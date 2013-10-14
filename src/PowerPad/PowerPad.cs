using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;

namespace PowerPad
{
	class PowerPad
	{
		private const int maxConsolePresentationNameLength = 40;

		private static readonly Application ppt = new Application();
		private static readonly Stopwatch watch = new Stopwatch();
		
		public static SlideShowWindow ActiveSlideShow;
		public static Cache ActiveSlideShowCache;

		static void Main()
		{
			Console.Title = "PowerPad";

			// Add global exception handler to log all exceptions
			AppDomain.CurrentDomain.UnhandledException += handleUnhandledException;

			// Start server
			using (var server = new PadServer(Settings.PortNumber))
			{
				printHelp();

				if (server.Start())
				{
					// Report which addresses server is listening on
					Log.Success("Server now listening on:");
					foreach (var addr in server.ListeningAddresses)
						Log.Success("\t" + addr);

					// Wire up PowerPoint events
					ppt.PresentationOpen += ppt_PresentationOpen;
					ppt.SlideShowBegin += ppt_SlideShowBegin;
					ppt.SlideShowEnd += ppt_SlideShowEnd;
					ppt.SlideShowNextSlide += ppt_SlideShowNextSlide;

					// Either hook into running instance or start a new instance of PowerPoint up
					if (ppt.Visible == MsoTriState.msoTrue)
					{
						Log.Line("Connected to running PowerPoint instance");

						// If there are any opened presentations, notify the user
						if (ppt.Presentations.Count == 0)
							Log.Line("\tNo open presentations");
						else
						{
							foreach (Presentation preso in ppt.Presentations)
								Log.Line("\t" + formatPresentationNameForConsole(preso));
						}

						// Do we need to connect to a running slide show?
						if (ppt.SlideShowWindows.Count > 0)
						{
							ppt_SlideShowBegin(ppt.SlideShowWindows[1]);

							// Has slide show been closed while we were starting it up?
							if (ActiveSlideShow != null)
								ppt_SlideShowNextSlide(ppt.SlideShowWindows[1]);
						}
					}
					else
					{
						Log.Line("Starting up new PowerPoint instance");
						ppt.Activate();
					}
				}

				// Wait for user command
				var quitting = false;
				while (!quitting)
				{
					var cmd = Console.ReadLine();

					switch (cmd)
					{
						case "quit":
							quitting = true;
							continue;

						default:
							Log.Warning("Unknown command: " + cmd);
							break;
					}
				}
			}
		}

		static void printHelp()
		{
			Log.Line("Available commands:");
			Log.Line("\tquit -- Ends the PowerPad process");
		}

		static string formatPresentationNameForConsole(Presentation preso)
		{
			string name = preso.Name;

			if (preso.Name.Length > maxConsolePresentationNameLength)
				name = preso.Name.Substring(0, maxConsolePresentationNameLength) + "...";

			return name + " (" + preso.Slides.Count + ")";
		}

		/// <summary>
		/// Takes care of global exception handling by crudely logging exceptions to a simple text file
		/// </summary>
		static void handleUnhandledException(object sender, UnhandledExceptionEventArgs e)
		{
			// This will overwrite any previous exceptions, but it'll do for now
			File.WriteAllText("Exception.txt", e.ExceptionObject.ToString());
		}
		
		/// <summary>
		/// Fires when user changes slide during an active slideshow
		/// </summary>
		static void ppt_SlideShowNextSlide(SlideShowWindow win)
		{
			if (ActiveSlideShow != null && win.HWND != ActiveSlideShow.HWND)
				return;

			Log.Line("\tCurrent slide: " + win.View.CurrentShowPosition);
		}

		/// <summary>
		/// Fires when a slideshow ends
		/// </summary>
		static void ppt_SlideShowEnd(Presentation pres)
		{
			ActiveSlideShow = null;
			Log.Warning("Ending slide show");
		}

		/// <summary>
		/// Fires when a slideshow begins
		/// </summary>
		static void ppt_SlideShowBegin(SlideShowWindow win)
		{
			if (ActiveSlideShow != null)
			{
				Log.Warning("Ignoring new slide show as another is already active");
				return;
			}

			Log.Success("Beginning slide show");
			Log.Success("\t" + formatPresentationNameForConsole(win.Presentation));

			// Start the timer & store presentation references
			ActiveSlideShow = win;
			ActiveSlideShowCache = new Cache(win.Presentation);
			watch.Reset();
			watch.Start();
		}

		/// <summary>
		/// Fires when user opens a PPTX file
		/// </summary>
		static void ppt_PresentationOpen(Presentation pres)
		{
			Log.Line("Presentation opened");
				Log.Success("\t" + formatPresentationNameForConsole(pres));
		}
	}
}