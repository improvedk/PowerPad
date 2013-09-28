using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPad
{
	class Program
	{
		private const int maxConsolePresentationNameLength = 40;

		private static readonly Application ppt = new Application();
		private static readonly Stopwatch watch = new Stopwatch();
		
		public static SlideShowWindow ActiveSlideShow;
		public static Cache ActiveSlideShowCache;

		static void Main()
		{
			// Add global exception handler to log all exceptions
			AppDomain.CurrentDomain.UnhandledException += handleUnhandledException;

			// Start server
			using (var server = new PadServer(Settings.PortNumber))
			{
				server.Start();

				// Report which addresses server is listening on
				writeSuccess("Server now listening on:");
				foreach (var addr in server.ListeningAddresses)
					writeSuccess("\t" + addr);

				printHelp();
			
				// Wire up PowerPoint events
				ppt.PresentationOpen += ppt_PresentationOpen;
				ppt.SlideShowBegin += ppt_SlideShowBegin;
				ppt.SlideShowEnd += ppt_SlideShowEnd;
				ppt.SlideShowNextSlide += ppt_SlideShowNextSlide;

				// Either hook into running instance or start a new instance of PowerPoint up
				if (ppt.Visible == MsoTriState.msoTrue)
				{
					writeLine("Connected to running PowerPoint instance");

					// If there are any opened presentations, notify the user
					if (ppt.Presentations.Count == 0)
						writeLine("\tNo open presentations");
					else
					{
						foreach (Presentation preso in ppt.Presentations)
							writeLine("\t" + formatPresentationNameForConsole(preso));
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
					writeLine("Starting up new PowerPoint instance");
					ppt.Activate();
				}

				// Wait for the user to close by typing "quit<Enter>"
				var quitting = false;
				while (!quitting)
				{
					var cmd = Console.ReadLine();

					switch (cmd)
					{
						case "quit":
							quitting = true;
							continue;

						case "cache":
							cacheSlides();
							break;

						default:
							writeWarning("Unknown command: " + cmd);
							break;
					}
				}
			}
		}

		static void printHelp()
		{
			writeLine("Available commands:");
			writeLine("\tcache -- Caches the currently active slideshow");
			writeLine("\tquit -- Ends the PowerPad process");
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
			
			writeLine("\tCurrent slide: " + win.View.CurrentShowPosition);
		}

		/// <summary>
		/// Fires when a slideshow ends
		/// </summary>
		static void ppt_SlideShowEnd(Presentation pres)
		{
			ActiveSlideShow = null;
			writeWarning("Ending slide show");
		}

		/// <summary>
		/// Fires when a slideshow begins
		/// </summary>
		static void ppt_SlideShowBegin(SlideShowWindow win)
		{
			if (ActiveSlideShow != null)
			{
				writeWarning("Ignoring new slide show as another is already active");
				return;
			}

			writeSuccess("Beginning slide show");
			writeSuccess("\t" + formatPresentationNameForConsole(win.Presentation));

			// Start the timer & store presentation references
			ActiveSlideShow = win;
			ActiveSlideShowCache = new Cache(computeHashForPresentation(win.Presentation));
			watch.Reset();
			watch.Start();

			// Report whether slideshow cache is already primed
			if (ActiveSlideShowCache.AreAllSlidesCached(ActiveSlideShow.Presentation.Slides.Count))
				writeSuccess("\tAll slides cached, ready to go!");
			else
				writeWarning("\tPresentation needs to be cached!");
		}

		static string computeHashForPresentation(Presentation preso)
		{
			var presoFile = new FileInfo(Path.Combine(preso.Path, preso.Name));
			var sha1 = SHA1.Create();
			var presoHashBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(presoFile.FullName + presoFile.LastWriteTimeUtc));
			
			return BitConverter.ToString(presoHashBytes).Replace("-", "");
		}

		static void cacheSlides()
		{
			if (ActiveSlideShow == null || ActiveSlideShow.Presentation == null)
			{
				writeWarning("Can't cache slides as there is no active slideshow");
				return;
			}

			writeLine("Caching slides...");
			writeLine("\t0%");

			// Calculate hash to store cache, based on the presentation and it's last modification time
			var preso = ActiveSlideShow.Presentation; 
			var presoHash = computeHashForPresentation(preso);
			var cache = ActiveSlideShowCache = new Cache(presoHash);

			// Create cache directory
			cache.EnsureDirectoryExists();

			// Loop slides, cache & report progress
			int totalSlides = preso.Slides.Count;
			int previousProgress = 0;
			for (int i = 1; i <= totalSlides; i++)
			{
				Slide slide = preso.Slides[i];
				
				// If the user closes the slide show while we're caching, abort
				if (ActiveSlideShow == null)
				{
					writeWarning("\tAborting cache since slide show has ended");
					return;
				}
				
				// Cache image
				if (!cache.ImageIsCached(i))
					preso.Slides[i].Export(cache.GetImagePath(i), "jpg");

				// Cache notes
				if (!cache.NoteIsCached(i))
				{
					if (slide.HasNotesPage == MsoTriState.msoTrue)
					{
						// Attempt to find the shape for the slide notes frmae
						var notesShape = slide.NotesPage.Shapes.Cast<Shape>()
						                      .Where(s => s.Name == "Notes Placeholder 2")
						                      .Where(s => s.HasTextFrame == MsoTriState.msoTrue)
						                      .Where(s => s.TextFrame.HasText == MsoTriState.msoTrue)
						                      .SingleOrDefault();

						// If found, export the note contents
						if (notesShape != null)
							cache.SetNotes(i, notesShape.TextFrame.TextRange.Text);
					}
				}

				// Report progress
				int percentage = (int)Math.Round((double)i / totalSlides * 100, 0);
				if (percentage - previousProgress > 10 || percentage == 100)
				{
					writeLine("\t" + percentage + "%");
					previousProgress = percentage;
				}
			}

			writeSuccess("\tDone!");
		}

		/// <summary>
		/// Fires when user opens a PPTX file
		/// </summary>
		static void ppt_PresentationOpen(Presentation pres)
		{
			writeLine("Presentation opened");
			writeLine("\t" + formatPresentationNameForConsole(pres));
		}

		static void writeLine(object msg)
		{
			string message = Regex.Replace(msg.ToString(), "\t", "   ");

			Console.WriteLine(DateTime.Now.ToString("hh:mm:ss") + ":   " + message);
		}

		static void writeWarning(object msg)
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			writeLine(msg);
			Console.ResetColor();
		}

		static void writeSuccess(object msg)
		{
			Console.ForegroundColor = ConsoleColor.Green;
			writeLine(msg);
			Console.ResetColor();
		}
	}
}