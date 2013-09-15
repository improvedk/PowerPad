using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPad
{
	class Program
	{
		private static readonly Application ppt = new Application();
		private static readonly Stopwatch watch = new Stopwatch();
		
		public static SlideShowWindow ActiveSlideShow;

		static void Main()
		{
			// Add global exception handler to log all exceptions
			AppDomain.CurrentDomain.UnhandledException += handleUnhandledException;

			// Clear cache
			Cache.Clear();

			// Start server
			using (var server = new PadServer(Settings.PortNumber))
			{
				server.Start();

				// Report which addresses server is listening on
				writeLine("Server now listening on:");
				foreach (var addr in server.ListeningAddresses)
					writeLine("\t" + addr);
			
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
							writeLine("\t" + preso.Name + " (" + preso.Slides.Count + " slides)");
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
				while (Console.ReadLine() != "quit")
				{ }
			}
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
			
			writeLine("Current slide: " + win.View.CurrentShowPosition);
		}

		/// <summary>
		/// Fires when a slideshow ends
		/// </summary>
		static void ppt_SlideShowEnd(Presentation pres)
		{
			ActiveSlideShow = null;
			writeLine("Ending slide show");
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

			writeLine("Beginning slide show");

			// Start the timer
			ActiveSlideShow = win;
			watch.Reset();
			watch.Start();
			
			// Save all slides as JPGs
			cacheSlides(win.Presentation);
		}

		static void cacheSlides(Presentation preso)
		{
			writeLine("Caching slides...");
			writeLine("0%");

			// Create cache directory
			Cache.EnsureDirectoryExists();

			// Loop slides, cache & report progress
			int totalSlides = preso.Slides.Count;
			int previousProgress = 0;
			for (int i = 1; i <= totalSlides; i++)
			{
				Slide slide = preso.Slides[i];

				// If the user closes the slide show while we're caching, abort
				if (ActiveSlideShow == null)
				{
					writeLine("Aborting cache since slide show has ended");
					return;
				}

				// Export slide image if it hasn't already been cached
				if (!Cache.ImageIsCached(i))
					preso.Slides[i].Export(Cache.GetImagePath(i), "jpg");

				// Export slide notes if they haven't already been cached
				if (!Cache.NotesAreCached(i))
				{
					string notes = "";

					// We can only export notes if slide has a notes page
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
							notes = notesShape.TextFrame.TextRange.Text;
					}

					Cache.SetNotes(i, notes);
				}

				// Report progress
				int percentage = (int)Math.Round((double)i / totalSlides * 100, 0);
				if (percentage - previousProgress > 10 || percentage == 100)
				{
					writeLine(percentage + "%");
					previousProgress = percentage;
				}
			}
		}

		/// <summary>
		/// Fires when user opens a PPTX file
		/// </summary>
		static void ppt_PresentationOpen(Presentation pres)
		{
			writeLine("Presentation opened: " + pres.Name + " (" + pres.Slides.Count + " slides)");
		}

		static void writeLine(object msg)
		{
			Console.WriteLine(DateTime.Now.ToString("hh:mm:ss") + ":\t" + msg);
		}

		/// <summary>
		/// Writes a formatted warning message to the console
		/// </summary>
		static void writeWarning(object msg)
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			writeLine(msg);
			Console.ResetColor();
		}
	}
}