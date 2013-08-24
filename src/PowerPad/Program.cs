using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPad
{
	class Program
	{
		private static readonly Application ppt = new Application();
		private static SlideShowWindow activeSlideShow;
		private static readonly Stopwatch watch = new Stopwatch();
		private static readonly string cacheDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Cache");

		static void Main()
		{
			// Clear cache
			if (Directory.Exists(cacheDirectory))
				Directory.Delete(cacheDirectory, true);
			Directory.CreateDirectory(cacheDirectory);

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
					ppt_SlideShowNextSlide(ppt.SlideShowWindows[1]);
				}
			}
			else
			{
				writeLine("Starting up new PowerPoint instance");
				ppt.Activate();
			}

			// Start server
			int portNumber = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
			using (var server = new PadServer(portNumber))
			{
				server.Start();

				// Wait for the user to close by typing "quit<Enter>"
				while (Console.ReadLine() != "quit")
				{ }
			}
		}
		
		/// <summary>
		/// Fires when user changes slide during an active slideshow
		/// </summary>
		static void ppt_SlideShowNextSlide(SlideShowWindow win)
		{
			if (activeSlideShow != null && win.HWND != activeSlideShow.HWND)
				return;
			
			writeLine("Current slide: " + win.View.CurrentShowPosition);
		}

		/// <summary>
		/// Fires when a slideshow ends
		/// </summary>
		static void ppt_SlideShowEnd(Presentation pres)
		{
			activeSlideShow = null;
			writeLine("Ending slide show");
		}

		/// <summary>
		/// Fires when a slideshow begins
		/// </summary>
		static void ppt_SlideShowBegin(SlideShowWindow win)
		{
			if (activeSlideShow != null)
			{
				writeWarning("Ignoring new slide show as another is already active");
				return;
			}

			writeLine("Beginning slide show");

			// Start the timer
			activeSlideShow = win;
			watch.Reset();
			watch.Start();
			
			// Save all slides as JPGs
			cacheSlides(win.Presentation);
		}

		static void cacheSlides(Presentation preso)
		{
			writeLine("Caching slides...");
			writeLine("0%");

			// Loop slides, cache & report progress
			int totalSlides = preso.Slides.Count;
			int previousProgress = 0;
			for (int i = 1; i <= totalSlides; i++)
			{
				// Only export if slide hasn't already been cached
				if (!File.Exists(Path.Combine(cacheDirectory, i + ".jpg")))
				{
					// Export slide
					Slide slide = preso.Slides[i];
					slide.Export(Path.Combine(cacheDirectory, i + ".jpg"), "jpg");
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