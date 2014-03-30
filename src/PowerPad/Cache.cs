using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPad
{
	internal class Cache
	{
		private readonly string cacheDirectory;
		private readonly string manifestPath;
		private readonly Dictionary<int, string> manifest = new Dictionary<int, string>();
		private readonly Presentation preso;
		private readonly SHA1 sha1 = SHA1.Create();

		internal Cache(Presentation preso)
		{
			this.preso = preso;

			// Calculate presentation hash based on full path
			var presoHashBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(preso.FullName));
			string hash = BitConverter.ToString(presoHashBytes).Replace("-", "");

			cacheDirectory = Path.Combine(Settings.CacheDirectory, hash);
			manifestPath = Path.Combine(cacheDirectory, "manifest.csv");
			ensureDirectoryExists();

			// Read existing manifest
			if (File.Exists(manifestPath))
			{
				foreach (var line in File.ReadAllLines(manifestPath))
				{
					var parts = line.Split(';');

					if (File.Exists(Path.Combine(cacheDirectory, parts[0] + ".jpg")))
						manifest.Add(Convert.ToInt32(parts[0]), parts[1]);
				}
			}

			cacheSlides();
		}

		private string sanitizeNotes(string notes)
		{
			if (notes == null)
				return notes;

			notes = Regex.Replace(notes, "\x0B", Environment.NewLine);

			return notes;
		}

		private void cacheSlides()
		{
			Log.Line("\tCaching slides...");
			Log.Line("\t\t0%");

			// Loop slides, cache & report progress
			int totalSlides = preso.Slides.Count;
			int previousProgress = 0;
			for (int i = 1; i <= totalSlides; i++)
			{
				// If the user closes the slide show while we're caching, abort
				if (PowerPad.ActiveSlideShow == null)
				{
					Log.Warning("\t\tAborting cache since slide show has ended");
					return;
				}

				// Cache slide if it isn't already cached
				var hash = computeSlideHash(preso.Slides[i]);

				if (!slideIsCached(i, hash))
				{
					// Export image
					preso.Slides[i].Export(GetImagePath(i), "jpg");

					// Export notes
					if (preso.Slides[i].HasNotesPage == MsoTriState.msoTrue)
					{
						// Attempt to find the shape for the slide notes frame
						var notesShape = preso.Slides[i].NotesPage.Shapes.Cast<Shape>()
														.Where(s => s.Name == "Notes Placeholder 2")
														.Where(s => s.HasTextFrame == MsoTriState.msoTrue)
														.Where(s => s.TextFrame.HasText == MsoTriState.msoTrue)
														.SingleOrDefault();

						// If found, export the note contents
						if (notesShape != null)
							SetNotes(i, sanitizeNotes(notesShape.TextFrame.TextRange.Text));
					}
					else
						SetNotes(i, null);

					// Add cached slide to manifest
					if (manifest.ContainsKey(i))
						manifest[i] = hash;
					else
						manifest.Add(i, hash);
				}

				// Report progress
				int percentage = (int)Math.Round((double)i / totalSlides * 100, 0);
				if (percentage - previousProgress > 10 || percentage == 100)
				{
					Log.Line("\t\t" + percentage + "%");
					previousProgress = percentage;
				}
			}

			// Write manifest to disk
			var sb = new StringBuilder();
			foreach (var key in manifest.Keys)
				sb.AppendLine(key + ";" + manifest[key]);
			File.WriteAllText(manifestPath, sb.ToString());

			Log.Success("\t\tDone!");
		}

		private string computeSlideHash(Slide slide)
		{
			var sb = new StringBuilder();

			sb.AppendLine(slide.SlideNumber.ToString());
			
			if (slide.HasNotesPage == MsoTriState.msoTrue)
			{
				foreach (Shape shape in slide.NotesPage.Shapes)
				{
					if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
						sb.AppendLine(shape.Left + ":" + shape.Top + ":" + shape.Width + ":" + shape.Height + ":" + shape.TextFrame2.TextRange.Text);
				}
			}

			foreach (Shape shape in slide.Shapes)
			{
				if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
					sb.AppendLine(shape.Left + ":" + shape.Top + ":" + shape.Width + ":" + shape.Height + ":" + shape.TextFrame2.TextRange.Text);
			}

			var slideHashBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
			return BitConverter.ToString(slideHashBytes).Replace("-", "");
		}

		private void ensureDirectoryExists()
		{
			Directory.CreateDirectory(cacheDirectory);
		}
		
		internal bool SlideIsCached(int slideNumber)
		{
			return slideIsCached(slideNumber, null);
		}

		private bool slideIsCached(int slideNumber, string hash)
		{
			if (hash != null)
				return manifest.ContainsKey(slideNumber) && manifest[slideNumber] == hash;
			else
				return manifest.ContainsKey(slideNumber);
		}

		internal string GetImagePath(int slideNumber)
		{
			return Path.Combine(cacheDirectory, slideNumber + ".jpg");
		}

		internal string GetNote(int slideNumber)
		{
			string path = GetNotePath(slideNumber);

			if (File.Exists(path))
				return File.ReadAllText(path);

			return null;
		}

		internal string GetNotePath(int slideNumber)
		{
			return Path.Combine(cacheDirectory, slideNumber + ".txt");
		}

		internal void SetNotes(int slideNumber, string noteValue)
		{
			string notePath = GetNotePath(slideNumber);

			if (noteValue == null)
			{
				if (File.Exists(notePath))
					File.Delete(notePath);
			}
			else
				File.WriteAllText(notePath, noteValue.Replace("\r", Environment.NewLine));
		}

		internal bool AreAllSlidesCached(int slideCount)
		{
			return manifest.Count == slideCount;
		}
	}
}