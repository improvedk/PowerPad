using System;
using System.IO;
using System.Linq;

namespace PowerPad
{
	internal class Cache
	{
		private readonly string cacheDirectory;

		internal Cache(string hash)
		{
			cacheDirectory = Path.Combine(Settings.CacheDirectory, hash);
		}

		internal void EnsureDirectoryExists()
		{
			Directory.CreateDirectory(cacheDirectory);
		}

		internal bool ImageIsCached(int slideNumber)
		{
			return File.Exists(GetImagePath(slideNumber));
		}

		internal string GetImagePath(int slideNumber)
		{
			return Path.Combine(cacheDirectory, slideNumber + ".jpg");
		}

		internal bool NoteIsCached(int slideNumber)
		{
			return File.Exists(GetNote(slideNumber));
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
			File.WriteAllText(GetNotePath(slideNumber), noteValue.Replace("\r", Environment.NewLine));
		}

		internal bool AreAllSlidesCached(int slideCount)
		{
			if (!Directory.Exists(cacheDirectory))
				return false;

			return Directory.GetFiles(cacheDirectory).Count(x => Path.GetExtension(x) == ".jpg") == slideCount;
		}
	}
}