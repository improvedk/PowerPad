using System.Collections.Generic;
using System.IO;

namespace PowerPad
{
	internal static class Cache
	{
		private static Dictionary<int, string> notes = new Dictionary<int, string>();

		internal static void Clear()
		{
			if (Directory.Exists(Settings.CacheDirectory))
				Directory.Delete(Settings.CacheDirectory, true);
		}

		internal static void EnsureDirectoryExists()
		{
			Directory.CreateDirectory(Settings.CacheDirectory);
		}

		internal static bool ImageIsCached(int slideNumber)
		{
			return File.Exists(GetImagePath(slideNumber));
		}

		internal static string GetImagePath(int slideNumber)
		{
			return Path.Combine(Settings.CacheDirectory, slideNumber + ".jpg");
		}

		internal static bool NotesAreCached(int slideNumber)
		{
			return notes.ContainsKey(slideNumber);
		}

		internal static string GetNotes(int slideNumber)
		{
			if (!NotesAreCached(slideNumber))
				return null;

			return notes[slideNumber];
		}

		internal static void SetNotes(int slideNumber, string noteValue)
		{
			notes[slideNumber] = noteValue;
		}
	}
}