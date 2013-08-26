using System;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace PowerPad
{
	internal static class Settings
	{
		/// <summary>
		/// This is the directory in which the presentation slide images are cached
		/// </summary>
		internal static readonly string CacheDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Cache");

		/// <summary>
		/// This is the port on which PowerPad will listen for incoming connections
		/// </summary>
		internal static readonly int PortNumber = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);

		/// <summary>
		/// This is the directory in which the frontend files are stored
		/// </summary>
		internal static string FrontendDirectory
		{
			get
			{
				// If we're in debug mode, we'll load the files directly from the source, assuming we're running from /bin/debug/*
				if (IsDebug)
					return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "../../Frontend");
				
				return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Frontend");
			}
		}

		/// <summary>
		/// When we're in debug mode, we'll enable certain developer-friendly settings
		/// </summary>
		internal static bool IsDebug
		{
			get { return System.Diagnostics.Debugger.IsAttached; }
		}
	}
}