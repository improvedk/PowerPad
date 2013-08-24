using System;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace PowerPad
{
	internal static class Settings
	{
		internal static readonly string CacheDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Cache");
		internal static readonly int PortNumber = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
		internal static readonly string FrontendDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Frontend");
	}
}