using System;
using System.Text.RegularExpressions;

namespace PowerPad
{
	internal static class Log
	{
		internal static void Line(object msg)
		{
			string message = Regex.Replace(msg.ToString(), "\t", "   ");

			Console.WriteLine(DateTime.Now.ToString("hh:mm:ss") + ":   " + message);
		}

		internal static void Warning(object msg)
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			Line(msg);
			Console.ResetColor();
		}

		internal static void Success(object msg)
		{
			Console.ForegroundColor = ConsoleColor.Green;
			Line(msg);
			Console.ResetColor();
		}
	}
}