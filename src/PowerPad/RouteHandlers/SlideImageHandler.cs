using System;
using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class SlideImageHandler : IRouteHandler
	{
		public void HandleRequest(HttpListenerContext context, StreamWriter sw)
		{
			// Validate parameters
			if (context.Request.QueryString["Number"] == null)
			{
				context.Response.StatusCode = 500;
				sw.WriteLine("Missing parameter 'Number'");
				return;
			}

			// Try to read slide number
			int slideNumber;
			try
			{
				slideNumber = Convert.ToInt32(context.Request.QueryString["Number"]);
			}
			catch (FormatException)
			{
				context.Response.StatusCode = 500;
				sw.WriteLine("Invalid parameter value: 'Number'");
				return;
			}

			// Ensure slide image exists in cache
			string slideCachePath = Path.Combine(Settings.CacheDirectory, slideNumber + ".jpg");
			if (!File.Exists(slideCachePath))
			{
				context.Response.StatusCode = 404;
				sw.WriteLine("Slide does not exist");
				return;
			}

			// Serve slide image to user
			var handler = new StaticFileHandler(slideCachePath);
			handler.HandleRequest(context, sw);
		}
	}
}