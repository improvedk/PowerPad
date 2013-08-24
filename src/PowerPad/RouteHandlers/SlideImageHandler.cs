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
				new ErrorHandler(500, "Missing parameter: 'Number'").HandleRequest(context, sw);
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
				new ErrorHandler(500, "Invalid parameter value: 'Number'").HandleRequest(context, sw);
				return;
			}

			// Ensure slide image exists in cache
			string slideCachePath = Path.Combine(Settings.CacheDirectory, slideNumber + ".jpg");
			if (!File.Exists(slideCachePath))
			{
				new ErrorHandler(404, "Slide does not exist").HandleRequest(context, sw);
				return;
			}

			// Serve slide image to user
			var handler = new StaticFileHandler(slideCachePath);
			handler.HandleRequest(context, sw);
		}
	}
}