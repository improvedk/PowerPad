using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class StaticFileHandler : IRouteHandler
	{
		private readonly byte[] source;
		private readonly string contentType;

		private Dictionary<string, string> knownContentTypes = new Dictionary<string, string> {
			{ ".js", "application/javascript" },
			{ ".htm", "text/html" },
			{ ".jpg", "image/jpeg" },
			{ ".png", "image/png" }
		};

		internal StaticFileHandler(string path)
		{
			if (path == null)
				throw new ArgumentException("Path can't be null");

			source = File.ReadAllBytes(path);

			string extension = Path.GetExtension(path);
			if (knownContentTypes.ContainsKey(extension))
				contentType = knownContentTypes[extension];
		}

		public void HandleRequest(HttpListenerContext context, StreamWriter sw)
		{
			if (contentType != null)
				context.Response.ContentType = contentType;

			context.Response.ContentLength64 = source.Length;
			context.Response.OutputStream.Write(source, 0, source.Length);
		}
	}
}