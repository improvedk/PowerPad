using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class StaticFileHandler : IRouteHandler
	{
		private readonly Func<byte[]> getSource;
		private readonly byte[] cachedSource;
		private readonly string contentType;

		private readonly Dictionary<string, string> knownContentTypes = new Dictionary<string, string> {
			{ ".js", "application/javascript" },
			{ ".htm", "text/html" },
			{ ".jpg", "image/jpeg" },
			{ ".png", "image/png" }
		};

		internal StaticFileHandler(string path)
		{
			if (path == null)
				throw new ArgumentException("Path can't be null");

			// If we're in debug mode, we'll always read the file from disk rather than caching it
			if (Settings.IsDebug)
				getSource = () => File.ReadAllBytes(path);
			else
			{
				cachedSource = File.ReadAllBytes(path);
				getSource = () => cachedSource;
			}

			// Set the content type depending on the file extension
			string extension = Path.GetExtension(path);

			if (extension == null)
				throw new ArgumentException("Unknown filetype: " + Path.GetFileName(path));

			if (knownContentTypes.ContainsKey(extension))
				contentType = knownContentTypes[extension];
		}

		public void HandleRequest(HttpListenerContext context, StreamWriter sw)
		{
			if (contentType != null)
				context.Response.ContentType = contentType;

			byte[] source = getSource();
			context.Response.ContentLength64 = source.Length;
			context.Response.OutputStream.Write(source, 0, source.Length);
		}
	}
}