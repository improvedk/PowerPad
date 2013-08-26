using PowerPad.RouteHandlers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading;

namespace PowerPad
{
	internal class PadServer : IDisposable
	{
		private readonly int portNumber;
		private Thread listenerThread;
		private HttpListener listener;
		private readonly Dictionary<string, IRouteHandler> routeHandlers = new Dictionary<string, IRouteHandler>();
		private readonly Dictionary<int, IRouteHandler> errorHandlers = new Dictionary<int, IRouteHandler>();
		
		internal PadServer(int portNumber)
		{
			this.portNumber = portNumber;
		}

		internal IEnumerable<string> ListeningAddresses
		{
			get
			{
				foreach (var ni in NetworkInterface.GetAllNetworkInterfaces())
					foreach (var ua in ni.GetIPProperties().UnicastAddresses.Where(ua => ua.Address.AddressFamily == AddressFamily.InterNetwork))
						yield return "http://" + ua.Address + ":" + portNumber + "/";
			}
		}

		internal void Start()
		{
			if (listenerThread != null)
				throw new InvalidOperationException("Can't start already started PadServer");

			// Setup routes
			routeHandlers.Add("/", new StaticFileHandler(Path.Combine(Settings.FrontendDirectory, "index.htm")));
			routeHandlers.Add("/jquery-2.0.3.min.js/", new StaticFileHandler(Path.Combine(Settings.FrontendDirectory, "jquery-2.0.3.min.js")));
			routeHandlers.Add("/slideimage/", new SlideImageHandler());
			routeHandlers.Add("/slideshowdata/", new SlideShowDataHandler());

			// Setup error handlers
			errorHandlers.Add(404, new ErrorHandler(404));

			// Setup listener and start listening
			listener = new HttpListener();
			listener.Prefixes.Add("http://+:" + portNumber + "/");

			listenerThread = new Thread(listenForRequests);
			listenerThread.Start();

			listener.Start();
		}

		/// <summary>
		/// Continually listens for incoming requests and calls the request handler
		/// </summary>
		private void listenForRequests()
		{
			while (listener.IsListening)
			{
				var context = listener.BeginGetContext(handleRequest, listener);
				context.AsyncWaitHandle.WaitOne();
			}
		}

		/// <summary>
		/// Handles requests
		/// </summary>
		private void handleRequest(IAsyncResult ar)
		{
			var listener = ar.AsyncState as HttpListener;
			var context = listener.EndGetContext(ar);

			// Make sure path always ends with a /
			string path = context.Request.Url.LocalPath;
			if (!path.EndsWith("/"))
				path += "/";

			// By default, we don't want to cache anything
			context.Response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate");
			context.Response.AddHeader("Pragma", "no-cache");
			context.Response.AddHeader("Expires", "0");

			// Locate route/error handler, if none found, return 404
			using (var sw = new StreamWriter(context.Response.OutputStream))
			{
				if (routeHandlers.ContainsKey(path))
				{
					context.Response.StatusCode = 200;
					routeHandlers[path].HandleRequest(context, sw);
				}
				else
				{
					context.Response.StatusCode = 404;
					errorHandlers[404].HandleRequest(context, sw);
				}
			}

			context.Response.Close();
		}

		internal void Stop()
		{
			// First stop the listener, then abort the request handler thread
			listener.Stop();
			listenerThread.Abort();
		}

		public void Dispose()
		{
			Stop();
		}
	}
}