using System;
using System.IO;
using System.Net;
using System.Threading;

namespace PowerPad
{
	internal class PadServer : IDisposable
	{
		private readonly int portNumber;
		private Thread listenerThread;
		private HttpListener listener;

		internal PadServer(int portNumber)
		{
			this.portNumber = portNumber;
		}

		internal void Start()
		{
			if (listenerThread != null)
				throw new InvalidOperationException("Can't start already started PadServer");

			listener = new HttpListener();
			listener.Prefixes.Add("http://+:" + portNumber + "/");

			listenerThread = new Thread(listenForRequests);
			listenerThread.Start();

			listener.Start();
		}

		private void listenForRequests()
		{
			while (listener.IsListening)
			{
				var context = listener.BeginGetContext(handleRequest, listener);
				context.AsyncWaitHandle.WaitOne();
			}
		}

		static void handleRequest(IAsyncResult ar)
		{
			var listener = ar.AsyncState as HttpListener;
			var context = listener.EndGetContext(ar);

			context.Response.StatusCode = 200;

			using (var sw = new StreamWriter(context.Response.OutputStream))
			{
				sw.WriteLine("Hello world!");
			}
		}

		internal void Stop()
		{
			listener.Stop();
			listenerThread.Abort();
		}

		public void Dispose()
		{
			Stop();
		}
	}
}