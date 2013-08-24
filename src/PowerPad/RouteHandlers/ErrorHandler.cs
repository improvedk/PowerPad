using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class ErrorHandler : IRouteHandler
	{
		private readonly int errorCode;
		private readonly string message;

		internal ErrorHandler(int errorCode)
		{
			this.errorCode = errorCode;
			message = "File does not exist";
		}

		internal ErrorHandler(int errorCode, string message)
		{
			this.errorCode = errorCode;
			this.message = message;
		}

		public void HandleRequest(HttpListenerContext context, StreamWriter sw)
		{
			context.Response.StatusCode = errorCode;
			sw.WriteLine("Error: " + message);
		}
	}
}