using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class Error404Handler : IRouteHandler
	{
		public void HandleRequest(HttpListenerContext context, StreamWriter sw)
		{
			context.Response.StatusCode = 404;
			sw.WriteLine("404");
		}
	}
}