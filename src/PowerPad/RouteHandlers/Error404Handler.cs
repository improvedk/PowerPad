using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class Error404Handler : IRouteHandler
	{
		public void HandleRequest(HttpListenerContext context)
		{
			context.Response.StatusCode = 404;
			
			using (var sw = new StreamWriter(context.Response.OutputStream))
				sw.WriteLine("404");
		}
	}
}