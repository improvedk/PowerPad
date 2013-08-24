using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal class NextSlideImageHandler : IRouteHandler
	{
		public void HandleRequest(HttpListenerContext context)
		{
			using (var sw = new StreamWriter(context.Response.OutputStream))
				sw.WriteLine("Next slide!");
		}
	}
}