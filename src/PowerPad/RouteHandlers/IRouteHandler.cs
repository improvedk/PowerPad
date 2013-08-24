using System.Net;

namespace PowerPad.RouteHandlers
{
	internal interface IRouteHandler
	{
		void HandleRequest(HttpListenerContext context);
	}
}