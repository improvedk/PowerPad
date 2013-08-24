using System.IO;
using System.Net;

namespace PowerPad.RouteHandlers
{
	internal interface IRouteHandler
	{
		void HandleRequest(HttpListenerContext context, StreamWriter writer);
	}
}