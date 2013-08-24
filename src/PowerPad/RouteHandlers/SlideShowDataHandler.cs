using System.IO;
using System.Net;
using System.Web.Script.Serialization;

namespace PowerPad.RouteHandlers
{
	internal class SlideShowDataHandler : IRouteHandler
	{
		public void HandleRequest(HttpListenerContext context, StreamWriter writer)
		{
			context.Response.ContentType = "application/json";

			// If there's no active slide show, we can't return any data
			if (Program.ActiveSlideShow == null)
			{
				new ErrorHandler(404, "No active slide show").HandleRequest(context, writer);
				return;
			}

			// Return current slide show state
			var preso = Program.ActiveSlideShow.Presentation;
			var state = new {
				numberOfSlides = preso.Slides.Count,
				currentSlideNumber = Program.ActiveSlideShow.View.CurrentShowPosition
			};

			var serializer = new JavaScriptSerializer();
			writer.Write(serializer.Serialize(state));
		}
	}
}