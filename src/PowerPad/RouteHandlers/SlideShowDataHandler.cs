using System.IO;
using System.Net;
using System.Runtime.InteropServices;
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

			// As we can't lock on the program itself, the slideshow might end while we're reading from it
			try
			{
				// Return current slide show state
				var preso = Program.ActiveSlideShow.Presentation;
				var currentSlideNumber = Program.ActiveSlideShow.View.CurrentShowPosition;

				// Return state to client
				var state = new {
					numberOfSlides = preso.Slides.Count,
					currentSlideNumber,
					currentSlideNotes = Program.ActiveSlideShowCache.GetNote(currentSlideNumber)
				};

				var serializer = new JavaScriptSerializer();
				writer.Write(serializer.Serialize(state));
			}
			catch (COMException ex)
			{
				if (ex.Message.Contains("There is currently no slide show view for this presentation"))
				{
					new ErrorHandler(404, "No active slide show").HandleRequest(context, writer);
					return;
				}

				throw;
			}
		}
	}
}