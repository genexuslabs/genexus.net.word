using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace Genexus.Word.Shapes
{
    public static class CustomShapeBuilder
    {
        public static OpenXmlCompositeElement BuildRectangle(MainDocumentPart docPart, uint id, string text, double width, double height, double left, double top, List<string> props)
        {
            RectangleShape rectangle = new RectangleShape()
            {
                Id = id,
                Properties = CustomShapeProperties.Create(props)
            };
            
            rectangle.Properties.InnerText = text;
            rectangle.Properties.Height = height;
            rectangle.Properties.Width = width;
            rectangle.Properties.PositionTop = top;
            rectangle.Properties.PositionLeft = left;

            rectangle.AddRequiredNamespaces(docPart);
            return rectangle.Build();

        }
    }
}
