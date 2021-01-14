using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace Genexus.Word.Shapes
{
    public static class CustomShapeBuilder
    {
        public static OpenXmlCompositeElement BuildRectangle(MainDocumentPart docPart, uint id, string text, double width, double height, double left, double top, List<string> props, List<string> textProperties)
        {
            RectangleShape rectangle = new RectangleShape()
            {
                Id = id,
                ShapeProperties = CustomShapeProperties.Create(props),
                TextProperties = textProperties
            };
            
            rectangle.ShapeProperties.InnerText = text;
            rectangle.ShapeProperties.Height = height;
            rectangle.ShapeProperties.Width = width;
            rectangle.ShapeProperties.PositionTop = top;
            rectangle.ShapeProperties.PositionLeft = left;

            rectangle.AddRequiredNamespaces(docPart);
            return rectangle.Build();

        }
    }
}
