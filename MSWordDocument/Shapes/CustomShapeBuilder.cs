using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Genexus.Word.Shapes
{
    public static class CustomShapeBuilder
    {
        public static OpenXmlCompositeElement BuildRectangle(MainDocumentPart docPart, uint id, string text, double width, double height)
        {
            RectangleShape rectangle = new RectangleShape()
            {
                Id = id,
                Height = height,
                Width = width,
                Text = text
            };

            rectangle.AddRequiredNamespaces(docPart);
            return rectangle.Build();

        }
    }
}
