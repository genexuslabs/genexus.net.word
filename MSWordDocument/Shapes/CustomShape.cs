using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;

namespace Genexus.Word.Shapes
{
    public abstract class CustomShape
    {
        public uint Id { get; set; }

        public CustomShapeProperties Properties = new CustomShapeProperties();

        protected Dictionary<string, string> RequiredImports = new Dictionary<string, string>();

        public virtual OpenXmlCompositeElement Build()
        {
            throw new NotImplementedException();
        }

        public void AddRequiredNamespaces(MainDocumentPart documentPart)
        {
            foreach (var item in RequiredImports)
            {
                documentPart.Document.RemoveNamespaceDeclaration(item.Key);
                documentPart.Document.AddNamespaceDeclaration(item.Key, item.Value);
            }
        }
    }

    public class CustomShapeProperties
    {
        private static char VALUE_SEPARATOR = ':';

        public double Height { get; set; }
        public double Width { get; set; }
        public double PositionLeft { get; set; }
        public double PositionTop { get; set; }
        public double StrokeWeight { get; set; }
        public string InnerText { get; set; }

        public static CustomShapeProperties Create(List<string> Props)
        {            
            CustomShapeProperties cProps = new CustomShapeProperties();
            if (Props != null)
            {
                foreach (var item in Props)
                {
                    string[] itemSplit = item.Split(VALUE_SEPARATOR);
                    if (itemSplit.Length == 2)
                    {
                        string itemKey = itemSplit[0].ToLower();
                        string itemValue = itemSplit[1];
                        switch (itemKey)
                        {
                            case "strokeweight":
                                cProps.StrokeWeight = Double.Parse(itemValue);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            return cProps;
        }
    }
}
