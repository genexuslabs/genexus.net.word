using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;

namespace Genexus.Word.Shapes
{
    public abstract class CustomShape
    {
        public uint Id { get; set; }

        public CustomShapeProperties ShapeProperties = new CustomShapeProperties();
        public List<string> TextProperties = new List<string>();

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
        public double StrokeWidth { get; set; }
        public string StrokeColor { get; set; }
        public string FillColor { get; set; }
        public string InnerText { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.Center;
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Middle;

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
                            case "strokewidth":
                                cProps.StrokeWidth = Double.Parse(itemValue);
                                break;
                            case "color":
                                cProps.StrokeColor = itemValue;
                                break;
                            case "fillcolor":
                                cProps.FillColor = itemValue;
                                break;
                            case "horizontalalignment":
                                cProps.HorizontalAlignment = (HorizontalAlignment)Enum.Parse(typeof(HorizontalAlignment), itemValue, true);
                                break;
                            case "verticalalignment":
                                cProps.VerticalAlignment = (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), itemValue, true);                                
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

    public enum HorizontalAlignment
    {
        //
        // Summary:
        //     Align Left.
        //     When the item is serialized out as xml, its value is "left".
        Left = 0,
        //
        // Summary:
        //     start.
        //     When the item is serialized out as xml, its value is "start".
        //     This item is only available in Office2010.
        Start = 1,
        //
        // Summary:
        //     Align Center.
        //     When the item is serialized out as xml, its value is "center".
        Center = 2,
        //
        // Summary:
        //     Align Right.
        //     When the item is serialized out as xml, its value is "right".
        Right = 3,
        //
        // Summary:
        //     end.
        //     When the item is serialized out as xml, its value is "end".
        //     This item is only available in Office2010.
        End = 4,
        //
        // Summary:
        //     Justified.
        //     When the item is serialized out as xml, its value is "both".
        Both = 5,
        //
        // Summary:
        //     Medium Kashida Length.
        //     When the item is serialized out as xml, its value is "mediumKashida".
        MediumKashida = 6,
        //
        // Summary:
        //     Distribute All Characters Equally.
        //     When the item is serialized out as xml, its value is "distribute".
        Distribute = 7,
        //
        // Summary:
        //     Align to List Tab.
        //     When the item is serialized out as xml, its value is "numTab".
        NumTab = 8,
        //
        // Summary:
        //     Widest Kashida Length.
        //     When the item is serialized out as xml, its value is "highKashida".
        HighKashida = 9,
        //
        // Summary:
        //     Low Kashida Length.
        //     When the item is serialized out as xml, its value is "lowKashida".
        LowKashida = 10,
        //
        // Summary:
        //     Thai Language Justification.
        //     When the item is serialized out as xml, its value is "thaiDistribute".
        ThaiDistribute = 11
    }

    public enum VerticalAlignment
    {
        Top = 0,
        Middle = 1,
        Bottom = 2
    }

    
}
