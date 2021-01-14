using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using MSWordDocument;
using Genexus.Word.Helpers;

namespace Genexus.Word.Shapes
{
    public class RectangleShape : CustomShape
    {
        private static string DEFAULT_STROKE_HEX_COLOR = "000000";
        private static string DEFAULT_FILL_HEX_COLOR = "FFFFFF";
        private static int MIN_STROKE_WIDTH = 12700;
        #region 
        public RectangleShape()
        {
            RequiredImports = new Dictionary<string, string>()
            {
                { "wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"},
                { "mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"},
                { "o", "urn:schemas-microsoft-com:office:office"},
                { "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
                { "v", "urn:schemas-microsoft-com:vml"},
                { "m", "http://schemas.openxmlformats.org/officeDocument/2006/math"},
                { "wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"},
                { "wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"},
                { "w10", "urn:schemas-microsoft-com:office:word"},
                { "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
                { "w14", "http://schemas.microsoft.com/office/word/2010/wordml"},
                { "w15", "http://schemas.microsoft.com/office/word/2012/wordml"},
                { "wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"},
                { "wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"},
                { "wne", "http://schemas.microsoft.com/office/word/2006/wordml"},
                { "wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"}
            };
        }

        #endregion

        public override OpenXmlCompositeElement Build()
        {
            string sElementId = Guid.NewGuid().ToString();
            uint docPropId = Id;

            AlternateContent altContent = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor = new Wp.Anchor()
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)114300U,
                DistanceFromRight = (UInt32Value)114300U,
                SimplePos = false,
                RelativeHeight = (UInt32Value)251645952U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                SimplePosition = new Wp.SimplePosition() { X = 0L, Y = 0L },
                HorizontalPosition = new Wp.HorizontalPosition()
                {
                    RelativeFrom = Wp.HorizontalRelativePositionValues.Column,
                    PositionOffset = new Wp.PositionOffset()
                    {
                        Text = MathOpenXml.CentimetersToEMU(Properties.PositionLeft).ToString()
                    }

                },
                VerticalPosition = new Wp.VerticalPosition()
                {
                    RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph,
                    PositionOffset = new Wp.PositionOffset()
                    {
                        Text = MathOpenXml.CentimetersToEMU(Properties.PositionTop).ToString()
                    }
                },
            };
            Wp.Extent extent1 = new Wp.Extent()
            {
                Cx = MathOpenXml.CentimetersToEMU(Properties.Width),
                Cy = MathOpenXml.CentimetersToEMU(Properties.Height)
            };

            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent()
            {
                LeftEdge = 15240L,
                TopEdge = 15875L,
                RightEdge = 12065L,
                BottomEdge = 10795L
            };

            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties()
            {
                Id = docPropId,
                Name = "Text Box " + docPropId
            };

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties(graphicFrameLocks1);


            A.Graphic graphic1 = new A.Graphic();

            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData()
            {
                Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties(new A.ShapeLocks()
            {
                NoChangeArrowheads = true
            })
            {
                TextBox = true
            };

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties()
            {
                BlackWhiteMode = A.BlackWhiteModeValues.Auto
            };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 382270L, Cy = 230505L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry()
            {
                Preset = A.ShapeTypeValues.Rectangle
            };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill(new A.RgbColorModelHex()
            {
                Val = Helper.ToRGBHexColor(Properties.FillColor, DEFAULT_FILL_HEX_COLOR)
            });
            
            A.Outline outline1 = new A.Outline() { 
                Width = Math.Max(MIN_STROKE_WIDTH, (Int32Value)(MIN_STROKE_WIDTH * Properties.StrokeWidth))
            };

            A.SolidFill solidFill2 = new A.SolidFill(new A.RgbColorModelHex()
            {
                Val = Helper.ToRGBHexColor(Properties.StrokeColor, DEFAULT_STROKE_HEX_COLOR)
            });
            

            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(solidFill2);
            outline1.Append(new A.Miter());
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);

            Wps.TextBoxInfo2 txtInfo = new Wps.TextBoxInfo2();

            TextBoxContent txtContent = new TextBoxContent();

            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProps = new ParagraphProperties(new Justification()
            {
                Val = JustificationValues.Center
            });

            Run textRun = WordServerDocument.GetTextRun(Properties.InnerText, new List<String>()
            {

            });
            
            paragraph.Append(paragraphProps);
            paragraph.Append(textRun);

            txtContent.Append(paragraph);

            txtInfo.Append(txtContent);
                        
            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(txtInfo);
            wordprocessingShape1.Append(new Wps.TextBodyProperties(new A.NoAutoFit())
            {
                Rotation = 0,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                LeftInset = 74295,
                TopInset = 8890,
                RightInset = 74295,
                BottomInset = 8890,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = true,
                UpRight = false
            });

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor.Append(extent1);
            anchor.Append(effectExtent1);
            anchor.Append(wrapNone1);
            anchor.Append(docProperties1);
            anchor.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor.Append(graphic1);
            anchor.Append(relativeWidth1);
            anchor.Append(relativeHeight1);

            drawing1.Append(anchor);

            alternateContentChoice1.Append(drawing1);
            AlternateContentFallback altContentFallback = CreateAlternateContent();

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype()
            {

            };

            V.Stroke stroke1 = new V.Stroke()
            {
                JoinStyle = V.StrokeJoinStyleValues.Miter
            };

            V.Path path1 = new V.Path()
            {
                AllowGradientShape = true,
                ConnectionPointType = Ovml.ConnectValues.Rectangle
            };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape()
            {
                Id = "Text Box " + sElementId
            };

            V.TextBox textBox1 = new V.TextBox()
            {

            };

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph2 = new Paragraph()
            {
                ParagraphId = "58F841CD",
                TextId = "77777777"
            };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification()
            {
                Val = JustificationValues.Center
            };

            paragraphProperties2.Append(justification2);

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts2 = new RunFonts();

            runProperties3.Append(runFonts2);
            Text text2 = new Text();
            //text2.Text = Properties.InnerText;

            run3.Append(runProperties3);
            run3.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run3);

            textBoxContent2.Append(paragraph2);

            textBox1.Append(textBoxContent2);

            shape1.Append(textBox1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            altContentFallback.Append(picture1);

            altContent.Append(alternateContentChoice1);
            altContent.Append(altContentFallback);

            return altContent;
        }

        private static AlternateContentFallback CreateAlternateContent()
        {
            AlternateContentFallback altContentFallback = new AlternateContentFallback();
            altContentFallback.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            altContentFallback.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            altContentFallback.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            altContentFallback.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            altContentFallback.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            altContentFallback.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            altContentFallback.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            altContentFallback.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            altContentFallback.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            altContentFallback.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            altContentFallback.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            altContentFallback.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            altContentFallback.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            altContentFallback.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            altContentFallback.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            return altContentFallback;
        }
    }
}
