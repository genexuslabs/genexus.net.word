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

namespace Genexus.Word.Shapes
{
    public class RectangleShape : CustomShape
    {
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
                HorizontalPosition = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column }
            };

            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset() { Text = "-493395" };

            anchor.HorizontalPosition.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "156210";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 382270L, Cy = 230505L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 15240L, TopEdge = 15875L, RightEdge = 12065L, BottomEdge = 10795L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = docPropId, Name = "Text Box " + docPropId };


            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 382270L, Cy = 230505L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            A.Outline outline1 = new A.Outline() { Width = 19050 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(solidFill2);
            outline1.Append(miter1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties1.Append(justification1);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            runProperties2.Append(runFonts1);
            Text text1 = new Text();
            text1.Text = Text;

            run2.Append(runProperties2);
            run2.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run2);

            textBoxContent1.Append(paragraph1);

            textBoxInfo21.Append(textBoxContent1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 74295, TopInset = 8890, RightInset = 74295, BottomInset = 8890, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

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
            
            anchor.Append(verticalPosition1);
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

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape() { Id = "Text Box " + sElementId, Style = "position:absolute;margin-left:-38.85pt;margin-top:12.3pt;width:30.1pt;height:18.15pt;z-index:251645952;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", StrokeWeight = "1.5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

            V.TextBox textBox1 = new V.TextBox() { Inset = "5.85pt,.7pt,5.85pt,.7pt" };

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00FC6179", RsidParagraphAddition = "00F60DF2", RsidParagraphProperties = "00596D2F", RsidRunAdditionDefault = "00F60DF2", ParagraphId = "58F841CD", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties2.Append(justification2);

            Run run3 = new Run() { RsidRunProperties = "00FC6179" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            runProperties3.Append(runFonts2);
            Text text2 = new Text();
            text2.Text = Text;

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
    }
}
