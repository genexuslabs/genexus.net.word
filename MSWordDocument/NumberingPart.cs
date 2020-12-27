using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MSWordDocument
{
    internal class NumberingPart
    {
        // Generates content of numberingDefinitionsPart1.
        internal static void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14" } };
            numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            numbering1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            numbering1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            numbering1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            numbering1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            numbering1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            numbering1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            numbering1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            numbering1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            numbering1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            numbering1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            numbering1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            numbering1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            numbering1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid1 = new Nsid() { Val = "3D494EF0" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "3A6EFE6E" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "380A0001" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText1 = new LevelText() { Val = "·" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation2 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties1.Append(runFonts2);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "380A0003", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText2 = new LevelText() { Val = "o" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties2.Append(runFonts3);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "380A0005", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText3 = new LevelText() { Val = "§" };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Start = "2160", Hanging = "360" };

            previousParagraphProperties3.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties3.Append(runFonts4);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "380A0001", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText4 = new LevelText() { Val = "·" };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties4.Append(runFonts5);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level() { LevelIndex = 4, TemplateCode = "380A0003", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText5 = new LevelText() { Val = "o" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties5.Append(runFonts6);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level() { LevelIndex = 5, TemplateCode = "380A0005", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText6 = new LevelText() { Val = "§" };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation() { Start = "4320", Hanging = "360" };

            previousParagraphProperties6.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties6.Append(runFonts7);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level() { LevelIndex = 6, TemplateCode = "380A0001", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText7 = new LevelText() { Val = "·" };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties7.Append(runFonts8);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level() { LevelIndex = 7, TemplateCode = "380A0003", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText8 = new LevelText() { Val = "o" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties8.Append(runFonts9);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level() { LevelIndex = 8, TemplateCode = "380A0005", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText9 = new LevelText() { Val = "§" };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation() { Start = "6480", Hanging = "360" };

            previousParagraphProperties9.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties9.Append(runFonts10);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);
            level9.Append(numberingSymbolRunProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            abstractNum2.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid2 = new Nsid() { Val = "47DE622F" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "7C9283E6" };

            Level level10 = new Level() { LevelIndex = 0, TemplateCode = "380A000F" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation11 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties10.Append(indentation11);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties10.Append(runFonts11);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties10);

            Level level11 = new Level() { LevelIndex = 1, TemplateCode = "380A000F" };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText11 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation12 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties11.Append(indentation12);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties11.Append(runFonts12);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties11);

            Level level12 = new Level() { LevelIndex = 2, TemplateCode = "380A0005" };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText12 = new LevelText() { Val = "§" };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation13 = new Indentation() { Start = "2160", Hanging = "360" };

            previousParagraphProperties12.Append(indentation13);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties12.Append(runFonts13);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);
            level12.Append(numberingSymbolRunProperties12);

            Level level13 = new Level() { LevelIndex = 3, TemplateCode = "380A0001" };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText13 = new LevelText() { Val = "·" };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();
            Indentation indentation14 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties13.Append(indentation14);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties13.Append(runFonts14);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);
            level13.Append(numberingSymbolRunProperties13);

            Level level14 = new Level() { LevelIndex = 4, TemplateCode = "380A0003", Tentative = true };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText14 = new LevelText() { Val = "o" };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();
            Indentation indentation15 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties14.Append(indentation15);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties14.Append(runFonts15);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);
            level14.Append(numberingSymbolRunProperties14);

            Level level15 = new Level() { LevelIndex = 5, TemplateCode = "380A0005", Tentative = true };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText15 = new LevelText() { Val = "§" };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();
            Indentation indentation16 = new Indentation() { Start = "4320", Hanging = "360" };

            previousParagraphProperties15.Append(indentation16);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties15.Append(runFonts16);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);
            level15.Append(numberingSymbolRunProperties15);

            Level level16 = new Level() { LevelIndex = 6, TemplateCode = "380A0001", Tentative = true };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText16 = new LevelText() { Val = "·" };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();
            Indentation indentation17 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties16.Append(indentation17);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

            numberingSymbolRunProperties16.Append(runFonts17);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);
            level16.Append(numberingSymbolRunProperties16);

            Level level17 = new Level() { LevelIndex = 7, TemplateCode = "380A0003", Tentative = true };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText17 = new LevelText() { Val = "o" };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();
            Indentation indentation18 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties17.Append(indentation18);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

            numberingSymbolRunProperties17.Append(runFonts18);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);
            level17.Append(numberingSymbolRunProperties17);

            Level level18 = new Level() { LevelIndex = 8, TemplateCode = "380A0005", Tentative = true };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText18 = new LevelText() { Val = "§" };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();
            Indentation indentation19 = new Indentation() { Start = "6480", Hanging = "360" };

            previousParagraphProperties18.Append(indentation19);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties18.Append(runFonts19);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);
            level18.Append(numberingSymbolRunProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            AbstractNum abstractNum3 = new AbstractNum() { AbstractNumberId = 2 };
            abstractNum3.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid3 = new Nsid() { Val = "4F9A041C" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "A1C80EC8" };

            Level level19 = new Level() { LevelIndex = 0, TemplateCode = "380A000F" };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText19 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();
            Indentation indentation20 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties19.Append(indentation20);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);

            Level level20 = new Level() { LevelIndex = 1, TemplateCode = "380A0019" };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText20 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();
            Indentation indentation21 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties20.Append(indentation21);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);

            Level level21 = new Level() { LevelIndex = 2, TemplateCode = "380A001B", Tentative = true };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText21 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();
            Indentation indentation22 = new Indentation() { Start = "2160", Hanging = "180" };

            previousParagraphProperties21.Append(indentation22);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);

            Level level22 = new Level() { LevelIndex = 3, TemplateCode = "380A000F", Tentative = true };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText22 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();
            Indentation indentation23 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties22.Append(indentation23);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);

            Level level23 = new Level() { LevelIndex = 4, TemplateCode = "380A0019", Tentative = true };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText23 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();
            Indentation indentation24 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties23.Append(indentation24);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);

            Level level24 = new Level() { LevelIndex = 5, TemplateCode = "380A001B", Tentative = true };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText24 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();
            Indentation indentation25 = new Indentation() { Start = "4320", Hanging = "180" };

            previousParagraphProperties24.Append(indentation25);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);

            Level level25 = new Level() { LevelIndex = 6, TemplateCode = "380A000F", Tentative = true };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText25 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();
            Indentation indentation26 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties25.Append(indentation26);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);

            Level level26 = new Level() { LevelIndex = 7, TemplateCode = "380A0019", Tentative = true };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText26 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();
            Indentation indentation27 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties26.Append(indentation27);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);

            Level level27 = new Level() { LevelIndex = 8, TemplateCode = "380A001B", Tentative = true };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText27 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();
            Indentation indentation28 = new Indentation() { Start = "6480", Hanging = "180" };

            previousParagraphProperties27.Append(indentation28);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);

            abstractNum3.Append(nsid3);
            abstractNum3.Append(multiLevelType3);
            abstractNum3.Append(templateCode3);
            abstractNum3.Append(level19);
            abstractNum3.Append(level20);
            abstractNum3.Append(level21);
            abstractNum3.Append(level22);
            abstractNum3.Append(level23);
            abstractNum3.Append(level24);
            abstractNum3.Append(level25);
            abstractNum3.Append(level26);
            abstractNum3.Append(level27);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 2 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 1 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 0 };

            numberingInstance3.Append(abstractNumId3);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(abstractNum3);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);
            numbering1.Append(numberingInstance3);

            numberingDefinitionsPart1.Numbering = numbering1;
        }
    }
}
