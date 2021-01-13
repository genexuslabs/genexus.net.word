using Genexus.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace MSWordUnitTesting
{
    [TestClass]
    public class WordDocumentUnits
    {
        private static string s_BasePath;

        [TestInitialize]
        public void Initialize()
		{
            s_BasePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
		}

        [TestMethod]
        public void OpenFromTemplate()
        {
            WordServerDocument doc = new WordServerDocument();
            Assert.AreEqual(doc.Open($"{s_BasePath}\\TemplateSample.docx", out _), OutputCode.OK);
            doc.Close();
        }



        [TestMethod]
        public void ReplaceFull()
        {
            using (WordServerDocument doc = new WordServerDocument())
            {
                File.Copy($"{s_BasePath}\\SampleFull.docx", $"{s_BasePath}\\SampleFullInstance.docx", true);

                Assert.AreEqual(doc.Open($"{s_BasePath}\\SampleFullInstance.docx", out _), OutputCode.OK);

                Assert.AreEqual(doc.ReplaceTextWithStyle("Imported:", "新潟県新潟市", true, new List<string>() { "color:blue", "italic", "fontsize:24", "fontfamily:MS PMincho" }), 4);
       //         Assert.AreEqual(doc.ReplaceTextWithStyle("REP02", "中央区米山", true, new List<string>() { "italic", "bold", "color:green", "fontfamily:MS PGothic" }), 1);
        //        Assert.AreEqual(doc.ReplaceTextWithStyle("REP03", "Remplazo con bold blue italic", true, new List<string>() { "italic", "bold", "color:blue" }), 1);


          //      doc.AddText("製品企画室", new List<string>() { "fontsize:110" });
            //    doc.AddText("TEL:", new List<string>() { "italic" });
             //   doc.AddText("FAX:", new List<string>() { "fontsize:54", "color:pink" });

                doc.Save();
            }

        }

        [TestMethod]
        public void ReplaceAndAddWithStyle()
        {
            using (WordServerDocument doc = new WordServerDocument())
            {
                File.Copy($"{s_BasePath}\\TemplateSample.docx", $"{s_BasePath}\\Sample.docx", true);
              
                Assert.AreEqual(doc.Open($"{s_BasePath}\\Sample.docx", out _), OutputCode.OK);
  
                Assert.AreEqual(doc.ReplaceTextWithStyle("REP01", "新潟県新潟市", true, new List<string>() { "color:blue", "italic" , "fontsize:24", "fontfamily:MS PMincho" }), 2);
                Assert.AreEqual(doc.ReplaceTextWithStyle("REP02", "中央区米山", true, new List<string>() {  "italic" , "bold" , "color:green", "fontfamily:MS PGothic"}), 1);
                Assert.AreEqual(doc.ReplaceTextWithStyle("REP03", "Remplazo con bold blue italic", true, new List<string>() { "italic", "bold" , "color:blue"}), 1);


                doc.AddText("製品企画室", new List<string>() { "fontsize:110" });
                doc.AddText("TEL:", new List<string>() { "italic" });
                doc.AddText("FAX:", new List<string>() { "fontsize:54", "color:pink" });

                doc.Save();
            }
            
        }

        [TestMethod]
        public void ReplaceAndAddWithImage()
        {
            using (WordServerDocument doc = new WordServerDocument())
            {
                File.Copy($"{s_BasePath}\\TemplateSample.docx", $"{s_BasePath}\\SampleImage.docx", true);

                Assert.AreEqual(doc.Open($"{s_BasePath}\\SampleImage.docx", out _), OutputCode.OK);


                Assert.AreEqual(doc.ReplaceTextWithStyle("REP01", "新潟県新潟市", true, new List<string>() { "color:blue", "italic", "fontsize:24", "fontfamily:MS PMincho" }), 2);
                Assert.AreEqual(doc.ReplaceTextWithStyle("REP02", "中央区米山", true, new List<string>() { "italic", "bold", "color:green", "fontfamily:MS PGothic" }), 1);
                Assert.AreEqual(doc.ReplaceTextWithImage("REP03", false, $"{s_BasePath}\\dos.png", 50, 50), 1);


                doc.AddText("製品企画室", new List<string>() { "fontsize:110" });
                doc.AddText("TEL:", new List<string>() { "italic" });
                doc.AddText("FAX:", new List<string>() { "fontsize:54", "color:pink" });

                doc.Save();
            }

        }




        [TestMethod]
        public void CreationTest()
        {
            WordServerDocument doc = new WordServerDocument();
            Assert.AreEqual(doc.Create($"{s_BasePath}\\test.docx", true, out _), OutputCode.OK);
            doc.Save();
            doc.Close();
        }

        [TestMethod]
        public void CreationTestWithoutOverwrite()
        {
            WordServerDocument doc = new WordServerDocument();
            Assert.AreEqual(doc.Create($"{s_BasePath}\\test.docx", false, out _), OutputCode.FILE_ALREADY_EXISTS);
            doc.Close();
        }

        [TestMethod]
        public void CreationParagraphsWithProperties()
        {
            WordServerDocument doc = new WordServerDocument();
            Assert.AreEqual(doc.Create($"{s_BasePath}\\testFormats.docx", true, out _), OutputCode.OK);


            doc.AddText("Bold", new List<string>() { "bold" });
            doc.AddText("Italic", new List<string>() { "italic" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic" });
            doc.AddText("This text without format", new List<string>() );


            doc.AddText("Bold", new List<string>() { "bold", "numbering" });
            doc.AddText("Italic", new List<string>() { "italic", "numbering" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54", "numbering" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic", "numbering" });
            doc.AddText("This text without format", new List<string>());

            doc.AddText("Bold", new List<string>() { "bold", "numbering", "level:0" });
            doc.AddText("Italic", new List<string>() { "italic", "numbering", "level:1" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54", "numbering", "level:1" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic", "numbering", "level:0" });
            doc.AddText("This text without format before page", new List<string>());

            doc.AddPageBreak();

            doc.AddText("Bold", new List<string>() { "bold", "bullet" });
            doc.AddText("Italic", new List<string>() { "italic", "bullet" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54", "bullet" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic", "bullet" });
            doc.AddText("This text without format after page", new List<string>());

            doc.AddText("Bold", new List<string>() { "bold", "bullet", "level:0" });
            doc.AddText("Italic", new List<string>() { "italic", "bullet", "level:1" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54", "bullet",  "level:1" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic", "bullet",  "level:2" });
            doc.AddText("This text without format", new List<string>());
            doc.AddRuledLine(60);


            doc.AddText("Bold", new List<string>() { "bold", "numbering" });
            doc.AddText("Italic", new List<string>() { "italic", "numbering" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54", "numbering" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic", "numbering" });
            doc.AddText("This text without format", new List<string>());
            doc.AddText("          d                                                ", new List<string>() { "underline"});
            doc.AddRuledLine(40);


            doc.Save();
            doc.Close();
        }

        [TestMethod]
        public void CreationWithImage()
        {
            WordServerDocument doc = new WordServerDocument();
            Assert.AreEqual(doc.Create($"{s_BasePath}\\testImage.docx", true, out _), OutputCode.OK);

            doc.AddText("Bold", new List<string>() { "bold" });

              doc.AddImage($"{s_BasePath}\\uno.gif", 100, 100);
                doc.AddImage($"{s_BasePath}\\dos.png", 50, 50);
            doc.AddImage($"{s_BasePath}\\tres.jpeg", 20, 20);

            doc.Save();
            doc.Close();
        }


        [TestMethod]
        public void CreationParagraphsWithRectangleShape()
        {
            WordServerDocument doc = new WordServerDocument();
            string filePath = $"{s_BasePath}\\rectangle-shape.docx";
            File.Delete(filePath);
            Assert.AreEqual(doc.Create($"{s_BasePath}\\rectangle-shape.docx", true, out _), OutputCode.OK);


            doc.AddText("Bold", new List<string>() { "bold" });
            doc.AddText("Italic", new List<string>() { "italic" });
            doc.AddText("FontSize 54", new List<string>() { "fontsize:54" });
            doc.AddText("color red italic", new List<string>() { "color:red", "italic" });
            doc.AddText("This text without format", new List<string>());
            doc.AddText("This text without format", new List<string>());
            doc.AddText("This text without format", new List<string>());
            doc.AddText("This text without format", new List<string>());

            doc.StartParagraph();
            doc.AddText("Inline text. ", new List<string>());
            doc.AddText("Inline text.", new List<string>());
            doc.EndParagraph();
            for (var i = 0; i < 10; i++)
            {
                doc.StartParagraph();
                doc.AddShapeWithText("", "R" + i, 300, 300, (0.15 * i), (-0.4 * i));
                doc.AddText("This text without format for line: " + i, new List<string>());
                doc.EndParagraph();
            }
            

            doc.Save();
            doc.Close();
        }
    }
}
