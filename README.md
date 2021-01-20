# genexus.net.word
GeneXus Word is an implementation of several functions for Server side generation of Microsoft docx documents.

This implementation is based on OpenXML.


## Classes

| **Class** | **Description** |
| :---: | :---: |
| WordServerDocument | The main class for creating Word documents |
| BasicStyle | There are some basic style that can be configured by using this class |
| OutputCodes | Output codes for methods in WordSeverDocument |

## WordServerDocument class

### Create , Open , Close, Save

With the following functions you can create or open a word document. Remember to Close the document after your operations over the document are finished.

```cs
/// <summary>
/// Open the document specified in fileName
/// </summary>
/// <param name="fileName">Is the path and name of the document. If only the name is specified, the document will be created in 
/// the default directory (DataXXX model directory).</param>
/// <param name="message">additional information related with the output code</param>
/// <returns>
/// 
///     0   -  OK
///     6   -  Could not open file
///     10  -  Could not complete operation
/// </returns>
public int Open(string fileName, out string message)

/// <summary>
/// Create a word document
/// </summary>
/// <param name="fileName">Path to new document</param>
/// <param name="overwriteIfExists">when true the file is created even if the file already exists on disk, 
/// overwriting the existing content</param>
/// <param name="message">additional information related with the output code</param>
/// <returns></returns>
public int Create(string fileName, bool overwriteIfExists, out string message)



/// <summary>
/// Save the current document state
/// </summary>
public void Save()

/// <summary>
/// Save the current document in other file, consider that previously to save to other document the current document is saved.
/// </summary>
/// <param name="fileName"></param>
public void SaveAs(string fileName)

/// <summary>
/// Close the file
/// </summary>
public void Close()
```


### Edition Methods

Several edition methods receive a List of style properties as parameter. 
The options to send in the list are:

- bold
- italic
- caps
- smallcaps
- strike
- doublestrike
- outline
- shadow
- underline
- fontsize:<number>               ie: fontsize:14
- color:<color name>              ie: color:red
- fontfamily:<font name>          ie: fontfamily:Arial
- numbering
- bullet
- level:<number>                  ie:  level:2  ( level number start with 0 and is the default level number for numbering and bullets)
  
           
```cs
/// <summary>
/// Add an image to the document with the original image width and height
/// </summary>
/// <param name="filePath"></param>
public void AddImage(string filePath)


/// <summary>
/// Add an image to the document with the given width and height
/// </summary>
/// <param name="filePath"></param>
/// <param name="width"></param>
/// <param name="height"></param>
public void AddImage(string filePath, int width, int height)

/// <summary>
/// Add a Style to the document with the given name and style properties. After adding the style it can be referenced by using AddTextWithStyle
/// </summary>
/// <param name="styleId">The Style id</param>
/// <param name="styleName">The Style name</param>
/// <param name="properties">Properties</param>
public void AddStyle(string styleId, string styleName, List<string> properties)


/// <summary>
/// Add the given text with the given style
/// </summary>
/// <param name="text">The text to add</param>
/// <param name="styleId"></param>
/// <returns></returns>
public int AddTextWithStyle(string text, string styleId)

/// <summary>
/// Add a page break to the current document
/// </summary>
/// <returns></returns>
public int AddPageBreak()

/// <summary>
/// Add text with the given properties 
/// </summary>
/// <param name="text">The text to be added</param>
/// <param name="properties">Properties to be set for the new paragraph</param>
/// <returns></returns>
public int AddText(string text, List<string> properties)

/// <summary>
/// Add text with te given basic style
/// </summary>
/// <param name="text"></param>
/// <param name="style"></param>
/// <returns></returns>
public int AddText(string text, BasicStyle style)

/// <summary>
/// Add a ruled line with the specified length
/// </summary>
/// <param name="size"></param>
public void AddRuledLine(int size)

/// <summary>
/// Replace the <paramref name="searchText"/> with the <paramref name="replaceText"/>
/// </summary>
/// <param name="searchText">the text to search</param>
/// <param name="replaceText">the new value</param>
/// <param name="matchCase">if true only match cases ocurrences are taking into account</param>
/// <returns></returns>
public int ReplaceText(string searchText, string replaceText, bool matchCase)


/// <summary>
/// Replace the <paramref name="searchText"/> with the <paramref name="replaceText"/> adding the style given in <paramref name="style"/>
/// </summary>
/// <param name="searchText"></param>
/// <param name="replaceText"></param>
/// <param name="matchCase"></param>
/// <param name="style"></param>
/// <returns></returns>
public int ReplaceTextWithStyle(string searchText, string replaceText, bool matchCase, BasicStyle style)

/// <summary>
/// Search for a given text  in <paramref name="searchText"/> and replace it with the <paramref name="replaceText"/> with the specified format in properties. Take into account tha this function
/// is not considering text that are splited in different Runs, so that if you have a document with mixed format for the same word this function is
/// not appropiated.
/// Tha algorithm for doing this well is complex and is implemented in the ReplaceText function but replacing the text without replacing the format.
/// </summary>
/// <param name="searchText"></param>
/// <param name="replaceText"></param>
/// <param name="matchCase"></param>
/// <param name="properties"></param>
/// <returns></returns>
public int ReplaceTextWithStyle(string searchText, string replaceText, bool matchCase, List<string> properties)

/// <summary>
/// Replace a given text in <paramref name="searchText"/> by an image specified in <paramref name="imageFile"/>
/// </summary>
/// <param name="searchText"></param>
/// <param name="matchCase"></param>
/// <param name="imageFile"></param>
/// <param name="width"></param>
/// <param name="height"></param>
/// <returns></returns>
public int ReplaceTextWithImage(string searchText, bool matchCase, string imageFile, double width, double height)

 /// <summary>
/// Adds predefined shape <paramref name="shapeId"/> with a custom inner shape text <paramref name="shapeText"/>. Only rectangle is supported
/// </summary>
/// <param name="shapeId"></param>
/// <param name="shapeInnetText"></param>
/// <param name="width"></param>
/// <param name="height"></param>
/// <param name="posLeft">Left Position of the Shape (in cm)</param>
/// <param name="posTop">Top Position of the Shape (in cm)</param>
/// <param name="shapeProperties">Shape style properties</param>
/// <returns></returns>
public int AddShapeWithText(string shapeId, string shapeText, double width, double height, double posLeft = 0, double posTop = 0, List<string> shapeProperties = null, List<string> textProperties = null)
        
```


### Samples

```cs
      [TestMethod]
        public void ReplaceAndAddWithStyle()
        {
            using (WordServerDocument doc = new WordServerDocument())
            {
                File.Copy($"{s_BasePath}\\TemplateSample.docx", $"{s_BasePath}\\Sample.docx", true);
              
                Assert.AreEqual(doc.Open($"{s_BasePath}\\Sample.docx", out _), OutputCode.OK);
            

                Assert.AreEqual(doc.ReplaceTextWithStyle("REP01", "新潟県新潟市", true, new List<string>() { "color:blue", "italic" , "fontsize:24", "fontfamily:MS PMincho" }), 1);
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


                Assert.AreEqual(doc.ReplaceTextWithStyle("REP01", "新潟県新潟市", true, new List<string>() { "color:blue", "italic", "fontsize:24", "fontfamily:MS PMincho" }), 1);
                Assert.AreEqual(doc.ReplaceTextWithStyle("REP02", "中央区米山", true, new List<string>() { "italic", "bold", "color:green", "fontfamily:MS PGothic" }), 1);
                Assert.AreEqual(doc.ReplaceTextWithImage("REP03", false, "d:\\temp\\dos.png", 50, 50), 1);


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
        public void CreateRectangleShape()
        {
            WordServerDocument doc = new WordServerDocument();
            string filePath = $"{s_BasePath}\\rectangle-shape.docx";
            File.Delete(filePath);
            Assert.AreEqual(doc.Create($"{s_BasePath}\\rectangle-shape.docx", true, out _), OutputCode.OK);

            doc.AddShapeWithText("", "SQUARE", 3, 3, 0, 0, new List<string>() {
                    "strokewidth:5",
                    "color:#32a852",
                    "fillcolor:silver"
                }
            , new List<string>() {
                    "fontsize:30",
                    "color:red"
                });

           
            doc.Save();
            doc.Close();
        }
        
        [TestMethod]
        public void CreateRectangleShapeWithSibilingText()
        {
            WordServerDocument doc = new WordServerDocument();
            string filePath = $"{s_BasePath}\\rectangle-shape-2.docx";
            File.Delete(filePath);
            Assert.AreEqual(doc.Create(filePath, true, out _), OutputCode.OK);

            doc.StartParagraph();
            
            doc.AddShapeWithText("", "TXT", 1.5, 1, -1, 0.15);
            doc.AddText("Item 1", new List<string>());
            doc.EndParagraph();
           
            doc.Save();
            doc.Close();
        }
    }
```



