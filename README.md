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

### Create or Open

```cs
/// <summary>
/// Open the document specified in fileName
/// </summary>
/// <param name="fileName">Is the path and name of the document. If only the name is specified, the document will be created in the default directory (DataXXX model directory).</param>
/// <param name="message">additional information related with the output code</param>
/// <returns>
/// 
///     0   -  OK
///     6   -  Could not open file
///     10  -  Could not complete operation
/// </returns>
public int Open(string fileName, out string message)
```

```cs
/// <summary>
  /// Create a word document
  /// </summary>
  /// <param name="fileName">Path to new document</param>
  /// <param name="overwriteIfExists">when true the file is created even if the file already exists on disk, overwriting the existing content</param>
  /// <param name="message">additional information related with the output code</param>
  /// <returns></returns>
  public int Create(string fileName, bool overwriteIfExists, out string message)
```


### Edition Methods

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
```



