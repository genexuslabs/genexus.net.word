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





