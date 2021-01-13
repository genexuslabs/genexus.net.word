using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MSWordDocument;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media.Imaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Genexus.Word.Shapes;

namespace Genexus.Word
{
	public class WordServerDocument : IDisposable
	{
		#region Private members
		// The underline document associated with this instance
		private WordprocessingDocument m_Document;
		// The underline dom main document associated with this document
		private MainDocumentPart m_DocumentPart;
		// The underline styles part associated to the document, we will create one for each document by default.
		private StyleDefinitionsPart m_StylesPart;
		// Numbering instance id
		private int m_LastNumberId = 0;
		// Should reset numbering
		private bool m_ResetNumbering = true;
		// The underline main body of the Document
		// Numbering document property id
		private uint m_LastDocumentPropertyId = 1;

		private Body m_Body;
		// The underline Styles part of the Document
		private Styles m_Styles;
		// Dispose pattern
		private bool disposedValue;
		#endregion

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
		{
			if (!File.Exists(fileName))
			{
				message = Messages.InvalidPath;
				return OutputCode.FILE_NOT_FOUND;
			}
			try
			{
				message = Messages.OK;
				OpenSettings settings = new OpenSettings
				{
					AutoSave = true
				};
				m_Document = WordprocessingDocument.Open(fileName, true, settings);
				m_DocumentPart = m_Document.MainDocumentPart;
				m_StylesPart = m_Document.MainDocumentPart.StyleDefinitionsPart;
				if (m_StylesPart != null)
					m_Styles = m_StylesPart.Styles;

				if (m_Document.MainDocumentPart != null && m_Document.MainDocumentPart.Document != null && m_Document.MainDocumentPart.Document.Body != null)
					m_Body = m_Document.MainDocumentPart.Document.Body;
				else
					ExceptionManager.LogException(Messages.NoBodyMessage);
				return OutputCode.OK;
			}
			catch (Exception ex)
			{
				message = ex.Message;
				ExceptionManager.HandleException(ex);
				return OutputCode.FAIL_OPEN;
			}
		}

		public void AddSampleCode()
		{
			TextBoxInfo2 textBoxInfo21 = new TextBoxInfo2();
			TextBoxContent textBoxContent1 = new TextBoxContent();

			Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00FC6179", RsidParagraphAddition = "00F60DF2", RsidParagraphProperties = "00DE46A1", RsidRunAdditionDefault = "00F60DF2", ParagraphId = "7C1AB0AA", TextId = "77777777" };

			ParagraphProperties paragraphProperties3 = new ParagraphProperties();
			Justification justification4 = new Justification() { Val = JustificationValues.Center };

			paragraphProperties3.Append(justification4);

			Run run7 = new Run() { RsidRunProperties = "00FC6179" };

			RunProperties runProperties2 = new RunProperties();
			RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

			runProperties2.Append(runFonts11);
			Text text2 = new Text();
			text2.Text = "H";

			run7.Append(runProperties2);
			run7.Append(text2);

			paragraph7.Append(paragraphProperties3);
			paragraph7.Append(run7);

			textBoxContent1.Append(paragraph7);
			textBoxInfo21.Append(textBoxContent1);


		}


		/// <summary>
		/// Create a word document
		/// </summary>
		/// <param name="fileName">Path to new document</param>
		/// <param name="overwriteIfExists">when true the file is created even if the file already exists on disk, overwriting the existing content</param>
		/// <param name="message">additional information related with the output code</param>
		/// <returns></returns>
		public int Create(string fileName, bool overwriteIfExists, out string message)
		{
			if (File.Exists(fileName) && !overwriteIfExists)
			{
				message = Messages.FileAlreadyExists;
				return OutputCode.FILE_ALREADY_EXISTS;
			}

			try
			{
				message = Messages.OK;
				m_Document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);

				// When creating a document we are creating the main document part and a body in order to easy further creational operations over the document.

				// Add a new main document part. 
				m_DocumentPart = m_Document.AddMainDocumentPart();
				//Create DOM tree for simple document. 
				m_DocumentPart.Document = new Document();
				m_Body = new Body();
				m_DocumentPart.Document.Append(m_Body);
				AddStylesPartToPackage();
				AddNumberingPartToPackage();
				// Save changes to the main document part. 
				m_DocumentPart.Document.Save();
				return OutputCode.OK;
			}
			catch (Exception ex)
			{
				message = ex.Message;
				ExceptionManager.HandleException(ex);
				return OutputCode.FAIL_CREATE;
			}
		}

		/// <summary>
		/// Save the current document state
		/// </summary>
		public void Save()
		{
			m_Document.Save();
		}


		/// <summary>
		/// Save the current document in other file, consider that previously to save to other document the current document is saved.
		/// </summary>
		/// <param name="fileName"></param>
		public void SaveAs(string fileName)
		{
			m_Document.SaveAs(fileName);
		}

		/// <summary>
		/// Close the file
		/// </summary>
		public void Close()
		{
			if (m_Document != null)
				m_Document.Close();
		}

		#region Styles

		// Add a StylesDefinitionsPart to the document.  Returns a reference to it.
		private void AddStylesPartToPackage()
		{
			m_StylesPart = m_DocumentPart.AddNewPart<StyleDefinitionsPart>();
			m_Styles = new Styles();
			m_Styles.Save(m_StylesPart);
		}

	
		#endregion

		#region openxml parts management

		/// <summary>
		/// Add Image part for the given filepath
		/// </summary>
		/// <param name="filepath"></param>
		/// <returns></returns>
		private string AddImagePart(string filepath)
		{
			if (!Enum.TryParse<ImagePartType>(Path.GetExtension(filepath), out ImagePartType imageType))
			{
				imageType = ImagePartType.Png;
			}
			ImagePart ip = m_DocumentPart.AddImagePart(imageType);
			using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
			{
				if (fs.Length == 0) return string.Empty;
				ip.FeedData(fs);
			}
			return m_DocumentPart.GetIdOfPart(ip);
		}

		/// <summary>
		/// Add a Numbering part to the package, we are adding 2 numbering formats + a bullet format
		/// </summary>
		private void AddNumberingPartToPackage()
		{
			NumberingDefinitionsPart numberingPart = m_DocumentPart.AddNewPart<NumberingDefinitionsPart>("gxNumberingPart");
			NumberingPart.GenerateNumberingDefinitionsPart1Content(numberingPart);
		}



		#endregion


		#region Edition Methods

		/// <summary>
		/// Add an image to the document with the original image width and height
		/// </summary>
		/// <param name="filePath"></param>
		public void AddImage(string filePath)
		{
			AddImageImpl(filePath);
		}

		/// <summary>
		/// Add an image to the document with the given width and height
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		public void AddImage(string filePath, int width, int height)
		{
			AddImageImpl(filePath, width, height);
		}

		private int AddImageImpl(string filePath, int width = -1, int height = -1)
		{
			int result = OutputCode.OK;
			try
			{
				string id = AddImagePart(filePath);
				AddImageToBody(id, filePath, width, height);
			}
			catch (FileNotFoundException)
            {
				result = AddImageOutputCode.IMAGE_NOT_FOUND;
			}
			return result;
		}

		/// <summary>
		/// Add a Style to the document with the given name and style properties. After adding the style it can be referenced by using AddTextWithStyle
		/// </summary>
		/// <param name="styleId">The Style id</param>
		/// <param name="styleName">The Style name</param>
		/// <param name="properties">Properties</param>
		public void AddStyle(string styleId, string styleName, List<string> properties)
		{
			Style style = new Style()
			{
				Type = StyleValues.Paragraph,
				StyleId = styleId,
				CustomStyle = true
			};
			StyleName styleName1 = new StyleName() { Val = styleName };
			style.Append(styleName1);
			StyleRunProperties styleRunProperties1 = new StyleRunProperties();
			styleRunProperties1.Append(GetProperties(properties));
			style.Append(styleRunProperties1);

			m_Styles.Append(style);
			m_Styles.Save(m_StylesPart);
		}


		/// <summary>
		/// Add the given text with the given style
		/// </summary>
		/// <param name="text">The text to add</param>
		/// <param name="styleId"></param>
		/// <returns></returns>
		public int AddTextWithStyle(string text, string styleId)
		{
			if (m_Document == null || m_Body == null)
				return OutputCode.INVALID_OPERATION;
			Paragraph p = new Paragraph
			{
				ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = styleId })
			};
			Run r = new Run();
			p.Append(r);
			Text t = new Text(text);
			r.Append(t);
			m_Body.Append(p);
			return OutputCode.OK;
		}

		/// <summary>
		/// Add a page break to the current document
		/// </summary>
		/// <returns></returns>
		public int AddPageBreak()
		{
			Paragraph paragraph = new Paragraph();
			Run run = new Run();
			Break br = new Break() { Type = BreakValues.Page };
			run.Append(br);
			paragraph.Append(run);
			m_Body.Append(paragraph);
			return OutputCode.OK;
		}


		/// <summary>
		/// Add text with the given properties 
		/// </summary>
		/// <param name="text">The text to be added</param>
		/// <param name="properties">Properties to be set for the new paragraph</param>
		/// <returns></returns>
		public int AddText(string text, List<string> properties)
        {
            if (m_Document == null || m_Body == null)
                return OutputCode.INVALID_OPERATION;
            Paragraph p = new Paragraph();
			p.Append(GetTextRun(p, text, properties));
			m_Body.Append(p);
            return OutputCode.OK;
        }

        private Run GetTextRun(Paragraph parentP, string text, List<string> properties)
        {
            if (properties.Contains("numbering") || properties.Contains("bullet"))
            {
                AddNumberingProperties(properties, parentP);
            }
            else
                m_ResetNumbering = true;

            Run r = new Run();
            r.Append(new RunProperties(GetProperties(properties)));
            
            Text t = new Text(text);
            r.Append(t);
			return r;
        }

        /// <summary>
        /// Numbering properties added to the given paragraph
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="p"></param>
        private void AddNumberingProperties(List<string> properties, Paragraph p)
		{
			if (properties.Contains("numbering") && m_ResetNumbering)
			{
				m_LastNumberId++;
				if (m_LastNumberId == 3)
					m_LastNumberId++;
				m_ResetNumbering = false;
			}
			NumberingProperties numberingProperties1 = new NumberingProperties();
			NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = GetNumberingLevel(properties) };
			NumberingId numberingId1 = new NumberingId() { Val = properties.Contains("bullet") ? 3 : m_LastNumberId };
			numberingProperties1.Append(numberingLevelReference1);
			numberingProperties1.Append(numberingId1);
			p.ParagraphProperties = new ParagraphProperties();
			p.ParagraphProperties.Append(numberingProperties1);
		}


		/// <summary>
		/// Add a ruled line with the specified length
		/// </summary>
		/// <param name="size"></param>
		public void AddRuledLine(int size)
		{
			Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "005D4D9C", RsidParagraphAddition = "00E416D0", RsidParagraphProperties = "005D4D9C", RsidRunAdditionDefault = "0044322C", ParagraphId = "4190A0B2", TextId = "72E25AC7" };
			ParagraphProperties paragraphProperties1 = new ParagraphProperties();
			ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
			Underline underline1 = new Underline() { Val = UnderlineValues.Single };
			paragraphMarkRunProperties1.Append(underline1);
			paragraphProperties1.Append(paragraphMarkRunProperties1);
			Run run1 = new Run();
			RunProperties runProperties1 = new RunProperties();
			Underline underline2 = new Underline() { Val = UnderlineValues.Single };
			runProperties1.Append(underline2);
			Text text1 = new Text();
			char nbrsp = '\u2007';
			text1.Text = new string(nbrsp, size);
			run1.Append(runProperties1);
			run1.Append(text1);
			paragraph2.Append(paragraphProperties1);
			paragraph2.Append(run1);
			m_Body.Append(paragraph2);
		}

		/// <summary>
		/// The properties can specify a level by adding a "level:" plus a number that indicate the nesting level. ie: level:3 
		/// </summary>
		/// <param name="properties"></param>
		/// <returns></returns>
		private int GetNumberingLevel(List<string> properties)
		{
			try
			{
				foreach (string prop in properties)
				{
					if (prop.ToLower().Contains("level:"))
						return int.Parse(prop.ToLower().Replace("level:", "").Trim());
				}
			}
			catch
			{

			}
			return 0;
		}

		/// <summary>
		/// Add text with te given basic style
		/// </summary>
		/// <param name="text"></param>
		/// <param name="style"></param>
		/// <returns></returns>
		public int AddTextWithBasicStyle(string text, BasicStyle style)
		{
			return AddText(text, style.GetProperties());
		}

		/// <summary>
		/// This function is in general used as a complement for previous addimagepart call, the width and height of the image is used
		/// </summary>
		/// <param name="relationshipId"></param>
		/// <param name="fileName"></param>
		private void AddImageToBody(string relationshipId, string fileName)
		{
			AddImageToBody(relationshipId, fileName, -1, -1);
		}

		/// <summary>
		/// Add the image with the specified width and height
		/// </summary>
		/// <param name="relationshipId">This is the reference to some graphic part added to the document</param>
		/// <param name="fileName">the path to an image</param>
		/// <param name="width">width in pixels</param>
		/// <param name="height">height in pixels</param>
		private void AddImageToBody(string relationshipId, string fileName, double width, double height)
		{
			using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
			{
				var img = new BitmapImage();
				img.BeginInit();
				img.StreamSource = fs;
				img.CacheOption = BitmapCacheOption.OnLoad;
				img.EndInit();

				width = (width >= 0) ? width: img.PixelWidth;
				height = (height >= 0) ? height : img.PixelHeight;

				var horzRezDpi = img.DpiX;
				var vertRezDpi = img.DpiY;
				var element = GetImageElement(relationshipId, fileName, Path.GetFileNameWithoutExtension(fileName), width, height, horzRezDpi, vertRezDpi);
				// Append the reference to body, the element should be in a Run.
				m_DocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
			}
		}

		/// <summary>
		/// Create a drawing element for the given fileName
		/// </summary>
		/// <param name="imagePartId"></param>
		/// <param name="fileName"></param>
		/// <param name="pictureName"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		/// <param name="horzRezDpi"></param>
		/// <param name="vertRezDpi"></param>
		/// <returns></returns>
		private static Drawing GetImageElement(string imagePartId, string fileName, string pictureName, double width, double height, double horzRezDpi, double vertRezDpi)
		{
			double emuWidth = width * Constants.EnglishMetricUnitsPerInch / horzRezDpi;
			double emuHeight = height * Constants.EnglishMetricUnitsPerInch / vertRezDpi;

			var element = new Drawing(
				new DW.Inline(
					new DW.Extent { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight },
					new DW.EffectExtent { LeftEdge = 400L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
					new DW.DocProperties { Id = (UInt32Value)1U, Name = pictureName },
					new DW.NonVisualGraphicFrameDrawingProperties(
					new A.GraphicFrameLocks { NoChangeAspect = true }),
					new A.Graphic(
						new A.GraphicData(
							new PIC.Picture(
								new PIC.NonVisualPictureProperties(
									new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = fileName },
									new PIC.NonVisualPictureDrawingProperties()),
								new PIC.BlipFill(
									new A.Blip(
										new A.BlipExtensionList(
											new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
									{
										Embed = imagePartId,
										CompressionState = A.BlipCompressionValues.Print
									},
									new A.Stretch(new A.FillRectangle())),
								new PIC.ShapeProperties(
									new A.Transform2D(
										new A.Offset { X = 0L, Y = 0L },
										new A.Extents { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight }),
									new A.PresetGeometry(
										new A.AdjustValueList())
									{ Preset = A.ShapeTypeValues.Rectangle })))
						{
							Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
						}))
				{
					DistanceFromTop = (UInt32Value)0U,
					DistanceFromBottom = (UInt32Value)0U,
					DistanceFromLeft = (UInt32Value)0U,
					DistanceFromRight = (UInt32Value)0U,
					EditId = "50D07946"
				});
			return element;
		}



		/// <summary>
		/// Get properties from string properties
		/// </summary>
		/// <param name="properties"></param>
		/// <returns></returns>
		private IEnumerable<OpenXmlElement> GetProperties(List<string> properties)
		{
			foreach (string prop in properties)
			{
				if (StyleProperties.Exists(prop))
					yield return StyleProperties.RunFunctionProperty(prop);

			}
		}

		/// <summary>
		/// Replace the <paramref name="searchText"/> with the <paramref name="replaceText"/>
		/// </summary>
		/// <param name="searchText">the text to search</param>
		/// <param name="replaceText">the new value</param>
		/// <param name="matchCase">if true only match cases ocurrences are taking into account</param>
		/// <returns></returns>
		public int ReplaceText(string searchText, string replaceText, bool matchCase)
		{
			if (m_Document == null || m_Body == null)
				return OutputCode.INVALID_OPERATION;

			// Avoid to use PowerTools at this time
			//TextReplacer.SearchAndReplace(m_Document, searchText, replaceText, matchCase);
			return ReplaceTextWithBasicStyle(searchText, replaceText, matchCase, new BasicStyle());
		}

		/// <summary>
		/// Replace the <paramref name="searchText"/> with the <paramref name="replaceText"/> adding the style given in <paramref name="style"/>
		/// </summary>
		/// <param name="searchText"></param>
		/// <param name="replaceText"></param>
		/// <param name="matchCase"></param>
		/// <param name="style"></param>
		/// <returns></returns>
		public int ReplaceTextWithBasicStyle(string searchText, string replaceText, bool matchCase, BasicStyle style)
		{
			return ReplaceTextWithStyle(searchText, replaceText, matchCase, style.GetProperties());
		}


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
		{
			if (m_Document == null || m_Body == null)
				return 0;

			int count = 0;
			var paras = m_Body.Elements<Paragraph>();
			foreach (var para in paras)
			{
				foreach (var run in para.Elements<Run>())
				{
					foreach (var text in run.Elements<Text>())
					{
						string currentText = matchCase ? text.Text : text.Text.ToLower();
						searchText = matchCase ? searchText : searchText.ToLower();
						int startIndex = currentText.IndexOf(searchText);
						if (startIndex >= 0)
						{
							string newText;
							if (startIndex > 0)
								newText = currentText.Substring(0, startIndex);
							else
								newText = String.Empty;
							newText += replaceText;
							newText += currentText.Substring(startIndex + searchText.Length);
							// Prepare the new properties for the text
							RunProperties newProperties = new RunProperties(GetProperties(properties));
							newProperties.Append(new Text(newText));

							run.PrependChild<RunProperties>(newProperties);
							text.Remove();
							count++;
						}
					}
				}
			}
			return count;
		}


	

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
		{
			if (m_Document == null || m_Body == null)
				return 0;

			string id = AddImagePart(imageFile);

			// Prepare the new properties for the text
			RunProperties newProperties = new RunProperties();
			var img = new BitmapImage();
			img.BeginInit();
			img.UriSource = new Uri(imageFile, UriKind.RelativeOrAbsolute);
			img.CacheOption = BitmapCacheOption.OnLoad;
			img.EndInit();
			var horzRezDpi = img.DpiX;
			var vertRezDpi = img.DpiY;
			
			
			

			newProperties.Append(GetImageElement(id, imageFile, "test", width, height, horzRezDpi, vertRezDpi));

			int count = 0;
			var paras = m_Body.Elements<Paragraph>();
			foreach (var para in paras)
			{
				foreach (var run in para.Elements<Run>())
				{
					foreach (var text in run.Elements<Text>())
					{
						if (string.Compare(text.Text, searchText, !matchCase) == 0)
						{
							run.PrependChild(newProperties);
							text.Remove();
							count++;
						}
					}
				}
			}
			return count;
		}


		/// <summary>
		/// Adds predefined shape <paramref name="shapeId"/> with a custom text <paramref name="shapeInnetText"/>
		/// </summary>
		/// <param name="shapeId"></param>
		/// <param name="shapeInnetText"></param>
		/// <param name="width"></param>
		/// <param name="height"></param>
		/// <returns></returns>
		public int AddShapeWithText(string shapeId, string shapeText, string text, double width, double height, List<string> properties = null)
		{
			if (m_Document == null || m_Body == null)
				return 0;
			
			Paragraph p = new Paragraph();
			
			Run r = new Run(new RunProperties(new NoProof()));
			
			r.Append(CustomShapeBuilder.BuildRectangle(m_DocumentPart, m_LastDocumentPropertyId++, shapeText, width, height));
			p.Append(r);
			p.Append(GetTextRun(p, text, properties));

			m_DocumentPart.Document.Body.AppendChild(p);
			return OutputCode.OK;
		}

		#endregion

		#region Dispose

		/// <summary>
		///  When disposing close the document
		/// </summary>
		/// <param name="disposing"></param>
		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					if (m_Document != null)
					{
						m_Document.Close();
						m_Document = null;
					}
				}

				disposedValue = true;
			}
		}


		/// <summary>
		/// Dispose method
		/// </summary>
		public void Dispose()
		{
			// Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
			Dispose(disposing: true);
			GC.SuppressFinalize(this);
		}

		#endregion
	}
}
