using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordsCommentsExtractor
{
	public class WordDocument
	{
		private string name;
		private string path;
		private char delimeter = '/';
		private readonly WordprocessingDocument wordDocument;

		//public WordDocument(string _path, string _name)
		//{
		//	DelimiterSwitch();
		//	name = _name;
		//	path = _path;
		//}

		public WordDocument(string _path)
		{
			DelimiterSwitch();
			path = _path;
			try
			{
				string[] _pathArr = _path.Split(delimeter);
				name = _pathArr[_pathArr.Length - 1];
				wordDocument = WordprocessingDocument.Open(path, false);
				

			}
			catch (Exception ex)
			{
				Console.WriteLine("Coudn't find the document: ", ex);
			}
		}

		public void Display()
		{
			Console.WriteLine("Word document has the name: " + name + " and the path is: " + path);
		}

		public void DeleteContentControls()
		{
			MainDocumentPart main = wordDocument.MainDocumentPart;
			Console.WriteLine("Looking for Content controls to be deleted");
			//SdtBlock[] sdtBlock = main.Document.Body.Descendants<SdtBlock>().ToArray();
			//Console.WriteLine(sdtBlock.Length);
			IEnumerable<SdtRun> sdtElement = main.Document.Body.Descendants<SdtRun>().ToArray();
			Console.WriteLine(sdtElement.Count());

			foreach (SdtRun sdtEl in sdtElement)
			{
				Console.WriteLine(sdtEl);
				sdtEl.Remove();
				Console.WriteLine("Deleting ContentControl Element");
			}

			Console.WriteLine(sdtElement.Count());

			//foreach (SdtBlock sdt in sdtBlock)
			//{
			//	sdt.Remove();
			//	Console.WriteLine("Deleting ContentCotnrols Blocks");
			//}


		}

		public void SaveAndClose()
		{
			wordDocument.Close();
		}

		public List<string> GetComments()
		{
			List<string> comments = new List<string>();

			WordprocessingCommentsPart commentsPart = wordDocument.MainDocumentPart.WordprocessingCommentsPart;
			if (commentsPart != null && commentsPart.Comments != null)
			{
				foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
				{
					comments.Add(comment.InnerText);
				}
			} else
			{
				return comments;
			}
			return comments;
		}

		public RecordsList GetCommentsWithText()
		{
			RecordsList records = new RecordsList();

			WordprocessingCommentsPart commentsPart = wordDocument.MainDocumentPart.WordprocessingCommentsPart;
			if (commentsPart != null && commentsPart.Comments != null)
			{
				foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
				{
					OpenXmlElement rangeStart = wordDocument.MainDocumentPart.Document.Descendants<CommentRangeStart>().Where(c => c.Id == comment.Id).FirstOrDefault();
					//bool breakLoop = false;
					//rangeStart = rangeStart.Parent;
					rangeStart = rangeStart.NextSibling();
					string commentText="";
					while (!(rangeStart is CommentRangeEnd))
					{
						try
						{
							if (!string.IsNullOrWhiteSpace(rangeStart.InnerText))
							commentText += rangeStart.InnerText;
							rangeStart = rangeStart.NextSibling();
						}
						catch (NullReferenceException ex)
						{
							Console.WriteLine(ex.Message);
							Console.WriteLine("NullReference Exception on " + comment.InnerText + " with highlited text: " + commentText);
							commentText += " !!!ERROR WHILE EXTRACTING THIS TEXT!!!";
							break;
						}
					}
					Record record = new Record(comment.Id, comment.InnerText, commentText);
					records.Add(record);

				}
			}
			else
			{
				return records;
			}
			return records;
		}


		private void DelimiterSwitch()
		{
			if (OSDetector.IsWindows())
			{
				Console.WriteLine("Running on Windows");
				delimeter = '\\';
			}
			Console.WriteLine("Running on Unix");
		}

	}
}
