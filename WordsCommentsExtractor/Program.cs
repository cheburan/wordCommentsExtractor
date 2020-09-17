using System;
using System.Collections.Generic;

namespace WordsCommentsExtractor
{
    public static class Extensions
    {
        public static T[] SubArray<T>(this T[] array, int offset, int length)
        {
            T[] result = new T[length];
            Array.Copy(array, offset, result, 0, length);
            return result;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string[] columns = {"A", "B", "C", "D", "E", "F", "G", "H", "I"};
            Console.Clear();
            Console.WriteLine("Please type the full address to the word document or folder with multiply word documents(docx) you want to exctract comments from:");
            string path = Console.ReadLine();
            Console.WriteLine("Thank you! Now Please provide the full puth to the Excel spreadsheet you want to create and put codes in to:");
            string excelPath = Console.ReadLine();
            Console.WriteLine("Thank you! \n\r Input from: " + path + "\n\r Output: " + excelPath);
            FilesProcessing files = new FilesProcessing(path);
            string[] titles = { "Id", "Comment", "Text" };
            ExcelDocument excelDocument = new ExcelDocument(excelPath, titles);
            excelDocument.Create();
            files.fileEntries.ForEach(delegate (TranscriptFile transcript)
            {
                transcript.ConsolePrint();
                WordDocument document = new WordDocument(transcript.path);
                document.DeleteContentControls();
				RecordsList records = document.GetCommentsWithText();
                //records.ConsolePrint();
                string[][] data = records.Transform(titles);
                columns = columns.SubArray(0, 3);
				//excelDocument.InsertWorksheet(transcript.title);
				excelDocument.InsertText(transcript.title, columns, data);
			});
        }
    }
}
