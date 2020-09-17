using System;
namespace WordsCommentsExtractor
{
	public class Record : IEntity
	{
		public string comment;
		public string text;
		public string id;

		public Record(string _id, string _comment, string _text)
		{
			id = _id;
			comment = _comment;
			text = _text;
		}

		public void ConsolePrint()
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			Console.WriteLine(id + " - " + comment + ": " + text);
			Console.ForegroundColor = ConsoleColor.White;
		}

		public string[] Export()
		{
			return new string[] { id, comment, text };
		}

		//return number of attributes
		public static int Count()
		{
			return typeof(Record).GetProperties().Length;
		}
	}
}
