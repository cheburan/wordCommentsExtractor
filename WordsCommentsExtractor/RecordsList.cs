using System;
using System.Collections.Generic;

namespace WordsCommentsExtractor
{
	public class RecordsList
	{
		public List<Record> records = new List<Record>();
		public int recordLengh = Record.Count();

		public RecordsList()
		{

		}

		public void Add(Record record)
		{
			records.Add(record);
		}

		public void ConsolePrint()
		{
			if (records.Count > 0)
			{
				Console.WriteLine("Number of comments in the file: " + records.Count);
				records.ForEach(delegate (Record record)
				{
					record.ConsolePrint();
				});
			}
			else
			{
				Console.WriteLine("No comments in the file");
			}
		}

		internal string[][] Transform(string[] titles)
		{
			if (records.Count > 0)
			{
				string[][] result = new string[records.Count+1][];
				result[0] = titles;
				for (int i = 1; i < result.Length; i++)
				{
					result[i] = records[i-1].Export();
				}
				return result;
			}
			else
			{
				throw new Exception("No comments in the file");
			}
		}
	}
}