using System;
namespace WordsCommentsExtractor
{
	public class TranscriptFile: IEntity
	{
		private char delimiter = '/';
		public string path, name, title;

		public TranscriptFile(string _path)
		{
			path = _path;
			name = _path.Substring(_path.LastIndexOf(delimiter) + 1);
			title = name.Split(".")[0];
		}

		private void DelimiterSwitch()
		{
			if (OSDetector.IsWindows())
			{
				delimiter = '\\';
			}
		}

		 public void ConsolePrint()
		{
			Console.WriteLine("File: " + name + "   Interviewer: " + title);
			Console.WriteLine(" [" + path + "]");
			Console.WriteLine("");
		}
	}
}
