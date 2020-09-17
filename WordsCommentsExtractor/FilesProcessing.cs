using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace WordsCommentsExtractor
{
    public class FilesProcessing
    {
        private Regex tempPattern = new Regex(@"\~\$?");
        public bool filePath = true;
        public List<TranscriptFile> fileEntries = new List<TranscriptFile>();

        public FilesProcessing(string path)
        {
            if (File.Exists(path))
            {
                // This path is a file
                filePath = true;
            }
            else if (Directory.Exists(path))
            {
                // This path is a directory
                ProcessDirectory(path);
            }
            else
            {
                Console.WriteLine("No file or directory were found");
                throw new IOException("No valid file or directory were found.");
            }
        }

        private void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] files = Directory.GetFiles(targetDirectory, "*.docx");
            foreach (string fileName in files)
                ProcessFile(fileName);
        }

        //Returning path to the file
        private void ProcessFile(string path)
        {
            if (!tempPattern.IsMatch(path))
                fileEntries.Add(new TranscriptFile(path));
        }

        public void Print()
        {
            fileEntries.ForEach(delegate (TranscriptFile transcript)
            {
                transcript.ConsolePrint();
            });
        }

        public List<TranscriptFile> GetTranscriptFiles()
		{
            return fileEntries;
		}
    }
}