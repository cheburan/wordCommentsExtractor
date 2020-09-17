using System;
using System.Runtime.InteropServices;

namespace WordsCommentsExtractor
{
	static public class OSDetector
	{
		/**
		 * Check if the Program is running on OS Windows
		 */
		static public bool IsWindows()
		{
			if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
			{
				return true;
			}
			return false;
		}

		/**
		 * Check if the Program is running on Mac OS
		 */
		static public bool IsOSX()
		{
			if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
			{
				return true;
			}
			return false;
		}

		/**
		 * Check if the Program is running on OS Linux
		 */
		static public bool IsLinux()
		{
			if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
			{
				return true;
			}
			return false;
		}


	}
}
