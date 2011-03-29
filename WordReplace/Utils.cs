using System;
using System.Diagnostics;
using System.IO;
using WordReplace.Extensions;

namespace WordReplace
{
	public static class Utils
	{
		public static string GetExecPath()
		{
			var mainModule = Process.GetCurrentProcess().MainModule;
			return mainModule == null ? String.Empty : Path.GetDirectoryName(mainModule.FileName);
		}

		/// <summary>
		/// Generates a new name for processed document.
		/// </summary>
		public static string GetNewName(string sourceFileName)
		{
			string result;
			var counter = 1;

			do
			{
				result = Path.Combine(Path.GetDirectoryName(sourceFileName), 
					"{0}-updated{1}{2}".Fill(Path.GetFileNameWithoutExtension(sourceFileName),
						(counter++ > 1) ? counter.ToString() : String.Empty, Path.GetExtension(sourceFileName)));
			}
			while (File.Exists(result));
			
			return result;
		}
	}
}
