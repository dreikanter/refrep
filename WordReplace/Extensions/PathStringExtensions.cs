using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using WordReplace.Auxiliary;

namespace WordReplace.Extensions
{
	public static class PathStringExtensions
	{
		public static long GetFileSize(this string fileName)
		{
			return new FileInfo(fileName).Length;
		}

		public static string ToFormattedFileSize(this long size)
		{
			return String.Format(new FileSizeFormatProvider(), "{0:fs}", size);
		}

		[DllImport("shlwapi.dll", CharSet = CharSet.Auto)]
		static extern bool PathCompactPathEx([Out] StringBuilder pszOut, string szPath, int cchMax, int dwFlags);

		public static string ReducePath(this string path, int length)
		{
			var sb = new StringBuilder();
			PathCompactPathEx(sb, path, length, 0);
			return sb.ToString();
		}
	}
}
