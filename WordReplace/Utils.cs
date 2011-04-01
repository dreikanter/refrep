using System;
using System.Diagnostics;
using System.IO;
using System.Text;
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

		/// <summary>
		/// Generates a clipboard contents for pasting by PasteSpecial() interop's method
		/// when DataType is WdPasteDataType.wdPasteHTML.
		/// </summary>
		public static string GetHtmlClipboardText(string text)
		{
			var html = @"Version:0.9
StartHTML:<<<<<<<1
EndHTML:<<<<<<<2
StartFragment:<<<<<<<3
EndFragment:<<<<<<<4
StartSelection:<<<<<<<3
EndSelection:<<<<<<<4
<!doctype><html><head><style>
body {
font-family:'Times New Roman';
font-size:14pt;
}
</style><title></title></head><body><!--Start-->" + text + @"<!--End--></body></html>";

			var s = Encoding.GetEncoding(0).GetString(Encoding.UTF8.GetBytes(html));
			var sb = new StringBuilder();

			sb.Append(s);
			sb.Replace("<<<<<<<1", (s.IndexOf("<html>") + "<html>".Length).ToString("D8"));
			sb.Replace("<<<<<<<2", (s.IndexOf("</html>")).ToString("D8"));
			sb.Replace("<<<<<<<3", (s.IndexOf("<!--Start-->") + "<!--Start-->".Length).ToString("D8"));
			sb.Replace("<<<<<<<4", (s.IndexOf("<!--End-->")).ToString("D8"));

			return sb.ToString();
		}
	}
}
