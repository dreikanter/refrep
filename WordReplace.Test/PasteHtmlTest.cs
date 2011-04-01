using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordReplace.Extensions;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordReplace.Test
{
	[TestClass]
	public class PasteHtmlTest
	{
		[TestMethod]
		public void PasteTest()
		{
			var fileName = TestFileNames.SourceFile;
			var word = new Application { Visible = false };
			var doc = word.Documents.Open(fileName);

			try
			{
				// ReSharper disable UseIndexedProperty
				var range = doc.Bookmarks.get_Item("Bibliography").Range;
				// ReSharper restore UseIndexedProperty

				var html = Utils.GetHtmlClipboardText("Hello <i>World!</i>");
				Clipboard.SetText(html, TextDataFormat.Html);
				range.PasteSpecial(DataType: WdPasteDataType.wdPasteHTML);

				var destFileName = Path.Combine(Path.GetDirectoryName(fileName), 
					Path.GetFileNameWithoutExtension(fileName) + "-updated" + Path.GetExtension(fileName));
				word.ActiveDocument.SaveAs(destFileName);
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.GetMessage());
			}
			finally
			{
				word.Quit(false);
			}
		}
	}
}
