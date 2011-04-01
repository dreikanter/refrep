using System.IO;

namespace WordReplace.Test
{
	public static class TestFileNames
	{
		public static string SourceFile
		{
			get { return Path.GetFullPath(@"..\..\..\TestData\source.docx"); }
		}

		public static string RefFile
		{
			get { return Path.GetFullPath(@"..\..\..\TestData\references.xlsx"); }
		}
	}
}
