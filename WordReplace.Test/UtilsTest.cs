using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WordReplace.Test
{
	[TestClass]
	public class UtilsTest
	{
		[TestMethod]
		public void GetExecPath()
		{
			var path = Utils.GetExecPath();
			Assert.IsNotNull(path);
			Assert.IsTrue(Directory.Exists(path));
		}

		[TestMethod]
		public void TestGetNewName()
		{
			var sourceFile = Path.Combine(Utils.GetExecPath(), "source.docx");
			var newFile = Utils.GetNewName(sourceFile);
			Assert.IsNotNull(newFile);
			Assert.IsFalse(File.Exists(newFile));
			Assert.IsTrue(newFile.Length > sourceFile.Length);
		}
	}
}
