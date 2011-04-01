using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordReplace.References;

namespace WordReplace.Test
{
	/// <summary>
	/// Summary description for UnitTest1
	/// </summary>
	[TestClass]
	public class ReferenceReaderTest
	{
		[TestMethod]
		public void ReadingTest()
		{
			using (var reader = new ReferenceReader(TestFileNames.RefFile))
			{
				foreach (var row in reader)
				{
					Assert.IsNotNull(row);
				}
			}
		}
	}
}
