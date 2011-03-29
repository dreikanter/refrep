using System.IO;
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
			var fileName = Path.GetFullPath(@"{0}\..\..\..\TestData\references.xlsx");
 
			using (var reader = new ReferenceReader(fileName))
			{
				foreach (var row in reader)
				{
					Assert.IsNotNull(row);
				}
			}
		}
	}
}
