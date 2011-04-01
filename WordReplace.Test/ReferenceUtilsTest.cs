using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordReplace.References;

namespace WordReplace.Test
{
	[TestClass]
	public class ReferenceUtilsTest
	{
		[TestMethod]
		public void GetAuthorsListTest()
		{
			var values = new Dictionary<string, string>
			             	{
			             		{"Иванов Петр Васильевич", "Иванов П.В."},
			             		{"Васильев Иван Петрович", "Васильев И.П."},
			             		{"Иванов Петр", "Иванов П."},
			             		{"Williams John", "Williams J."},
			             		{"Иванов", "Иванов"},
			             		{"", ""},
			             	};

			foreach(var pair in values)
			{
				Assert.IsTrue(ReferenceUtils.GetAuthorsList(pair.Key, false, false) == pair.Value);
			}
		}
	}
}
