using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordReplace.Extensions;

namespace WordReplace.Test
{
	[TestClass]
	public class StringExtensionsTest
	{
		[TestMethod]
		public void TitleComparisionTest()
		{
			var values = new Dictionary<string, string>
			             	{
			             		{"Иванов", "Петров"},
			             		{"Johnson", "Wilson"},
			             		{"Петров", "Wilson"},
			             		{"Барабанов", ""},
			             		{"Иванов А.Б. Cold fusion in Siberia", "Johnson M. Some book title"},
			             	};

			foreach(var pair in values)
			{
				Assert.IsTrue(pair.Key.CompareTitleStrings(pair.Value) < 0);
				Assert.IsTrue(pair.Value.CompareTitleStrings(pair.Key) > 0);
			}

			Assert.IsTrue(((string) null).CompareTitleStrings("Non-null value") > 0);
			Assert.IsTrue("Non-null value".CompareTitleStrings(null) < 0);
		}
	}
}
