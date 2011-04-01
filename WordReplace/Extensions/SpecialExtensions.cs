using System;
using System.Collections.Generic;
using System.Linq;
using WordReplace.References;

namespace WordReplace.Extensions
{
	/// <summary>
	/// Application domain-specific extension methods.
	/// </summary>
	public static class SpecialExtensions
	{
		public static bool Defined(this int? value)
		{
			return value != null;
		}

		public static bool Defined(this string value)
		{
			return !String.IsNullOrEmpty(value);
		}

		public static bool Defined(this DateTime? value)
		{
			return value != null;
		}

		/// <summary>
		/// Returns comma-separated list of sorted reference numbers.
		/// </summary>
		public static string GetOrderedRefNumList(this IEnumerable<Reference> refs)
		{
			var refNums = from r in refs orderby r.RefNum ascending select r.RefNum.ToString();
			return refNums.CommaSeparatedNb(false);
		}
	}
}