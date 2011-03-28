using System;
using System.Collections.Generic;

namespace WordReplace.Extensions
{
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
		/// Returns a shortened list of author names.
		/// </summary>
		public static string GetAuthorsList(this string authors, bool initialsInFront)
		{
			var result = new List<string>();

			foreach (var author in authors.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
			{
				var nameParts = new List<string>();

				foreach (var part in author.Split(new[] { " ", "\t", "." }, StringSplitOptions.RemoveEmptyEntries))
				{
					nameParts.Add((nameParts.Count == 1) ? part.Trim().Capitilize() : "{0}.".Fill(part.GetInitial()));
				}

				if (initialsInFront)
				{
					nameParts.Add(nameParts[0]);
					nameParts.RemoveAt(0);
				}

				result.Add(nameParts.JoinWith(Constants.NbSp));
			}

			return result.CommaSeparated();
		}
	}
}