using System;
using System.Collections.Generic;
using System.Text;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Хелпер для ReferenceStringGenerator.
	/// </summary>
	public static class ReferenceUtils
	{
		public static string GetCity(string cityName)
		{
			if (cityName.Equals("Москва", StringComparison.InvariantCultureIgnoreCase))
			{
				return "M.";
			}

			if (cityName.Equals("Санкт-Петербург", StringComparison.InvariantCultureIgnoreCase))
			{
				return "СПб.";
			}

			if (cityName.Equals("Ленинград", StringComparison.InvariantCultureIgnoreCase))
			{
				return "Л.";
			}

			return cityName.Capitilize();
		}

		public static string GetVolumeTemplate(string lang)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return "Т." + Constants.NbSp + "{0}.";
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return "Vol." + Constants.NbSp + "{0}.";
			}

			throw new ArgumentException("Bad language value");
		}

		public static string GetPagesIntervalTemplate(string lang)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return "C." + Constants.NbSp + "{0}" + Constants.EnDash + "{1}";
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return "P." + Constants.NbSp + "{0}" + Constants.EnDash + "{1}";
			}

			throw new ArgumentException("Bad language value");
		}

		public static string GetPagesTemplate(string lang)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return "{0}" + Constants.NbSp + "с.";
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return "{0}" + Constants.NbSp + "p.";
			}

			throw new ArgumentException("Bad language value");
		}

		public static string GetEdition(string lang, int edition)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return GetOrdinal(lang, edition) + Constants.NbSp + "изд.";
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return GetOrdinal(lang, edition) + Constants.NbSp + "edn.";
			}

			throw new ArgumentException("Bad language value");
		}

		private static string GetOrdinal(string lang, int edition)
		{
			return String.Concat(edition.ToString(), GetOrdinalSuffix(lang, edition));
		}

		private static string GetOrdinalSuffix(string lang, int edition)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return Constants.NbHyphen + "е";
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				switch (edition % 100)
				{
					case 11:
					case 12:
					case 13:
						return "th";
				}

				switch (edition % 10)
				{
					case 1:
						return "st";
					case 2:
						return "nd";
					case 3:
						return "rd";
					default:
						return "th";
				}
			}

			throw new ArgumentException("Bad language value");
		}

		public static string GetIssueNumber(string lang, int number)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return "№" + Constants.NbSp + number;
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return "#" + number;
			}

			throw new ArgumentException("Bad language value");
		}

		public static string GetUrlDate(DateTime dateTime)
		{
			return "(дата обращения: {0})".Fill(dateTime.ToString(Constants.UrlDateFormat));
		}

		/// <summary>
		/// Returns a shortened list of author names.
		/// </summary>
		public static string GetAuthorsList(string authors, bool initialsInFront, bool useNbSp)
		{
			var result = new List<string>();
			var space = useNbSp ? Constants.NbSp : " ";

			foreach (var author in authors.Split(new[] { Constants.AuthorsDelimiter }, StringSplitOptions.RemoveEmptyEntries))
			{
				var name = new StringBuilder();
				var parts = author.Split(new[] {" ", "\t", "."}, StringSplitOptions.RemoveEmptyEntries);

				if(parts.Length == 1)
				{
					name.Append(parts[0]);
				}
				else
				{
					for (var i = 1; i < parts.Length; i++)
					{
						name.Append(parts[i].GetInitial() + ".");
					}

					if (initialsInFront) name.Append(space + parts[0]); else name.Insert(0, parts[0] + space);
				}

				result.Add(name.ToString());
			}

			return result.CommaSeparated();
		}
	}
}
