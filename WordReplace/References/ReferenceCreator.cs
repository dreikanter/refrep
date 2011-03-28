using System;
using System.Collections.Generic;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Генератор текста ссылок для библиографического списка
	/// </summary>
	public class ReferenceCreator
	{
		public static string GetReferenceText(Reference reference)
		{
			ICollection<string> errors;
			if (!ReferenceValidator.Validate(reference, out errors))
			{
				throw new Exception(errors.SemicolonSeparated());
			}

			switch (reference.Type)
			{
				case ReferenceType.Book:
					return GetBookRef(reference);

				case ReferenceType.WebPage:
					return GetWebPageRef(reference);

				case ReferenceType.WebSite:
					return GetWebSiteRef(reference);

				case ReferenceType.Article:
					return GetArticleRef(reference);

				default:
					throw new ArgumentOutOfRangeException();
			}
		}

		private static string GetBookRef(Reference reference)
		{
			var builder = new ReferenceBuilder();

			if (reference.Authors.Defined())
			{
				builder.Append(reference.Authors.GetAuthorsList(false));
				builder.Space();
			}

			builder.Append(reference.Title).Dot().Space();

			var lang = reference.Language;
			var edition = reference.Edition;

			if (edition != null)
			{
				builder.Append(GetEdition(lang, (int) edition)).Space();
			}

			var city = GetCity(reference.City);
			var publisher = reference.Publisher;

			if (city.Defined())
			{
				builder.Append(city);
				if (publisher.Defined()) builder.Colon().NbSp();
			}

			builder.Append(publisher);

			if (reference.Year != null)
			{
				if (city.Defined() || publisher.Defined() || edition.Defined()) builder.Comma();
				builder.Space().Append(reference.Year.ToString());
			}

			builder.Dot();

			if (reference.Volume != null)
			{
				builder.Space().Append(GetVolumeTemplate(lang).Fill(reference.Volume));
			}

			if (reference.From.Defined() && reference.To.Defined())
			{
				builder.Space().Append(GetPagesIntervalTemplate(lang).Fill(reference.From, reference.To)).Dot();
			}
			else if (reference.Pages != null)
			{
				builder.Space().Append(GetPagesTemplate(lang).Fill(reference.Pages)).Dot();
			}

			if (reference.Isbn.Defined())
			{
				builder.Space().Append("ISBN {0}.".Fill(reference.Isbn)).Dot();
			}

			if (reference.Description.Defined())
			{
				builder.Space().Append("({0})".Fill(reference.Description)).Dot();
			}

			return builder.Text;
		}

		/// <summary>
		/// Ефимова Т. Н., Кусакин А. В. Охрана и рациональное использование болот 
		/// в Республике Марий Эл // Проблемы региональной экологии. 2007. № 1. С. 80-86.
		/// </summary>
		private static string GetArticleRef(Reference reference)
		{
			var builder = new ReferenceBuilder();
			var lang = reference.Language;

			if (reference.Authors.Defined())
			{
				builder.Append(reference.Authors.GetAuthorsList(false));
				builder.Space();
			}

			builder.Append(reference.Title).NbSp().Append("//").NbSp();
			builder.Append(reference.Magazine).Dot().Space();
			builder.Append(reference.Year.ToString()).Dot();
			
			if (reference.Issue != null)
			{
				builder.Space().Append(GetIssueNumber(lang, (int)reference.Issue));
			}

			if (reference.From.Defined() && reference.To.Defined())
			{
				builder.Space().Append(GetPagesIntervalTemplate(lang).Fill(reference.From, reference.To)).Dot();
			}
			else if (reference.Pages != null)
			{
				builder.Space().Append(GetPagesTemplate(lang).Fill(reference.Pages)).Dot();
			}

			if (reference.Isbn.Defined())
			{
				builder.Space().Append("ISSN {0}.".Fill(reference.Isbn)).Dot();
			}

			if (reference.Description.Defined())
			{
				builder.Space().Append("({0})".Fill(reference.Description)).Dot();
			}

			return builder.Text;
		}

		/// <summary>
		/// Лэтчфорд Е. У. С Белой армией в Сибири [Электронный ресурс] // Восточный фронт 
		/// армии адмирала А. В. Колчака: [сайт]. [2004]. URL: http://east-front.narod.ru/memo/latchford.htm (дата обращения: 23.08.2007).
		///
		/// Дирина А. И. Право военнослужащих Российской Федерации на свободу ассоциаций // Военное право: 
		/// сетевой журн. 2007. URL: http://www.voennoepravo.ru/node/2149 (дата обращения: 19.09.2007).
		/// </summary>
		private static string GetWebPageRef(Reference reference)
		{
			return null;
		}

		/// <summary>
		/// Список документов «Информационно-справочной системы архивной отрасли» (ИССАО) и ее 
		/// приложения — «Информационной системы архивистов России» (ИСАР) // Консалтинговая группа 
		/// «Термика»: [сайт]. URL: http://www.termika.ru/dou/progr/spisok24.html (дата обращения: 16.11.2007).
		/// 
		/// 78. Лэтчфорд Е. У. С Белой армией в Сибири [Электронный ресурс] // Восточный фронт армии адмирала 
		/// А. В. Колчака: [сайт]. [2004]. URL: http://east-front.narod.ru/memo/latchford.htm (дата обращения: 23.08.2007).
		/// </summary>
		private static string GetWebSiteRef(Reference reference)
		{
			return null;
		}

		private static string GetCity(string cityName)
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

		private static string GetVolumeTemplate(string lang)
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

		private static string GetPagesIntervalTemplate(string lang)
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

		private static string GetPagesTemplate(string lang)
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

		private static string GetEdition(string lang, int edition)
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
				return Constants.Hyphen + "е";
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

		private static string GetIssueNumber(string lang, int number)
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

		private static string GetUrlDate(string lang, DateTime dateTime)
		{
			if (lang.Equals("ru", StringComparison.InvariantCultureIgnoreCase))
			{
				return "(дата обращения: {0})".Fill(dateTime.ToString("dd-MM-yy"));
			}

			if (lang.Equals("en", StringComparison.InvariantCultureIgnoreCase))
			{
				return "(reference date: {0})".Fill(dateTime.ToString("MM-dd-yy"));
			}

			throw new ArgumentException("Bad language value");
		}
	}
}
