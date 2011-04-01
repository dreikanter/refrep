using System;
using System.Collections.Generic;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Генератор текста ссылок для библиографического списка (хелпер для класса Reference).
	/// </summary>
	public static class ReferenceStringGenerator
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
					return ReferenceUtils.GetWebSiteRef(reference);

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
				builder.Append("<i>{0}</i>".Fill(ReferenceUtils.GetAuthorsList(reference.Authors, false, true))).Space();
			}

			builder.Append(reference.Title).Dot().Space();

			var lang = reference.Language;
			var edition = reference.Edition;

			if (edition != null)
			{
				builder.Append(ReferenceUtils.GetEdition(lang, (int)edition)).Space();
			}

			var city = ReferenceUtils.GetCity(reference.City);
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
				builder.Space().Append(ReferenceUtils.GetVolumeTemplate(lang).Fill(reference.Volume));
			}

			if (reference.From.Defined() && reference.To.Defined())
			{
				builder.Space().Append(ReferenceUtils.GetPagesIntervalTemplate(lang).Fill(reference.From, reference.To)).Dot();
			}
			else if (reference.Pages != null)
			{
				builder.Space().Append(ReferenceUtils.GetPagesTemplate(lang).Fill(reference.Pages));
			}

			if (reference.Isbn.Defined())
			{
				builder.Space().Append("ISBN ").Append(reference.Isbn).Dot();
			}

			if (reference.Description.Defined())
			{
				builder.Space().Append(reference.Description.InParentheses()).Dot();
			}

			return builder.Text;
		}

		/// <example>
		/// Ефимова Т. Н., Кусакин А. В. Охрана и рациональное использование болот 
		/// в Республике Марий Эл // Проблемы региональной экологии. 2007. № 1. С. 80-86.
		/// </example>
		private static string GetArticleRef(Reference reference)
		{
			var builder = new ReferenceBuilder();
			var lang = reference.Language;

			if (reference.Authors.Defined())
			{
				builder.Append("<i>{0}</i>".Fill(ReferenceUtils.GetAuthorsList(reference.Authors, false, true))).Space();
			}

			builder.Append(reference.Title).NbSp().Append("//").NbSp();
			builder.Append(reference.Source).Dot().Space();
			builder.Append(reference.Year.ToString()).Dot();
			
			if (reference.Issue != null)
			{
				builder.Space().Append(ReferenceUtils.GetIssueNumber(lang, (int)reference.Issue));
			}

			if (reference.From.Defined() && reference.To.Defined())
			{
				builder.Space().Append(ReferenceUtils.GetPagesIntervalTemplate(lang).Fill(reference.From, reference.To)).Dot();
			}
			else if (reference.Pages != null)
			{
				builder.Space().Append(ReferenceUtils.GetPagesTemplate(lang).Fill(reference.Pages));
			}

			if (reference.Isbn.Defined())
			{
				builder.Space().Append("ISSN ").Append(reference.Isbn).Dot();
			}

			if (reference.Description.Defined())
			{
				builder.Space().Append(reference.Description.InParentheses()).Dot();
			}

			return builder.Text;
		}

		/// <example>
		/// Лэтчфорд Е. У. С Белой армией в Сибири [Электронный ресурс] // Восточный фронт 
		/// армии адмирала А. В. Колчака: [сайт]. [2004]. URL: http://east-front.narod.ru/memo/latchford.htm 
		/// (дата обращения: 23.08.2007).
		///
		/// Дирина А. И. Право военнослужащих Российской Федерации на свободу ассоциаций // Военное право: 
		/// сетевой журн. 2007. URL: http://www.voennoepravo.ru/node/2149 (дата обращения: 19.09.2007).
		/// </example>
		private static string GetWebPageRef(Reference reference)
		{
			var builder = new ReferenceBuilder();
			
			if(!reference.Authors.IsNullOrBlank())
			{
				builder.Append("<i>{0}</i>".Fill(ReferenceUtils.GetAuthorsList(reference.Authors, false, true))).Space();
			}

			builder.Append(reference.Title);

			builder.Append(Constants.WebResource).Space();

			if(!reference.Source.IsNullOrBlank())
			{
				builder.Append(" // " + reference.Source).Dot().Space();
				if(reference.Year != null)
				{
					builder.Append(": [сайт]. {0}. ".Fill(reference.Year));
				}
			}

			builder.Append("URL: " + reference.Url);

			if(reference.SiteVisited != null)
			{
				builder.Space().Append(ReferenceUtils.GetUrlDate((DateTime)reference.SiteVisited)).Dot();
			}

			return builder.Text;
		}
	}
}
