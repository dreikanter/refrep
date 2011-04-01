using System.Collections.Generic;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Валидатор целостности данных для объектов Reference (хелпер Reference).
	/// </summary>
	/// <remarks>
	/// По сути есть всего одна проверка: наличие того или иного поля (плюс более сложное 
	/// исключение для Pages/From+To). Можно это реализовать с помощью атрибутов и сделать 
	/// обобщенный валидатор.
	/// </remarks>
	public static class ReferenceValidator
	{
		/// <summary>
		/// Метод, который проверяет достаточность данных в объекте Reference для того или иного
		/// типа библиографической ссылки.
		/// </summary>
		/// <returns>
		/// Возвращает true или false, в зависимости от валидности проверяемого объекта. В последнем 
		/// случае список найденных ошибок сохраняется в коллекции validationErrors.
		/// </returns>
		public static bool Validate(Reference reference, out ICollection<string> validationErrors)
		{
			validationErrors = new List<string>();

			switch (reference.Type)
			{
				case ReferenceType.Book:
					return ValidateBook(reference, ref validationErrors);

				case ReferenceType.Article:
					return ValidateArticle(reference, ref validationErrors);

				case ReferenceType.WebPage:
					return ValidateWebPage(reference, ref validationErrors);

				case ReferenceType.WebSite:
					return ValidateWebSite(reference, ref validationErrors);

				default:
					validationErrors.Add("Type not defined or has incorrect value");
					return false;
			}
		}

		private static bool ValidateBook(Reference reference, ref ICollection<string> errors)
		{
			if (!reference.Title.Defined()) errors.Add("Title not defined");

			if (!reference.Publisher.Defined()) errors.Add("Publisher not defined");

			if (!reference.Year.Defined()) errors.Add("Year not defined");

			if (!reference.Pages.Defined() && !(reference.From.Defined() && reference.To.Defined()))
			{
				errors.Add("Pages number (or interval) not defined");
			}

			return errors.Count == 0;
		}

		private static bool ValidateArticle(Reference reference, ref ICollection<string> errors)
		{
			if (!reference.Title.Defined()) errors.Add("Title not defined");

			if (!reference.Source.Defined()) errors.Add("Magazine name not defined");

			if (!reference.Year.Defined()) errors.Add("Year not defined");

			if (!reference.Issue.Defined()) errors.Add("Issue not defined");

			if (!reference.Pages.Defined() && !(reference.From.Defined() && reference.To.Defined()))
			{
				errors.Add("Pages number (or interval) not defined");
			}

			return errors.Count == 0;
		}

		private static bool ValidateWebPage(Reference reference, ref ICollection<string> errors)
		{
			if (!reference.Title.Defined()) errors.Add("Title not defined");

			if (!reference.Url.Defined()) errors.Add("URL not defined");

			if (!reference.SiteVisited.Defined()) errors.Add("URL visiting time not defined");

			return errors.Count == 0;
		}

		private static bool ValidateWebSite(Reference reference, ref ICollection<string> errors)
		{
			if (!reference.Title.Defined()) errors.Add("Title not defined");

			if (!reference.Url.Defined()) errors.Add("URL not defined");

			if (!reference.SiteVisited.Defined()) errors.Add("URL visiting time  not defined");

			return errors.Count == 0;
		}
	}
}
