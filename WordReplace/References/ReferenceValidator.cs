using System;
using System.Collections.Generic;
using WordReplace.Extensions;

namespace WordReplace.References
{
	public class ReferenceValidator
	{
		private List<string> _errors;

		public ICollection<string> Errors { get { return _errors; } }

		public bool Validate(Reference reference)
		{
			_errors = new List<string>();

			switch (reference.Type)
			{
				case ReferenceType.Book:
					return ValidateBook(reference);

				case ReferenceType.Article:
					return ValidateArticle(reference);

				case ReferenceType.WebPage:
					return ValidateWebPage(reference);

				case ReferenceType.WebSite:
					return ValidateWebSite(reference);

				default:
					_errors.Add("Nothing to validate.");
					return false;
			}
		}

		private bool ValidateBook(Reference reference)
		{
			throw new NotImplementedException();
			return _errors.Count > 0;
		}

		private bool ValidateArticle(Reference reference)
		{
			if (!reference.Title.Defined()) _errors.Add("Title not defined");

			if (!reference.Magazine.Defined()) _errors.Add("Magazine name not defined");

			if (!reference.Year.Defined()) _errors.Add("Year not defined");

			if (!reference.Issue.Defined()) _errors.Add("Issue not defined");

			if (!reference.Pages.Defined() && !reference.From.Defined() && !reference.To.Defined())
			{
				_errors.Add("Pages not defined");
			}

			return _errors.Count > 0;
		}

		private bool ValidateWebPage(Reference reference)
		{
			throw new NotImplementedException();
			return _errors.Count > 0;
		}

		private bool ValidateWebSite(Reference reference)
		{
			throw new NotImplementedException();
			return _errors.Count > 0;
		}
	}
}
