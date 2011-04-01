using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using WordReplace.Extensions;

namespace WordReplace.References
{
	/// <summary>
	/// Библиографическая ссылка. Объекты этого типа генерируются на основе данных из таблицы Excel.
	/// </summary>
    public class Reference : IComparable
    {
        public int RefNum { get; set; }

		/// <summary>
		/// Excel row number.
		/// </summary>
		public int RowNum { get; private set; }
        
        public int Id { get; private set; }

		public string Tag { get; private set; }

        public ReferenceType Type { get; private set; }
        
        public string Category { get; private set; }

		public int? Relevance { get; private set; }

		public string Authors { get; private set; }

		public string Title { get; private set; }
        
		/// <summary>
		/// Название журнала или сайта
		/// </summary>
		public string Source { get; private set; }

		public int? Issue { get; private set; }

		public string Publisher { get; private set; }

		public int? Edition { get; private set; }

		public int? Pages { get; private set; }

		public int? From { get; private set; }

		public int? To { get; private set; }

		public string Isbn { get; private set; }

		public int? Year { get; private set; }

		public DateTime? SiteVisited { get; private set; }

		public string Country { get; private set; }

		public string City { get; private set; }

		public string Language { get; private set; }

		public int? Volume { get; private set; }

		public string Url { get; private set; }

		public string Description { get; private set; }

		private string SortingTitle
		{
			get
			{
				return (new[] {Authors, Title, Source}).CommaSeparated();
			}
		}

    	public Reference(int rowNum, IEnumerable<ReferenceFields> order, IEnumerable<string> values)
        {
            if (values.IsNullOrEmpty()) throw new ArgumentException("values");
            if (order.IsNullOrEmpty()) throw new ArgumentException("order");

    		RowNum = rowNum;
            var vs = values.ToArray();
            var fs = order.ToArray();
            
            if (vs.Length != fs.Length) throw new InvalidOperationException();

            var fillErrors = new List<ReferenceFields>();

            for (var i = 0; i < vs.Length; i++)
            {
                try
                {
                    SetValue(fs[i], vs[i]);
                }
                catch
                {
                    fillErrors.Add(fs[i]);
                }
            }
            
            if (fillErrors.Any())
            {
                throw new Exception("Error reading following fields: {0}".
                    Fill(fillErrors.Cast<String>().CommaSeparated()));
            }
		}

        private void SetValue(ReferenceFields field, string value)
        {
            switch (field)
            {
            	case ReferenceFields.Unknown:
            		break;

            	case ReferenceFields.Id:
            		Id = Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Tag:
            		Tag = value;
            		break;

            	case ReferenceFields.Type:
            		Type = value.GetEnumValueOrDefault<ReferenceType>();
            		break;

            	case ReferenceFields.Category:
            		Category = value;
            		break;

            	case ReferenceFields.Relevance:
            		Relevance = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Authors:
            		Authors = value;
            		break;

            	case ReferenceFields.Title:
            		Title = value;
            		break;

            	case ReferenceFields.Magazine:
            		Source = value;
            		break;

            	case ReferenceFields.Issue:
            		Issue = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Publisher:
            		Publisher = value;
            		break;

            	case ReferenceFields.Edition:
            		Edition = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Volume:
            		Volume = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Pages:
            		Pages = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.From:
            		From = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.To:
            		To = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.Isbn:
            		Isbn = value;
            		break;

            	case ReferenceFields.Year:
            		Year = value.IsNullOrEmpty() ? (int?) null : Convert.ToInt32(value);
            		break;

            	case ReferenceFields.SiteVisited:
            		SiteVisited = (value.IsNullOrEmpty()) ? (DateTime?) null : DateTime.ParseExact(value, Constants.DateTimeFormat, CultureInfo.InvariantCulture);
            		break;

            	case ReferenceFields.Country:
            		Country = value;
            		break;

            	case ReferenceFields.City:
            		City = value;
            		break;

            	case ReferenceFields.Language:
            		Language = value;
            		break;

            	case ReferenceFields.Url:
            		Url = value.TrimEnd(new[] {'/'});
                    break;

				case ReferenceFields.Description:
					Description = value;
                    break;
                
                default:
                    throw new ArgumentOutOfRangeException("field");
            }
        }

        public int CompareTo(object obj)
        {
            if (!(obj is Reference)) return 1;
            var other = obj as Reference;
            return SortingTitle.CompareTitleStrings(other.SortingTitle);
        }

		public override string ToString()
		{
			return ReferenceStringGenerator.GetReferenceText(this);
		}
    }
}