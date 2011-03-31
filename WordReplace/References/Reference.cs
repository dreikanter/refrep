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
        
        public int Id { get; set; }

		public string Tag { get; set; }

        public ReferenceType Type { get; set; }
        
        public string Category { get; set; }
        
        public int? Relevance { get; set; }
        
        public string Authors { get; set; }
        
        public string Title { get; set; }
        
		/// <summary>
		/// Название журнала или сайта
		/// </summary>
        public string Source { get; set; }
        
        public int? Issue { get; set; }
        
        public string Publisher { get; set; }
        
        public int? Edition { get; set; }
        
        public int? Pages { get; set; }
        
        public int? From { get; set; }
        
        public int? To { get; set; }
        
        public string Isbn { get; set; }
        
        public int? Year { get; set; }
        
        public DateTime? SiteVisited { get; set; }
        
        public string Country { get; set; }
        
        public string City { get; set; }
        
        public string Language { get; set; }

		public int? Volume { get; set; }

        public string Url { get; set; }

		public string Description { get; set; }

		private string SortingTitle { get { return (Type == ReferenceType.Book) ? Authors : Title; } }

		public Reference() { }

    	public Reference(IEnumerable<ReferenceFields> order, IEnumerable<string> values)
        {
            if (values.IsNullOrEmpty()) throw new ArgumentException("values");
            if (order.IsNullOrEmpty()) throw new ArgumentException("order");

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
                    SiteVisited = (value.IsNullOrEmpty()) ? (DateTime?) null : 
                        DateTime.ParseExact(value, Constants.DateTimeFormat, CultureInfo.InvariantCulture);
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
                    Url = value;
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
			return ReferenceCreator.GetReferenceText(this);
		}
    }
}