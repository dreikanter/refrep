using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordReplace.Extensions
{
	public static class StringExtensions
	{
		public static readonly string Ellipsis = "...";

		public static string CommaSeparated(this IEnumerable<string> list)
		{
			return list.Separated(Constants.CommaWithSpace);
		}

		public static string CommaSeparated(this IEnumerable<string> list, int limit)
		{
			return list.Separated(Constants.CommaWithSpace, limit);
		}

		public static string CommaSeparatedNb(this IEnumerable<string> list)
		{
			return list.Separated(Constants.CommaWithNbSp);
		}

		public static string SemicolonSeparated(this IEnumerable<string> list)
		{
			return list.Separated(Constants.SemicolonWithSpace);
		}

		/// <summary>
		/// Returns separated list of the first N elements from the string collection
		/// </summary>
		public static string Separated(this IEnumerable<string> list, string separator, int limit)
		{
			if (limit < 0) throw new ArgumentOutOfRangeException("limit");
			return (list.Count() <= limit) ? list.Separated(separator) :
				(list.ToList().GetRange(0, limit).Separated(separator) + Constants.Ellipsis);
		}

		/// <summary>
		/// Returns separated list of the collection elements
		/// </summary>
		public static string Separated(this IEnumerable<string> list, string separator)
		{
			return list.IsNullOrEmpty() ? String.Empty :  String.Join(separator, list.ToArray());
		}

		/// <summary>
		/// Parses the DateTime from string value, using format from configuration. 
		/// </summary>
		/// <remarks>
		/// For empty, blank and null strings null will be returned with succeeded = true.
		/// </remarks>
		/// <param name="value">DateTime string representation to be converted.</param>
		/// <param name="format">DateTime format.</param>
		/// <param name="succeeded">Parsing result.</param>
		public static DateTime? ToDateTime(this string value, string format, out bool succeeded)
		{
			if (value.IsNullOrBlank())
			{
				succeeded = true;
				return null;
			}

			DateTime result;

			succeeded = DateTime.TryParseExact(value, format, 
				CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
			
			return succeeded ? (DateTime?) result : null;
		}

		/// <summary>
		/// Convert the DateTime value from string representation, using a set 
		/// of available format alternatives. The first matching format from the 
		/// collection will be used.
		/// </summary>
		public static DateTime? ToDateTime(this string value, IEnumerable<string> availableFormats, out bool succeeded)
		{
			if(availableFormats.IsNullOrEmpty()) throw new ArgumentException("availableFormats");

			foreach(var format in availableFormats)
			{
				var result = value.ToDateTime(format, out succeeded);
				if (succeeded) return result;
			}

			succeeded = false;
			return null;
		}

		public static string Fill(this string format, params object[] args)
		{
			return String.Format(format, args);
		}

		public static string OrDefault(this string value, string defaultValue)
		{
			return String.IsNullOrEmpty(value) ? defaultValue : value;
		}

		public static string Ellip(this string value, int maxLen)
		{
			return value.Ellip(maxLen, Ellipsis);
		}

		public static string Ellip(this string value, int maxLen, string suffix)
		{
			if (maxLen < 0) throw new ArgumentException("maxLen");
			if (suffix == null) throw new ArgumentNullException("suffix");

			if (value.IsNullOrEmpty() || maxLen == 0 || value.Length <= maxLen) return value;

			return (value.Length <= suffix.Length) ? value.Substring(0, maxLen) : 
				(value.Substring(0, maxLen - suffix.Length) + suffix);
		}

        public static bool IsNullOrBlank(this string value)
        {
            return String.IsNullOrEmpty(value) || value.Trim().Length == 0;
        }

        public static string JoinWith(this IEnumerable<string> parts, string delimiter)
        {
            if (parts == null) throw new ArgumentNullException("parts");
            if (delimiter == null) throw new ArgumentNullException("delimiter");

            return String.Join(delimiter, parts.ToArray());
        }

        public static string Capitilize(this string value)
        {
            if (value.IsNullOrBlank()) return value;
            return value.First().ToString().ToUpper() + value.Substring(1).ToLower();
        }

		public static string GetInitial(this string name)
		{
			if (name.IsNullOrBlank()) return String.Empty;
			return name.ToLower().StartsWith("дж") ? "Дж" : name.Substring(0, 1).ToUpper();
		}

		public static int CompareTitleStrings(this string x, string y)
		{
			var xEmpty = x.IsNullOrEmpty();
			var yEmpty = y.IsNullOrEmpty();
			if (xEmpty && yEmpty) return 0;
			if (xEmpty) return 1;
			if (yEmpty) return -1;

			var ruRe = new Regex(@"[а-я]+", RegexOptions.IgnoreCase);

			var xIsRu = ruRe.IsMatch(x);
			var yIsRu = ruRe.IsMatch(y);

			if ((xIsRu && yIsRu) || !(xIsRu || yIsRu)) return x.CompareTo(y);

			return xIsRu ? 1 : -1;
		}

		public static string InBrackets(this string value)
		{
			return String.Format("[{0}]", value);
		}

		public static string InBraces(this string value)
		{
			return String.Format(@"\{{0}\}", value);
		}

		public static string InParentheses(this string value)
		{
			return String.Format("({0})", value);
		}
	}
}
