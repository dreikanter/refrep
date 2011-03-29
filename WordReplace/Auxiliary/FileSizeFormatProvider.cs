using System;

namespace WordReplace.Auxiliary
{
	/// <summary>
	/// Format file size with GB/MB/kB/B suffixes.
	/// </summary>
	/// <remarks>
	/// Credits for http://flimflan.com/blog/FileSizeFormatProvider.aspx
	///</remarks>
	public class FileSizeFormatProvider : IFormatProvider, ICustomFormatter
	{
		private const string _fileSizeFormat = "fs";
		private const Decimal _oneKiloByte = 1024M;
		private const Decimal _oneMegaByte = _oneKiloByte * 1024M;
		private const Decimal _oneGigaByte = _oneMegaByte * 1024M;

		public object GetFormat(Type formatType)
		{
			return formatType == typeof(ICustomFormatter) ? this : null;
		}

		public string Format(string format, object arg, IFormatProvider formatProvider)
		{
			if (format == null || !format.StartsWith(_fileSizeFormat))
			{
				return DefaultFormat(format, arg, formatProvider);
			}

			if (arg is string)
			{
				return DefaultFormat(format, arg, formatProvider);
			}

			Decimal size;

			try
			{
				size = Convert.ToDecimal(arg);
			}
			catch (InvalidCastException)
			{
				return DefaultFormat(format, arg, formatProvider);
			}

			string suffix;
			if (size > _oneGigaByte)
			{
				size /= _oneGigaByte;
				suffix = "GB";
			}
			else if (size > _oneMegaByte)
			{
				size /= _oneMegaByte;
				suffix = "MB";
			}
			else if (size > _oneKiloByte)
			{
				size /= _oneKiloByte;
				suffix = "kB";
			}
			else
			{
				suffix = "B";
			}

			var precision = (size < _oneKiloByte) ? 0 : 1;
			return String.Format("{0:N" + precision + "} {1}", size, suffix);
		}

		private static string DefaultFormat(string format, object arg, IFormatProvider formatProvider)
		{
			var tmp = arg as IFormattable;
			return tmp != null ? tmp.ToString(format, formatProvider) : arg.ToString();
		}
	}
}