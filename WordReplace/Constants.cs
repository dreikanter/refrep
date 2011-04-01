namespace WordReplace
{
    public static class Constants
    {
		/// <summary>
		/// Используется при парсинге считываемых из Excel значений
		/// </summary>
        public const string DateTimeFormat = "yyyy-MM-dd";

		/// <summary>
		/// Формат даты, используемый при генерации текста ссылки 
		/// на веб-страницы ("дата обращения: ...")
		/// </summary>
		public const string UrlDateFormat = "dd.MM.yyyy";

		/// <summary>
		/// Разделитель ID ссылок и тегов в группе ссылок ("[#1, #2, tag1]")
		/// </summary>
        public const char RefIdDelimiter = ',';

        public const char AuthorsDelimiter = ',';

		public const string CommaWithSpace = ", ";

		public const string CommaWithNbSp = "," + NbSp;

		public const string SemicolonWithSpace = "; ";

        public const string Ellipsis = "...";

        /// <summary>
        /// Non-breaking space
        /// </summary>
		public const string NbSp = "&nbsp;"; //" ";

		public const string AndOthers = "и" + NbSp + "др."; //"и др.";

        public const string PagesSuffix = NbSp + "с.";

        public const string EnDash = "–";

        public const string EmDash = "—";

		public const string Hyphen = "-";

    	public const string RefListBookmark = "Bibliography";

    	public const string WebResource = "[Электронный ресурс]";
	}
}
