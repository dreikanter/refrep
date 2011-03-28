using System.ComponentModel;

namespace WordReplace.References
{
	/// <summary>
	/// Полный список полей таблицы с библиографическими данными
	/// </summary>
    public enum ReferenceFields
    {
        Unknown = 0,

        [Description("ID")]
        Id,
        
        [Description("Type")]
        Type,
        
        [Description("Category")]
        Category,
        
        [Description("Relevance")]
        Relevance,
        
		/// <summary>
		/// Список имён авторов, разделённых запятыми.
		/// Формат: "Фамилия Имя Отчество" или "И.О. Фамилия"/ или "И. Фамилия".
		/// </summary>
        [Description("Authors")]
        Authors,
        
		/// <summary>
		/// Заголовок книги, статьи или веб-сайта.
		/// </summary>
        [Description("Title")]
        Title,
        
		/// <summary>
		/// Название периодического издания.
		/// </summary>
        [Description("Magazine")]
        Magazine,
        
		/// <summary>
		/// Номер выпуска периодического издания.
		/// </summary>
        [Description("Issue")]
        Issue,

        /// <summary>
        /// Издательство.
        /// </summary>
        [Description("Publisher")]
        Publisher,
        
        [Description("Edition")]
        Edition,
        
        [Description("Volume")]
        Volume,
        
        [Description("Pages")]
        Pages,
        
        [Description("From")]
        From,
        
        [Description("To")]
        To,

        [Description("ISBN/ISSN")]
        Isbn,
        
        [Description("Year")]
        Year,
        
        [Description("Site Visited")]
        SiteVisited,
        
        [Description("Country")]
        Country,
        
        [Description("City")]
        City,
        
        [Description("Language")]
        Language,
        
        [Description("Url")]
        Url,

		[Description("Description")]
		Description,
    }
}
