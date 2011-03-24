using System.ComponentModel;

namespace WordReplace.References
{
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
        
        [Description("Authors")]
        Authors,
        
        [Description("Title")]
        Title,
        
        [Description("Magazine")]
        Magazine,
        
        [Description("Issue")]
        Issue,
        
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
