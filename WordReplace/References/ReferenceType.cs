using System.ComponentModel;

namespace WordReplace.References
{
    public enum ReferenceType
    {
        Unknown,

        [Description("Book")]
        Book,

        [Description("Article")]
        Article,

        [Description("Web page")]
        WebPage,

        [Description("Web site")]
        WebSite,
    }
}
