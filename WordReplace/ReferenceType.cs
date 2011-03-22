using System.ComponentModel;

namespace WordReplace
{
    public enum ReferenceType
    {
        Unknown,

        [Description("Book")]
        Book,

        [Description("Web page")]
        WebPage,

        [Description("Web site")]
        WebSite,

        [Description("Article")]
        Article,
    }
}
