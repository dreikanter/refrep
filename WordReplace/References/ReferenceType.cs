using System.ComponentModel;

namespace WordReplace.References
{
	/// <summary>
	/// Тип библиографической ссылки, определяющий то, как будет сгенерирован ее текст.
	/// </summary>
    public enum ReferenceType
    {
        Unknown = 0,

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
