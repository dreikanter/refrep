using System.ComponentModel;

namespace WordReplace.References
{
	/// <summary>
	/// Порядок сортировки списка библиографических ссылок.
	/// </summary>
    public enum ReferenceOrder
    {
        [Description("Mention")]
        Mention = 0,

        [Description("Alpha")]
        Alpha,
    }
}
