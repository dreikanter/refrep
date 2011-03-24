using System;

namespace WordReplace
{
	public class TemplateAttribute : Attribute
	{
		public string Template { get; private set; }

		public TemplateAttribute(string template)
		{
			Template = template;
		}
	}
}
