using System;
using WordReplace.References;

namespace WordReplace.Attributes
{
	public class MandatoryFieldAttribute : Attribute
	{
		public ReferenceFields[] Fields { get; private set; }

		public MandatoryFieldAttribute(params ReferenceFields[] fields)
		{
			Fields = fields;
		}
	}
}
