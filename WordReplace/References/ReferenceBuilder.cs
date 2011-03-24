using System.Text;

namespace WordReplace.References
{
	public class ReferenceBuilder
	{
		private readonly StringBuilder _builder;

		public string Text { get { return _builder.ToString(); } }

		public ReferenceBuilder()
		{
			_builder = new StringBuilder();
		}

		public ReferenceBuilder Append(string text)
		{
			_builder.Append(text);
			return this;
		}

		public ReferenceBuilder NbSp()
		{
			return Append(Constants.NbSp);
		}

		public ReferenceBuilder Space()
		{
			return Append(" ");
		}

		public ReferenceBuilder Dot()
		{
			return Append(".");
		}

		public ReferenceBuilder EnDash()
		{
			return Append("‏–");
		}

		public ReferenceBuilder EmDash()
		{
			return Append("‏—");
		}

		public ReferenceBuilder Colon()
		{
			return Append(":");
		}

		public ReferenceBuilder Semicolon()
		{
			return Append(";");
		}

		public ReferenceBuilder Comma()
		{
			return Append(",");
		}

		public override string ToString()
		{
			return Text;
		}
	}
}
