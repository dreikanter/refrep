using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordReplace.References
{
	/// <summary>
	/// Класс, готовящий данные для автозамены ссылок в тексте.
	/// </summary>
	/// <remarks>
	/// В тексте могут присутствовать как индивидуальные ссылки ("[[#1]]"), так и группы ("[[#1, Иванов24]]"). 
	/// Сами ссылки могу обозначаться с помощью идентификаторов - уникальных числовых обозначений ссылок 
	/// (например, #1, #24), так и по тегам - так же уникальным строчным значениям (например, Иванов24, Петров2001).
	/// </remarks>
	public class ReferenceReplacer
	{
		public const string RefDelimiter = ",";

		public const string RefIdPrefix = "#";

		private readonly ReferenceCollection _usedReferences;
	
		/// <summary>
		/// Коллекция упомянутых в тексте ссылок — подмножество общей коллекции ссылок.
		/// </summary>
		public ReferenceCollection UsedReferences { get { return _usedReferences; } }

		private readonly Dictionary<string, Reference[]> _replacements;

		/// <summary>
		/// Основной результат работы класса: словарь, в котором каждой ссылке и группе 
		/// ссылок в тексте сопоставлена коллекция соответствующих объектов-ссылок.
		/// </summary>
		public IDictionary<string, Reference[]> Replacements { get { return _replacements; } }

		private readonly HashSet<string> _unknownTags;

		/// <summary>
		/// Коллекция встретившихся в тексте неизвестных тегов ссылок.
		/// </summary>
		public string[] UnknownTags { get { return _unknownTags.ToArray(); } }

		private readonly HashSet<int> _unknownIds;

		/// <summary>
		/// Коллекция встретившихся в тексте неизвестных идентификаторов ссылок.
		/// </summary>
		public int[] UnknownIds { get { return _unknownIds.ToArray(); } }

		private readonly HashSet<string> _badIds;

		/// <summary>
		/// Коллекция синтаксически-некорректных идентификаторов ссылок.
		/// </summary>
		public string[] BadIds { get { return _badIds.ToArray(); } }

		public ReferenceReplacer(string text, ref ReferenceCollection references, ReferenceOrder order)
		{
			_usedReferences = new ReferenceCollection();
			_badIds = new HashSet<string>();
			_unknownIds = new HashSet<int>();
			_unknownTags = new HashSet<string>();
			_replacements = new Dictionary<string, Reference[]>();

			var re = new Regex(@"\[\[(([" + RefIdPrefix + @"\d\w\s" + RefDelimiter + @"]+)?)\]\]", RegexOptions.IgnoreCase);

			foreach(Match match in re.Matches(text))
			{
				var refText = match.Groups[0].Captures[0].Value;
				var refList = match.Groups[1].Captures[0].Value;

				if (Replacements.ContainsKey(refText)) continue;

				var parts = refList.Split(new[] {RefDelimiter}, StringSplitOptions.RemoveEmptyEntries);
				if (parts.Length < 1)
				{
					Replacements.Add(refText, null);
					continue;
				}

				var refs = new ReferenceCollection();
				foreach (var part in parts.Select(part => part.Trim()).Where(p => p.Length != 0))
				{
					Reference reference;
					if (part[0] == '#')
					{
						int id;
						if (!Int32.TryParse(part.Substring(1), NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite, CultureInfo.InvariantCulture, out id))
						{
							_badIds.Add(part);
							continue;
						}

						reference = references.GetById(id);

						if (reference == null)
						{
							_unknownIds.Add(id);
							continue;
						}
					}
					else
					{
						reference = references.GetByTag(part);

						if (reference == null)
						{
							_unknownTags.Add(part);
							continue;
						}
					}

					refs.Add(reference);
					if (!UsedReferences.Contains(reference)) UsedReferences.Add(reference);
				}

				refs.Sort(order);
				Replacements.Add(refText, refs.ToArray());
			}
		}
	}
}
