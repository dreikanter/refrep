using System;
using System.Collections.Generic;
using System.Linq;
using WordReplace.Extensions;

namespace WordReplace.References
{
	public class ReferenceCollection : List<Reference>
	{
		public Reference GetByTag(string tag)
		{
			return this.FirstOrDefault(r => r.Tag == tag);
		}

		public Reference GetById(int id)
		{
			return this.FirstOrDefault(r => r.Id == id);
		}

		public bool ContainsId(int id)
		{
			return this.Any(r => r.Id == id);
		}

		public bool ContainsTag(int id)
		{
			return this.Any(r => r.Id == id);
		}

		/// <summary>
		/// Enumerate references according to the order the are stored in the collection.
		/// </summary>
		public void Enumerate()
		{
			for (var i = 0; i < Count; i++) this[i].RefNum = i + 1;
		}

		public void Sort(ReferenceOrder order)
		{
			switch(order)
			{
				case ReferenceOrder.Mention:
					Sort((x, y) => x.Id - y.Id);
					break;

				case ReferenceOrder.Alpha:
					Sort((x, y) => x.SortingTitle.CompareTitleStrings(y.SortingTitle));
					break;

				default:
					throw new ArgumentOutOfRangeException("order");
			}
		}
	}
}