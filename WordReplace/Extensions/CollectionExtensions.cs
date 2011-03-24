using System;
using System.Collections.Generic;
using System.Linq;

namespace WordReplace.Extensions
{
	public static class CollectionExtensions
	{
		public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
		{
			TValue result;
			return dictionary.TryGetValue(key, out result) ? result : defaultValue;
		}

		public static void Update<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue value)
		{
			if (dictionary == null) throw new NullReferenceException();

			if (dictionary.ContainsKey(key))
			{
				dictionary[key] = value;
			}
			else
			{
				dictionary.Add(key, value);
			}
		}

		public static bool IsNullOrEmpty<T>(this IEnumerable<T> collection)
		{
			return collection == null || !collection.Any();
		}
	}
}
