using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace WordReplace.Extensions
{
    public static class EnumExtensions
    {
		/// <summary>
		/// Cache for GetEnumDescriptionsMap()
		/// </summary>
		private static Type _lastEnumType;

		/// <summary>
		/// Cache for GetEnumDescriptionsMap()
		/// </summary>
    	private static object _lastMap;

		/// <summary>
		/// Generates { enum value => description } dictionary
		/// </summary>
        public static Dictionary<T, string> GetEnumDescriptionsMap<T>()
        {
			if (_lastEnumType == typeof(T)) return _lastMap as Dictionary<T, string>;

			_lastEnumType = typeof(T);
			_lastMap = Enum.GetValues(_lastEnumType).
				Cast<T>().ToDictionary(value => value, value => GetDescription(value));

			return _lastMap as Dictionary<T, string>;
        }

		/// <summary>
		/// Extension method to extract Description attribute 
		/// value from enumeration type member
		/// </summary>
		public static string ToDescription(this Enum value)
		{
			if (value == null) return null;
			return GetDescription(value, value.GetType()) ?? value.GetName();
		}

		private static string GetDescription<T>(T enumValue)
        {
            var type = typeof (T);
            var memInfo = type.GetMember(enumValue.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {
                var attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                {
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return null;
        }

        private static string GetDescription(Enum enumValue, Type type)
        {
            var memInfo = type.GetMember(enumValue.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {
                var attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                {
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return null;
        }

		public static T GetEnumValue<T>(this string description) where T : new()
		{
			return description.GetEnumValue<T>(false);
		}

		public static T GetEnumValue<T>(this string description, bool ignoreCase) where T : new()
		{
			T result;

			if (description.TryGetEnumValue(ignoreCase, out result))
			{
				return result;
			}

			throw new ArgumentException();
		}

		public static T GetEnumValueOrDefault<T>(this string description) where T : new()
		{
			return description.GetEnumValueOrDefault<T>(true);
		}

    	public static T GetEnumValueOrDefault<T>(this string description, bool ignoreCase) where T : new()
		{
			T result;
			description.TryGetEnumValue(ignoreCase, out result);
			return result;
		}

		public static bool TryGetEnumValue<T>(this string description, bool ignoreCase, out T result) where T : new()
		{
			if (!typeof(T).IsEnum)
			{
				throw new InvalidOperationException("Function can be used with Enum types only");
			}

			foreach (var pair in GetEnumDescriptionsMap<T>())
			{
				var ic = ignoreCase ? StringComparison.InvariantCultureIgnoreCase : StringComparison.InvariantCulture;
				if (String.Equals(pair.Value, description, ic))
				{
					result = pair.Key;
					return true;
				}
			}

			result = new T();
			return false;
		}

		public static string GetName(this Enum value)
		{
			return Enum.GetName(value.GetType(), value);
		}

		public static IEnumerable<Enum> GetFlags(this Enum value)
		{
			var intValue = Convert.ToInt32(value);

			if (Enum.IsDefined(value.GetType(), 0) && intValue == 0)
			{
				return new[] { value };
			}

			return Enum.GetValues(value.GetType()).Cast<Enum>().
				Where(flag => (Convert.ToInt32(flag) & intValue) != 0);
		}

		public static IEnumerable<string> GetFlagNames(this Enum value)
		{
			return value.GetFlags().Select(f => f.ToDescription());
		}
    }
}
