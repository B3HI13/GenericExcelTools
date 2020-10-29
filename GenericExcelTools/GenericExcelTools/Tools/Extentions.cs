using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace GenericExcelTools
{
    public static class Extentions
    {
        public static string GetPropertyDisplayName(this Type modelType, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) return null;

            var displayAttribute = modelType
                .GetProperties()
                .FirstOrDefault(q => q.Name == propertyName)
                .GetCustomAttributes(typeof(DisplayAttribute), true)
                .FirstOrDefault() as DisplayAttribute;

            return displayAttribute.Name;
        }

        public static bool HaveNumbers(this string input)
            => input.Any(char.IsDigit);

        public static string ToStandardNumbers(this string input)
        {
            string[] persian = new string[10] { "۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹" };

            for (int j = 0; j < persian.Length; j++)
                input = input.Replace(persian[j], j.ToString());

            return input;
        }
    }
}
