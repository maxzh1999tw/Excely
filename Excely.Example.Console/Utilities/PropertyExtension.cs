using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace Excely.Example.Console.Utilities
{
    internal static class PropertyExtension
    {
        public static string GetDisplayName(this PropertyInfo propertyInfo)
        {
            var displayAttribute = propertyInfo.GetCustomAttribute<DisplayAttribute>();
            return displayAttribute?.Name ?? propertyInfo.Name;
        }
    }
}
