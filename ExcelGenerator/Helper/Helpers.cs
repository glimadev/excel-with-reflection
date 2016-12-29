using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ExcelGenerator.Helper
{
    public static class Helpers
    {
        public static string GetDisplayName(PropertyInfo prop)
        {
            if (prop.CustomAttributes == null || prop.CustomAttributes.Count() == 0)
                return prop.Name;

            var displayNameAttribute = prop.CustomAttributes.Where(x => x.AttributeType == typeof(DisplayNameAttribute)).FirstOrDefault();

            if (displayNameAttribute == null || displayNameAttribute.ConstructorArguments == null || displayNameAttribute.ConstructorArguments.Count == 0)
                return prop.Name;

            return displayNameAttribute.ConstructorArguments[0].Value.ToString() ?? prop.Name;
        }

        public static bool HiddenAttribute(PropertyInfo prop)
        {
            if (prop.CustomAttributes == null || prop.CustomAttributes.Count() == 0)
                return false;

            var hiddenAttribute = prop.CustomAttributes.Where(x => x.AttributeType == typeof(DesignerSerializationVisibilityAttribute)).FirstOrDefault();

            if (hiddenAttribute != null)
            {
                return !Convert.ToBoolean(hiddenAttribute.ConstructorArguments[0].Value);
            }

            return false;
        }
    }
}
