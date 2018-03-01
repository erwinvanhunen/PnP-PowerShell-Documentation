using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Generate.Model
{
    internal static class Extensions
    {
        public static T GetAttributeValue<T>(this CustomAttributeData attribute, string name)
        {
            var argument = attribute.NamedArguments.FirstOrDefault(p => p.MemberName == name);
            if (argument != null && argument.TypedValue != null)
            {
                if (argument.TypedValue.Value != null)
                {
                    return (T)argument.TypedValue.Value;
                }
            }
            return default(T);
        }
    }
}
