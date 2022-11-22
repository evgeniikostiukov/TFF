using System;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace tff.main.Extensions;

public static class DisplayExtension
{
    public static string GetDisplayName(this PropertyInfo pi)
    {
        if (pi == null)
        {
            throw new ArgumentNullException(nameof(pi));
        }

        return pi.IsDefined(typeof(DisplayAttribute)) ? pi.GetCustomAttribute<DisplayAttribute>()?.GetName() : pi.Name;
    }

    public static string GetDisplayName<T>(this T _, string propertyName)
    {
        return typeof(T).GetProperty(propertyName).GetDisplayName();
    }
}