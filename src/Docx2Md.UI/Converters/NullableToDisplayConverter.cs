using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Docx2Md.Core.Models;

namespace Docx2Md.UI.Converters;

/// <summary>
/// Converts nullable int (heading level) to display string.
/// null displays as "(Original)", numbers display as "H1", "H2", etc.
/// </summary>
public class HeadingLevelDisplayConverter : IValueConverter
{
    public static readonly HeadingLevelDisplayConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is int level)
        {
            return $"H{level}";
        }
        return "(Original)";
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string str)
        {
            if (str == "(Original)" || string.IsNullOrEmpty(str))
                return null;
            if (str.StartsWith("H") && int.TryParse(str.Substring(1), out var level))
                return level;
        }
        return null;
    }
}

/// <summary>
/// Converts nullable SegmentType to display string.
/// null displays as "(Original)", otherwise displays type name.
/// </summary>
public class SegmentTypeDisplayConverter : IValueConverter
{
    public static readonly SegmentTypeDisplayConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is SegmentType type)
        {
            return type.ToString();
        }
        return "(Original)";
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is string str)
        {
            if (str == "(Original)" || string.IsNullOrEmpty(str))
                return null;
            if (Enum.TryParse<SegmentType>(str, out var type))
                return type;
        }
        return null;
    }
}
