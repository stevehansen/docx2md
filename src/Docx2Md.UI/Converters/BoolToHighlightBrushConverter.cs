using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace Docx2Md.UI.Converters;

/// <summary>
/// Converts a boolean IsSelected value to a highlight brush for segment selection.
/// Returns a semi-transparent accent color when true, transparent when false.
/// </summary>
public class BoolToHighlightBrushConverter : IValueConverter
{
    public static readonly BoolToHighlightBrushConverter Instance = new();

    // Semi-transparent blue highlight
    private static readonly IBrush HighlightBrush = new SolidColorBrush(Color.FromArgb(60, 66, 133, 244));
    private static readonly IBrush TransparentBrush = Brushes.Transparent;

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is bool isSelected && isSelected)
        {
            return HighlightBrush;
        }
        return TransparentBrush;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
