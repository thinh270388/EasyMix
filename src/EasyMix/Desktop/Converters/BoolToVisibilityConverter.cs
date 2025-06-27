using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Desktop.Converters
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public bool CollapseWhenFalse { get; set; } = false;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Visible : (CollapseWhenFalse ? Visibility.Collapsed : Visibility.Hidden);
            }
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is Visibility visibility && visibility == Visibility.Visible;
        }
    }
}
