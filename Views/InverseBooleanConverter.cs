using System;
using System.Globalization;
using System.Windows.Data;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Inverts a boolean value. Used for the OR radio button binding.
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : value;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b ? !b : value;
    }
}
