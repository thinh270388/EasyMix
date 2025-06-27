using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace Desktop.Helpers
{
    public class MessageHelper
    {
        public static MessageBoxResult Success(string message, string title = "Thông báo", MessageBoxImage icon = MessageBoxImage.Information)
        {
            return MessageBox.Show(message, title, MessageBoxButton.OK, icon);
        }

        public static MessageBoxResult Error(string message, string title = "Lỗi", MessageBoxImage icon = MessageBoxImage.Error)
        {
            return MessageBox.Show("Thất bại!\nLỗi: " + message, title, MessageBoxButton.OK, icon);
        }

        public static MessageBoxResult Error(Exception ex, string title = "Lỗi", MessageBoxImage icon = MessageBoxImage.Error)
        {
            return MessageBox.Show("Thất bại!\nLỗi: " + ex.Message, title, MessageBoxButton.OK, icon);
        }
    }
}
