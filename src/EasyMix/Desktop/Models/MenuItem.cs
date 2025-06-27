using System.Collections.ObjectModel;

namespace Desktop.Models
{
    public class MenuItem
    {
        public string Icon { get; }
        public string Title { get; }
        public Type? ViewModelType { get; }
        public ObservableCollection<MenuItem> Children { get; }

        public bool HasChildren => Children.Count > 0;

        public MenuItem(string icon, string title, Type? viewModelType = null)
        {
            Icon = icon;
            Title = title;
            ViewModelType = viewModelType;
            Children = new ObservableCollection<MenuItem>();
        }
    }
}
