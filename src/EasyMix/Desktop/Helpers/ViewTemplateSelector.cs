using Desktop.Services.Interfaces;
using System.Windows;
using System.Windows.Controls;

namespace Desktop.Helpers
{
    public class ViewTemplateSelector : DataTemplateSelector
    {
        public static IViewLocator? ViewLocator { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item == null || ViewLocator == null) return null!;

            var view = ViewLocator.GetViewForViewModel(item);

            // Tạo một DataTemplate mới cho View
            var template = new DataTemplate();
            var factory = new FrameworkElementFactory(view.GetType());
            template.VisualTree = factory;
            return template;
        }
    }
}
