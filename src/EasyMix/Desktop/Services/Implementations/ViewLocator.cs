using Desktop.Services.Interfaces;
using System.Windows;
using System.Windows.Controls;

namespace Desktop.Services.Implementations
{
    public class ViewLocator : IViewLocator
    {
        private readonly IServiceProvider _serviceProvider;

        public ViewLocator(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }
        public UserControl GetViewForViewModel(object viewModel)
        {
            if (viewModel == null)
                return WrapText("Null ViewModel");

            var viewModelType = viewModel.GetType();
            var viewTypeName = viewModelType.FullName!.Replace("ViewModel", "View");
            var viewAssembly = viewModelType.Assembly;

            var viewType = viewAssembly.GetType(viewTypeName);
            if (viewType == null)
                return WrapText($"Không tìm thấy View cho {viewModelType.Name}");

            var view = _serviceProvider.GetService(viewType) as UserControl;
            if (view == null)
                return WrapText($"Không resolve được {viewType.Name} từ DI");

            view.DataContext = viewModel;
            return view;
        }
        private UserControl WrapText(string message)
        {
            return new UserControl
            {
                Content = new TextBlock
                {
                    Text = message,
                    Foreground = System.Windows.Media.Brushes.Red,
                    Margin = new Thickness(10)
                }
            };
        }
    }
}
