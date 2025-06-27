using System.Windows.Controls;

namespace Desktop.Services.Interfaces
{
    public interface IViewLocator
    {
        UserControl GetViewForViewModel(object viewModel);
    }
}
