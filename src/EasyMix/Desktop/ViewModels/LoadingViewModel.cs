using CommunityToolkit.Mvvm.ComponentModel;

namespace Desktop.ViewModels
{
    public partial class LoadingViewModel : ObservableObject
    {
        [ObservableProperty] private string message = "Loading...";
    }
}
