using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Desktop.Models;

namespace Desktop.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly IServiceProvider? _serviceProvider;
        public ObservableCollection<MenuItem> Menus { get; }

        [ObservableProperty] private ObservableObject? currentViewModel;
        [ObservableProperty] private bool isMenuExpanded = true;
        [ObservableProperty] public double menuWidth = 180;
        [ObservableProperty] public string appVersion = string.Empty;

        public MainViewModel(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
            Menus = new ObservableCollection<MenuItem>
            {
                new("🏠", "Trang chủ", typeof(HomeViewModel)),
                new("⚙️", "Chuẩn hóa", typeof(NormalizationViewModel)),
                new("🧪", "Trộn đề", typeof(MixViewModel)),
                new("❓", "Hỗ trợ", null)
                {
                    Children =
                    {
                        new("🔄", "Cập nhật", null),
                        new("📞", "Liên hệ", null),
                        new("📘", "Hướng dẫn", null),
                    }
                }
            };

            AppVersion = $"v{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}";
            CurrentViewModel = _serviceProvider?.GetService(typeof(HomeViewModel)) as ObservableObject;
        }

        [RelayCommand]
        private void ToggleMenu()
        {
            IsMenuExpanded = !IsMenuExpanded;
            MenuWidth = IsMenuExpanded ? 180 : 90;
        }

        [RelayCommand]
        private void ChangeView(MenuItem? menu)
        {
            if (menu == null || menu?.ViewModelType == null) return;

            var vm = _serviceProvider!.GetService(menu.ViewModelType) as ObservableObject;
            if (vm != null)
            {
                CurrentViewModel = vm;
            }
        }
    }
}
