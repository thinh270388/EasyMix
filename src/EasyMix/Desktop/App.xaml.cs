using Desktop.DependencyInjection;
using Desktop.Helpers;
using Desktop.Services.Implementations;
using Desktop.Services.Interfaces;
using Desktop.ViewModels;
using Desktop.Views;
using EasyUpdater.Core.Models;
using EasyUpdater.Core.ViewModels;
using EasyUpdater.Core.Views;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Reflection;
using System.Windows;

namespace Desktop
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private readonly IServiceProvider _serviceProvider;

        public App()
        {
            var serviceCollection = new ServiceCollection();
            serviceCollection.AddService();
            _serviceProvider = serviceCollection.BuildServiceProvider();
        }
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (await CheckInternet.IsInternetAvailableAsync())
            {
                var versionChecker = _serviceProvider.GetRequiredService<IVersionChecker>();

                // 📦 Kiểm tra & cập nhật EasyUpdater
                var updaterJson = Path.Combine(AppContext.BaseDirectory, "EasyUpdater.json");
                var updaterLocal = await versionChecker.ReadLocalAsync(updaterJson);
                var updaterRemote = await versionChecker.ReadRemoteAsync(updaterLocal?.VersionUrl ?? string.Empty);
                if (updaterLocal != null && updaterRemote != null && updaterLocal.Version != updaterRemote.Version)
                {
                    Console.WriteLine($"🔄 Có bản mới cho {updaterLocal.AppName}: {updaterLocal.Version} → {updaterRemote.Version}");

                    var http = new HttpClient();
                    var zipData = await http.GetByteArrayAsync(updaterRemote.ZipUrl);
                    var zipPath = Path.Combine(Path.GetDirectoryName(updaterJson)!, updaterRemote.File);
                    await File.WriteAllBytesAsync(zipPath, zipData);

                    ZipFile.ExtractToDirectory(zipPath, Path.GetDirectoryName(updaterJson)!, overwriteFiles: true);
                    File.Delete(zipPath);

                    Console.WriteLine($"✅ Đã cập nhật {updaterRemote.AppName}!");
                }

                // 🚀 Kiểm tra & cập nhật ứng dụng chính
                var appJson = Path.Combine(AppContext.BaseDirectory, "Version.json");
                var appLocal = await versionChecker.ReadLocalAsync(appJson);
                var appRemote = await versionChecker.ReadRemoteAsync(appLocal?.VersionUrl ?? string.Empty);
                if (appLocal != null && appRemote != null && appLocal.Version != appRemote.Version)
                {
                    var ctx = new UpdateContext
                    {
                        Url = appRemote.ZipUrl,
                        FileName = appLocal.File,
                        AppExe = Environment.ProcessPath!
                    };

                    var vm = new UpdateViewModel(ctx);
                    var win = new UpdateView(vm);
                    win.ShowDialog();
                    return;
                }
            }

            // 🚪 Khởi động giao diện chính
            ViewTemplateSelector.ViewLocator = _serviceProvider.GetRequiredService<IViewLocator>();
            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow!.DataContext = _serviceProvider.GetRequiredService<MainViewModel>();
            mainWindow.Show();
        }
    }
}
