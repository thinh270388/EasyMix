using Desktop.Services.Implementations;
using Desktop.Services.Interfaces;
using Desktop.ViewModels;
using Desktop.Views;
using Microsoft.Extensions.DependencyInjection;

namespace Desktop.DependencyInjection
{
    public static class ServiceContainer
    {
        public static IServiceCollection AddService(this IServiceCollection services)
        {
            // Đăng ký Services
            services.AddSingleton<IViewLocator, ViewLocator>();
            services.AddTransient<IExcelAnswerExporter, ExcelAnswerExporter>();
            services.AddTransient<IOpenXMLService, OpenXMLService>();
            services.AddTransient<IInteropWordService, InteropWordService>();
            services.AddTransient<IVersionChecker, VersionChecker>();

            // Đăng ký Views
            services.AddSingleton<HomeView>();
            services.AddTransient<MainWindow>();
            services.AddTransient<NormalizationView>();
            services.AddTransient<MixView>();

            // Đăng ký ViewModels
            services.AddTransient<HomeViewModel>();
            services.AddSingleton<MainViewModel>();
            services.AddTransient<NormalizationViewModel>();
            services.AddTransient<MixViewModel>();

            return services;
        }
    }
}
