using Microsoft.Extensions.Logging;
using SharePointAuthApp.Helpers;
using SharePointAuthApp.Models;
using SharePointAuthApp.Services;
using SharePointAuthApp.ViewModels;
using SharePointAuthApp.Views;
using System.Text.Json;

namespace SharePointAuthApp
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>();

#if DEBUG
    		builder.Logging.AddDebug();
#endif

            // 1. Load configuration settings synchronously for DI registration
            AppConfig config;
            try
            {
                using var stream = FileSystem.OpenAppPackageFileAsync("appsettings.json").GetAwaiter().GetResult();
                using var reader = new StreamReader(stream);
                var json = reader.ReadToEnd();
                config = JsonSerializer.Deserialize<AppConfig>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new AppConfig();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading config in MauiProgram: {ex.Message}");
                config = new AppConfig(); // Fallback empty config
            }

            // 2. Register Singleton Configuration
            builder.Services.AddSingleton(config);

            // 3. Register Core Services
            builder.Services.AddSingleton<IAuthenticationService, AuthenticationService>();
            builder.Services.AddSingleton<ISharePointService, SharePointService>();
            builder.Services.AddSingleton<INavigationService, NavigationService>();

            // 4. Register ViewModels
            builder.Services.AddTransient<LoginViewModel>();
            builder.Services.AddTransient<HomeViewModel>();

            // 5. Register Views (Pages)
            builder.Services.AddTransient<LoginPage>();
            builder.Services.AddTransient<HomePage>();

            return builder.Build();
        }
    }
}
