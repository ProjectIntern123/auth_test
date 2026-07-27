using System.Text.Json;
using SharePointAuthApp.Models;

namespace SharePointAuthApp.Helpers
{
    /// <summary>
    /// Helper class to load configuration settings from embedded/raw appsettings.json in MAUI
    /// </summary>
    public static class ConfigHelper
    {
        private static AppConfig? _config;

        /// <summary>
        /// Loads AppConfig asynchronously from Maui assets (appsettings.json)
        /// </summary>
        public static async Task<AppConfig> LoadConfigAsync()
        {
            if (_config != null)
                return _config;

            try
            {
                using var stream = await FileSystem.OpenAppPackageFileAsync("appsettings.json");
                using var reader = new StreamReader(stream);
                var contents = await reader.ReadToEndAsync();
                
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };

                _config = JsonSerializer.Deserialize<AppConfig>(contents, options);
            }
            catch (Exception ex)
            {
                // In case of failure, fall back to empty config or log
                System.Diagnostics.Debug.WriteLine($"Failed to load appsettings.json: {ex.Message}");
            }

            return _config ?? new AppConfig();
        }
    }
}
