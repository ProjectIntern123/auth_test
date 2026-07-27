namespace SharePointAuthApp.Models
{
    /// <summary>
    /// Model mapping the settings loaded from appsettings.json
    /// </summary>
    public class AppConfig
    {
        public string TenantId { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
        public string SharePointSite { get; set; } = string.Empty;
        public string ListName { get; set; } = string.Empty;
    }
}
