using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using SharePointAuthApp.Models;
using System.Text.Json;

namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Service to communicate with SharePoint Online via Microsoft Graph API
    /// </summary>
    public class SharePointService : ISharePointService
    {
        private readonly AppConfig _config;
        private readonly IAuthenticationService _authService;

        public SharePointService(AppConfig config, IAuthenticationService authService)
        {
            _config = config;
            _authService = authService;
        }

        /// <summary>
        /// Instantiates a GraphServiceClient using the bearer token provider
        /// </summary>
        private GraphServiceClient GetGraphClient()
        {
            var tokenProvider = new GraphAccessTokenProvider(_authService);
            var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
            return new GraphServiceClient(authProvider);
        }

        public async Task<List<UserModel>> GetAllUsersAsync()
        {
            // Verify internet connectivity first
            if (Connectivity.Current.NetworkAccess != NetworkAccess.Internet)
            {
                throw new HttpRequestException("No internet connection available. Please check your network status.");
            }

            var graphClient = GetGraphClient();
            var users = new List<UserModel>();

            try
            {
                // Parse host and path from SharePoint Site URL
                // Format: https://hartek.sharepoint.com/sites/OPSTestSite
                var siteUri = new Uri(_config.SharePointSite);
                var hostname = siteUri.Host;
                var sitePath = siteUri.AbsolutePath.TrimStart('/');

                // 1. Fetch site identifier
                var site = await graphClient.Sites[$"{hostname}:/{sitePath}"].GetAsync();
                if (site == null || string.IsNullOrEmpty(site.Id))
                {
                    throw new Exception("Unable to find the specified SharePoint site.");
                }

                // 2. Fetch list items and expand fields
                // We use the List Name/Title or ID directly.
                var itemsCollection = await graphClient.Sites[site.Id]
                    .Lists[_config.ListName]
                    .Items
                    .GetAsync(requestConfiguration =>
                    {
                        // Expand the fields property to get item columns (Title, Password, etc.)
                        requestConfiguration.QueryParameters.Expand = new[] { "fields" };
                    });

                if (itemsCollection?.Value == null)
                {
                    return users; // Empty list
                }

                foreach (var item in itemsCollection.Value)
                {
                    if (item.Fields?.AdditionalData != null)
                    {
                        var fields = item.Fields.AdditionalData;

                        // Retrieve the Title (Username) and Password columns case-insensitively
                        var username = GetFieldValue(fields, "Title");
                        var password = GetFieldValue(fields, "Password");

                        if (!string.IsNullOrEmpty(username))
                        {
                            users.Add(new UserModel
                            {
                                Username = username,
                                Password = password ?? string.Empty
                            });
                        }
                    }
                }
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError graphEx)
            {
                System.Diagnostics.Debug.WriteLine($"Graph API Error: {graphEx.Error?.Message}");
                throw new Exception($"Graph API Error: {graphEx.Error?.Message ?? "Access denied or resource not found."}", graphEx);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SharePoint fetch exception: {ex.Message}");
                throw;
            }

            return users;
        }

        public async Task<bool> ValidateCredentialsAsync(string username, string password)
        {
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                return false;
            }

            var allUsers = await GetAllUsersAsync();
            if (allUsers == null || !allUsers.Any())
            {
                throw new Exception("The SharePoint authentication list is empty or could not be loaded.");
            }

            // Find matching username and password (case sensitive/insensitive based on requirements, standard is Username case insensitive, password case sensitive)
            var matchedUser = allUsers.FirstOrDefault(u => 
                u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) && 
                u.Password.Equals(password, StringComparison.Ordinal));

            return matchedUser != null;
        }

        private string? GetFieldValue(IDictionary<string, object> fields, string key)
        {
            // Case-insensitive key lookup in field dictionary
            var match = fields.FirstOrDefault(f => f.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
            if (match.Value == null)
            {
                return null;
            }

            // Graph API returns values as JsonElement or basic types depending on Kiota serialization
            if (match.Value is JsonElement jsonElem)
            {
                return jsonElem.ValueKind switch
                {
                    JsonValueKind.String => jsonElem.GetString(),
                    JsonValueKind.Number => jsonElem.GetRawText(),
                    JsonValueKind.True => "True",
                    JsonValueKind.False => "False",
                    _ => jsonElem.ToString()
                };
            }

            return match.Value.ToString();
        }
    }

    /// <summary>
    /// Custom access token provider bridging IAuthenticationService to Graph SDK Kiota client
    /// </summary>
    internal class GraphAccessTokenProvider : IAccessTokenProvider
    {
        private readonly IAuthenticationService _authService;

        public GraphAccessTokenProvider(IAuthenticationService authService)
        {
            _authService = authService;
        }

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri, 
            Dictionary<string, object>? additionalAuthenticationContext = null, 
            CancellationToken cancellationToken = default)
        {
            // Fetch token. Active interactive flow is executed in ViewModels, so here we assume it is cached or we do a silent check.
            var token = _authService.CurrentAccessToken ?? await _authService.AcquireTokenSilentAsync();
            return token ?? string.Empty;
        }

        public AllowedHostsValidator AllowedHostsValidator { get; } = new AllowedHostsValidator();
    }
}
