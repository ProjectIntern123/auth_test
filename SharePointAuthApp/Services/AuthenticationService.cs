using Microsoft.Identity.Client;
using SharePointAuthApp.Models;

namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Implementation of AuthenticationService handling MSAL flow for Android
    /// </summary>
    public class AuthenticationService : IAuthenticationService
    {
        private readonly AppConfig _config;
        private IPublicClientApplication? _pca;
        private string? _currentAccessToken;
        
        // Default scopes for querying Microsoft Graph API for SharePoint Sites.
        private readonly string[] _scopes = new[] { "https://graph.microsoft.com/Sites.Read.All" };

        public string? CurrentAccessToken => _currentAccessToken;

        public AuthenticationService(AppConfig config)
        {
            _config = config;
        }

        private async Task EnsureInitializedAsync()
        {
            if (_pca != null) return;

            var authority = string.IsNullOrWhiteSpace(_config.TenantId) || _config.TenantId.Equals("common", StringComparison.OrdinalIgnoreCase)
                ? "https://login.microsoftonline.com/common"
                : $"https://login.microsoftonline.com/{_config.TenantId}";

            var builder = PublicClientApplicationBuilder.Create(_config.ClientId)
                .WithAuthority(authority)
                .WithRedirectUri($"msal{_config.ClientId}://auth");

#if ANDROID
            // For Android, we register the parent activity locator
            builder.WithParentActivityOrWindow(() => Microsoft.Maui.ApplicationModel.Platform.CurrentActivity);
#endif

            _pca = builder.Build();

            // Setup cache serialization if needed (token caching in secure storage)
            await SecureStorageTokenCache.Bind(_pca.UserTokenCache);
        }

        public async Task<string?> AcquireTokenSilentAsync()
        {
            await EnsureInitializedAsync();
            if (_pca == null) return null;

            var accounts = await _pca.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            if (firstAccount == null)
            {
                return null;
            }

            try
            {
                var result = await _pca.AcquireTokenSilent(_scopes, firstAccount)
                    .ExecuteAsync();

                _currentAccessToken = result.AccessToken;
                return _currentAccessToken;
            }
            catch (MsalUiRequiredException)
            {
                // Must authenticate interactively
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error acquiring token silently: {ex.Message}");
                return null;
            }
        }

        public async Task<string?> AcquireTokenInteractiveAsync()
        {
            await EnsureInitializedAsync();
            if (_pca == null) return null;

            try
            {
                var accounts = await _pca.GetAccountsAsync();
                var firstAccount = accounts.FirstOrDefault();

                var builder = _pca.AcquireTokenInteractive(_scopes);

                if (firstAccount != null)
                {
                    builder = builder.WithAccount(firstAccount);
                }

#if ANDROID
                // Specify parent activity for Android so that the browser prompt displays correctly
                builder = builder.WithParentActivityOrWindow(Microsoft.Maui.ApplicationModel.Platform.CurrentActivity);
#endif

                var result = await builder.ExecuteAsync();
                _currentAccessToken = result.AccessToken;
                return _currentAccessToken;
            }
            catch (MsalClientException ex) when (ex.ErrorCode == MsalError.AuthenticationCanceledError)
            {
                System.Diagnostics.Debug.WriteLine("Authentication cancelled by user.");
                throw new OperationCanceledException("Login was cancelled by the user.", ex);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Interactive login failed: {ex.Message}");
                throw;
            }
        }

        public async Task SignOutAsync()
        {
            await EnsureInitializedAsync();
            if (_pca == null) return;

            var accounts = (await _pca.GetAccountsAsync()).ToList();
            while (accounts.Any())
            {
                await _pca.RemoveAsync(accounts.First());
                accounts = (await _pca.GetAccountsAsync()).ToList();
            }

            _currentAccessToken = null;
        }
    }

    /// <summary>
    /// Helper to store MSAL tokens securely in platform SecureStorage
    /// </summary>
    internal static class SecureStorageTokenCache
    {
        private const string CacheKey = "MsalTokenCache";

        public static async Task Bind(ITokenCache tokenCache)
        {
            tokenCache.SetBeforeAccessAsync(async (args) =>
            {
                try
                {
                    var data = await SecureStorage.Default.GetAsync(CacheKey);
                    if (data != null)
                    {
                        args.DeserializeMsalV3(Convert.FromBase64String(data));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to load token cache: {ex.Message}");
                }
            });

            tokenCache.SetAfterAccessAsync(async (args) =>
            {
                if (args.HasStateChanged)
                {
                    try
                    {
                        var data = args.SerializeMsalV3();
                        var base64 = Convert.ToBase64String(data);
                        await SecureStorage.Default.SetAsync(CacheKey, base64);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to save token cache: {ex.Message}");
                    }
                }
            });
            
            await Task.CompletedTask;
        }
    }
}
