namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Service contract for handling Microsoft Entra ID authentication via MSAL
    /// </summary>
    public interface IAuthenticationService
    {
        /// <summary>
        /// Gets the current active access token. Null if not authenticated.
        /// </summary>
        string? CurrentAccessToken { get; }

        /// <summary>
        /// Attempts to acquire a token silently from the cache.
        /// </summary>
        /// <returns>Access token, or null if silent acquisition fails</returns>
        Task<string?> AcquireTokenSilentAsync();

        /// <summary>
        /// Displays the Microsoft interactive login dialog to acquire a token.
        /// </summary>
        /// <returns>Access token, or null if login is cancelled/failed</returns>
        Task<string?> AcquireTokenInteractiveAsync();

        /// <summary>
        /// Clears tokens from the MSAL cache.
        /// </summary>
        Task SignOutAsync();
    }
}
