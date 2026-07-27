namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Service contract for handling page routing within MAUI AppShell
    /// </summary>
    public interface INavigationService
    {
        /// <summary>
        /// Navigates to a specific page route with optional parameters
        /// </summary>
        Task NavigateToAsync(string route, IDictionary<string, object>? parameters = null);

        /// <summary>
        /// Navigates back one step in the navigation stack
        /// </summary>
        Task NavigateBackAsync();

        /// <summary>
        /// Returns to the root route (typically LoginPage)
        /// </summary>
        Task NavigateToRootAsync();
    }
}
