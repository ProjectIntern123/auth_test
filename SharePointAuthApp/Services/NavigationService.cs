namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Implementation of NavigationService wrapping MAUI AppShell GoToAsync routing
    /// </summary>
    public class NavigationService : INavigationService
    {
        public async Task NavigateToAsync(string route, IDictionary<string, object>? parameters = null)
        {
            if (Shell.Current == null)
            {
                System.Diagnostics.Debug.WriteLine("Shell.Current is null. Cannot navigate.");
                return;
            }

            if (parameters != null)
            {
                await Shell.Current.GoToAsync(route, parameters);
            }
            else
            {
                await Shell.Current.GoToAsync(route);
            }
        }

        public async Task NavigateBackAsync()
        {
            if (Shell.Current == null) return;
            await Shell.Current.GoToAsync("..");
        }

        public async Task NavigateToRootAsync()
        {
            if (Shell.Current == null) return;
            // Go to the LoginPage route directly. Using "//login" is typical for route changes that reset the stack.
            await Shell.Current.GoToAsync("//login");
        }
    }
}
