using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointAuthApp.Services;

namespace SharePointAuthApp.ViewModels
{
    /// <summary>
    /// ViewModel for the Home screen. Handles welcoming the logged-in user and signing out.
    /// </summary>
    [QueryProperty(nameof(Username), "username")]
    public partial class HomeViewModel : BaseViewModel
    {
        private readonly IAuthenticationService _authService;
        private readonly INavigationService _navigationService;

        [ObservableProperty]
        private string _username = string.Empty;

        public HomeViewModel(
            IAuthenticationService authService,
            INavigationService navigationService)
        {
            _authService = authService;
            _navigationService = navigationService;
            Title = "Welcome Home";
        }

        /// <summary>
        /// Command executed when clicking the Logout button
        /// </summary>
        [RelayCommand]
        private async Task LogoutAsync()
        {
            if (IsBusy) return;

            IsBusy = true;

            try
            {
                // Clear MSAL credentials and cache
                await _authService.SignOutAsync();

                // Navigate back to the login root route
                await _navigationService.NavigateToRootAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Logout error: {ex.Message}");
                ErrorMessage = "Failed to log out correctly.";
                HasError = true;
            }
            finally
            {
                IsBusy = false;
            }
        }
    }
}
