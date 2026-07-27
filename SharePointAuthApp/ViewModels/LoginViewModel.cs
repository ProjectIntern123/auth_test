using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SharePointAuthApp.Services;
using System.Windows.Input;

namespace SharePointAuthApp.ViewModels
{
    /// <summary>
    /// ViewModel for handling the Login page logic, validation, MSAL auth, and SharePoint checks
    /// </summary>
    public partial class LoginViewModel : BaseViewModel
    {
        private readonly IAuthenticationService _authService;
        private readonly ISharePointService _sharePointService;
        private readonly INavigationService _navigationService;

        [ObservableProperty]
        private string _username = string.Empty;

        [ObservableProperty]
        private string _password = string.Empty;

        public LoginViewModel(
            IAuthenticationService authService,
            ISharePointService sharePointService,
            INavigationService navigationService)
        {
            _authService = authService;
            _sharePointService = sharePointService;
            _navigationService = navigationService;
            Title = "Sign In";
        }

        /// <summary>
        /// Command executed when clicking the Login button
        /// </summary>
        [RelayCommand]
        private async Task LoginAsync()
        {
            if (IsBusy) return;

            // Reset error states
            ErrorMessage = string.Empty;
            HasError = false;

            // 1. Inputs validation
            if (string.IsNullOrWhiteSpace(Username))
            {
                ErrorMessage = "Username is required.";
                HasError = true;
                return;
            }

            if (string.IsNullOrWhiteSpace(Password))
            {
                ErrorMessage = "Password is required.";
                HasError = true;
                return;
            }

            IsBusy = true;

            try
            {
                // Check network availability
                if (Connectivity.Current.NetworkAccess != NetworkAccess.Internet)
                {
                    ErrorMessage = "No internet connection. Please check your network and try again.";
                    HasError = true;
                    return;
                }

                // 2. Authenticate with Microsoft via MSAL
                System.Diagnostics.Debug.WriteLine("Starting Microsoft Authentication...");
                string? token = await _authService.AcquireTokenInteractiveAsync();
                
                if (string.IsNullOrEmpty(token))
                {
                    ErrorMessage = "Microsoft login failed. Access token could not be obtained.";
                    HasError = true;
                    return;
                }

                // 3. Connect to SharePoint and validate entered credentials
                System.Diagnostics.Debug.WriteLine("Fetching SharePoint list and validating credentials...");
                bool isValidUser = await _sharePointService.ValidateCredentialsAsync(Username, Password);

                if (isValidUser)
                {
                    // 4. Success -> Clear fields and navigate to Home Screen
                    var enteredUser = Username;
                    Username = string.Empty;
                    Password = string.Empty;

                    // Pass username to the home page via query parameters
                    await _navigationService.NavigateToAsync($"home?username={Uri.EscapeDataString(enteredUser)}");
                }
                else
                {
                    // 5. Failure -> Show invalid credentials error
                    ErrorMessage = "Invalid Username or Password";
                    HasError = true;
                }
            }
            catch (OperationCanceledException)
            {
                ErrorMessage = "Login was cancelled.";
                HasError = true;
            }
            catch (Microsoft.Identity.Client.MsalClientException ex) when (ex.ErrorCode == "authentication_canceled")
            {
                ErrorMessage = "Login was cancelled.";
                HasError = true;
            }
            catch (HttpRequestException)
            {
                ErrorMessage = "Network connection failure. Failed to connect to SharePoint Online.";
                HasError = true;
            }
            catch (Exception ex)
            {
                // Handle different error messages based on exception details
                var message = ex.Message.ToLower();
                if (message.Contains("accessdenied") || message.Contains("forbidden") || message.Contains("access denied"))
                {
                    ErrorMessage = "Access denied. Please verify your Entra ID application registration has 'Sites.Read.All' permissions.";
                }
                else if (message.Contains("notfound") || message.Contains("does not exist"))
                {
                    ErrorMessage = "SharePoint connection failure. Site or list was not found.";
                }
                else if (message.Contains("empty") || message.Contains("could not be loaded"))
                {
                    ErrorMessage = "The SharePoint authentication list is empty or could not be loaded.";
                }
                else
                {
                    ErrorMessage = $"An unexpected error occurred: {ex.Message}";
                }
                
                HasError = true;
                System.Diagnostics.Debug.WriteLine($"Login exception: {ex}");
            }
            finally
            {
                IsBusy = false;
            }
        }
    }
}
