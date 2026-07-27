using SharePointAuthApp.Models;

namespace SharePointAuthApp.Services
{
    /// <summary>
    /// Service contract for querying SharePoint Online list details via Microsoft Graph
    /// </summary>
    public interface ISharePointService
    {
        /// <summary>
        /// Retrieves all users stored in the auth_test list from SharePoint
        /// </summary>
        Task<List<UserModel>> GetAllUsersAsync();

        /// <summary>
        /// Validates that the entered username and password exist in the SharePoint list
        /// </summary>
        Task<bool> ValidateCredentialsAsync(string username, string password);
    }
}
