namespace SharePointAuthApp.Models
{
    /// <summary>
    /// Model representing a user mapping credentials stored in SharePoint Online auth_test list
    /// </summary>
    public class UserModel
    {
        /// <summary>
        /// Mapped from the 'Title' column in SharePoint
        /// </summary>
        public string Username { get; set; } = string.Empty;

        /// <summary>
        /// Mapped from the 'Password' column in SharePoint
        /// </summary>
        public string Password { get; set; } = string.Empty;
    }
}
