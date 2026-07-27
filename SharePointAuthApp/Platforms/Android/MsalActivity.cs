using Android.App;
using Android.Content;
using Android.Content.PM;

namespace SharePointAuthApp
{
    /// <summary>
    /// Activity that intercepts the custom redirect URI from Microsoft Entra ID.
    /// Android will route the web browser redirect (msal{ClientId}://auth) to this activity,
    /// which then resumes the MSAL authentication flow.
    /// </summary>
    [Activity(Exported = true)]
    [IntentFilter(new[] { Intent.ActionView },
        Categories = new[] { Intent.CategoryDefault, Intent.CategoryBrowsable },
        // IMPORTANT: Replace 'YOUR_CLIENT_ID_HERE' with the Client ID from appsettings.json
        // e.g. DataScheme = "msal12345678-1234-1234-1234-1234567890ab"
        DataScheme = "msalYOUR_CLIENT_ID_HERE",
        DataPath = "/auth")]
    public class MsalActivity : BrowserTabActivity
    {
    }
}
