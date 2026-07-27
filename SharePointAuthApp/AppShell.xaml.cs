using SharePointAuthApp.Views;

namespace SharePointAuthApp
{
    public partial class AppShell : Shell
    {
        public AppShell()
        {
            InitializeComponent();

            // Register routes for pages that are navigated to dynamically
            Routing.RegisterRoute("home", typeof(HomePage));
        }
    }
}
