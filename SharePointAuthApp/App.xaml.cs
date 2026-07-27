namespace SharePointAuthApp
{
    public partial class App : Application
    {
        public App()
        {
            InitializeComponent();

            // Set AppShell as the entry point of the app
            MainPage = new AppShell();
        }
    }
}
