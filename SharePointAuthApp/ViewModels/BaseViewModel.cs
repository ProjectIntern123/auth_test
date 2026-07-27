using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointAuthApp.ViewModels
{
    /// <summary>
    /// Base class for all ViewModels implementing INotifyPropertyChanged via ObservableObject
    /// </summary>
    public partial class BaseViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsNotBusy))]
        private bool _isBusy;

        [ObservableProperty]
        private string _title = string.Empty;

        [ObservableProperty]
        private string _errorMessage = string.Empty;

        public bool IsNotBusy => !_isBusy;

        [ObservableProperty]
        private bool _hasError;
    }
}
