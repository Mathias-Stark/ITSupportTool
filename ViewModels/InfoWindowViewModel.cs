using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace ITSupportToolGUI.ViewModels
{
    public class InfoWindowViewModel : INotifyPropertyChanged
    {
        private bool _isVisible;
        private string _title;
        private object _content;

        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(); }
        }

        public string Title
        {
            get => _title;
            set { _title = value; OnPropertyChanged(); }
        }

        public object Content
        {
            get => _content;
            set { _content = value; OnPropertyChanged(); }
        }

        public ICommand CloseCommand { get; }

        public InfoWindowViewModel()
        {
            CloseCommand = new RelayCommand(_ => IsVisible = false);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
