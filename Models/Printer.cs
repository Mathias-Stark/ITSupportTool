using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ITSupportToolGUI.Models
{
    public class Printer : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string? Name { get; set; }
        public string? ServerName { get; set; }
        public string? DriverName { get; set; }
        public string? Comment { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
