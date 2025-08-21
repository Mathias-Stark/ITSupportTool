using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ITSupportToolGUI.Models
{
    public class InfoItem : INotifyPropertyChanged
    {
        private string _label;
        public string Label
        {
            get => _label;
            set { _label = value; OnPropertyChanged(); }
        }

        private string _value;
        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(); }
        }

        public string Id { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
