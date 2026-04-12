using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Windows.Input;
using NAudio.CoreAudioApi;

namespace AudioRecorder
{
    /// <summary>
    /// Логика взаимодействия для DeviceControl.xaml
    /// </summary>
    public partial class DeviceControl : UserControl, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }


        private bool _isSpeaker = true;
        private string _deviceName = "Unknown device";
        private string _volume = "0%";
        private bool _isHighlighted = false;
        private MMDevice _model;

        public event MouseButtonEventHandler Click;
        private void InvokeClickEvent(object sender, MouseButtonEventArgs e) => Click?.Invoke(this, e);


        public bool IsSpeaker
        {
            get => _isSpeaker;
            set
            {
                if (_isSpeaker != value)
                {
                    _isSpeaker = value;
                    OnPropertyChanged(nameof(DeviceIconPath));
                }
            }
        }

        public string DeviceName
        {
            get => _deviceName;
            set
            {
                if (_deviceName != value)
                {
                    _deviceName = value;
                    OnPropertyChanged(nameof(DeviceName));
                }
            }
        }

        public string Volume
        {
            get => _volume;
            set
            {
                if (_volume != value)
                {
                    _volume = value;
                    OnPropertyChanged(nameof(Volume));
                }
            }
        }

        public bool IsHighlighted
        {
            get => _isHighlighted;
            set
            {
                if (_isHighlighted != value)
                {
                    _isHighlighted = value;
                    ToggleSelection(value);
                }
            }
        }

        public MMDevice Model
        {
            get => _model;
            set => _model = value;
        }

        public string DeviceIconPath => IsSpeaker ? "/Icons/speaker.png" : "/Icons/microphone.png";


        public DeviceControl()
        {
            InitializeComponent();
            DataContext = this;
            bgRect.MouseLeftButtonDown += InvokeClickEvent;
        }

        private void ToggleSelection(bool value)
        {
            if (value)
            {
                IsHighlighted = value;
                bgRect.Stroke = TryFindResource("Accent") as SolidColorBrush ?? new SolidColorBrush(Colors.BlueViolet);
                bgRect.Fill = TryFindResource("Dark3") as SolidColorBrush ?? new SolidColorBrush(Colors.LightGray);
            }
            else
            {
                IsHighlighted = value;
                bgRect.Stroke = TryFindResource("Light2") as SolidColorBrush ?? new SolidColorBrush(Colors.LightSlateGray);
                bgRect.Fill = TryFindResource("Dark2") as SolidColorBrush ?? new SolidColorBrush(Colors.DarkGray);
            }
        }
    }
}
