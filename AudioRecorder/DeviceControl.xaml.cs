using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace AudioRecorder
{
    /// <summary>
    /// Логика взаимодействия для DeviceControl.xaml
    /// </summary>
    public partial class DeviceControl : UserControl, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


        private bool _isSpeaker = true;
        private string _deviceName = "Unknown device";
        private string _volume = "0%";
        private bool _isHighlighted = false;

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

        public string DeviceIconPath => IsSpeaker ? "/Icons/speaker.png" : "/Icons/microphone.png";


        public static readonly DependencyProperty IsSpeakerProperty = DependencyProperty.Register(
            nameof(IsSpeaker), typeof(bool), typeof(DeviceControl), new PropertyMetadata(true, OnIsSpeakerPropertyChanged));

        public static readonly DependencyProperty DeviceNameProperty = DependencyProperty.Register(
            nameof(DeviceName), typeof(string), typeof(DeviceControl), new PropertyMetadata("Unknown device", OnDeviceNamePropertyChanged));

        public static readonly DependencyProperty VolumeProperty = DependencyProperty.Register(
            nameof(Volume), typeof(string), typeof(DeviceControl), new PropertyMetadata("0%", OnVolumePropertyChanged));


        private static void OnIsSpeakerPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DeviceControl control)
                control.IsSpeaker = (bool)e.NewValue;
        }

        private static void OnDeviceNamePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DeviceControl control)
                control.DeviceName = (string)e.NewValue;
        }

        private static void OnVolumePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DeviceControl control)
                control.Volume = (string)e.NewValue;
        }


        public static readonly RoutedEvent ClickEvent = EventManager.RegisterRoutedEvent(
            nameof(Click), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(DeviceControl));

        public event RoutedEventHandler Click
        {
            add => AddHandler(ClickEvent, value);
            remove => RemoveHandler(ClickEvent, value);
        }


        public DeviceControl()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void ToggleSelection(bool value)
        {
            if (value)
            {
                IsHighlighted = value;
                bgRect.Stroke = TryFindResource("Accent") as SolidColorBrush ?? new SolidColorBrush(Colors.BlueViolet);
            }
            else
            {
                IsHighlighted = value;
                bgRect.Stroke = TryFindResource("Light2") as SolidColorBrush ?? new SolidColorBrush(Colors.LightSlateGray);
            }
        }
    }
}
