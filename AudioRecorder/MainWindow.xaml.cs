using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace AudioRecorder
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        #region I_NOTIFY_PROPERTY_CHANGED
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
        #endregion


        private WasapiLoopbackCapture _loopbackCapture;
        private WaveFileWriter _waveWriter;
        private MMDeviceEnumerator _deviceEnumerator;
        private DispatcherTimer _refreshTimer;
        private DispatcherTimer _recordingTimer;
        private Stopwatch _recordingElapsedTimer;
        private uint _countdownCounter = 3;

        public ObservableCollection<DeviceControl> Devices = new ObservableCollection<DeviceControl>();
        public ObservableCollection<TextBlock> Processes = new ObservableCollection<TextBlock>();

        private AppState _state = AppState.Idle;
        private enum AppState
        {
            Idle,
            Countdown,
            Recording
        }

        private static readonly string _defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                                                                   $"recording_{DateTime.Now:dd_MM_yyyy}");
        private string _filePath = _defaultPath;
        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                }
            }
        }
        public string FullFilePath
        {
            get => FilePath + GetFileExtension();
        }


        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;
            devicesPanel.ItemsSource = Devices;
            processesPanel.ItemsSource = Processes;
            _deviceEnumerator = new MMDeviceEnumerator();
            RefreshAudioDevices();

            _refreshTimer = new DispatcherTimer(DispatcherPriority.Normal);
            _refreshTimer.Interval = TimeSpan.FromSeconds(1);
            _refreshTimer.Tick += OnRefreshTimerTick;
            _refreshTimer.Start();
        }
        

        #region RESTORE DEFAULT WINDOW ANIMATIONS
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(HandleRef hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(HandleRef hWnd, int nIndex, IntPtr dwNewLong);

        public IntPtr myHWND;
        public const int GWL_STYLE = -16;

        public static class WS
        {
            public static readonly long
            WS_BORDER = 0x00800000L,
            WS_CAPTION = 0x00C00000L,
            WS_CHILD = 0x40000000L,
            WS_CHILDWINDOW = 0x40000000L,
            WS_CLIPCHILDREN = 0x02000000L,
            WS_CLIPSIBLINGS = 0x04000000L,
            WS_DISABLED = 0x08000000L,
            WS_DLGFRAME = 0x00400000L,
            WS_GROUP = 0x00020000L,
            WS_HSCROLL = 0x00100000L,
            WS_ICONIC = 0x20000000L,
            WS_MAXIMIZE = 0x01000000L,
            WS_MAXIMIZEBOX = 0x00010000L,
            WS_MINIMIZE = 0x20000000L,
            WS_MINIMIZEBOX = 0x00020000L,
            WS_OVERLAPPED = 0x00000000L,
            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
            WS_POPUP = 0x80000000L,
            WS_POPUPWINDOW = WS_POPUP | WS_BORDER | WS_SYSMENU,
            WS_SIZEBOX = 0x00040000L,
            WS_SYSMENU = 0x00080000L,
            WS_TABSTOP = 0x00010000L,
            WS_THICKFRAME = 0x00040000L,
            WS_TILED = 0x00000000L,
            WS_TILEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
            WS_VISIBLE = 0x10000000L,
            WS_VSCROLL = 0x00200000L;
        }

        public static IntPtr SetWindowLongPtr(HandleRef hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8)
            {
                return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            }
            else
            {
                return new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            myHWND = new WindowInteropHelper(this).Handle;
            IntPtr myStyle = new IntPtr(WS.WS_CAPTION | WS.WS_CLIPCHILDREN | WS.WS_MINIMIZEBOX | WS.WS_MAXIMIZEBOX | WS.WS_SYSMENU | WS.WS_SIZEBOX);
            SetWindowLongPtr(new HandleRef(null, myHWND), GWL_STYLE, myStyle);
        }
        #endregion

        #region WINDOW CONTROLS
        private void RedCircle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Close();
        }

        private void YellowCircle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }

        private void GreenCircle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        protected override void OnClosed(EventArgs e)
        {
            FreeResources();
            base.OnClosed(e);
        }
        #endregion


        #region APP CONTROLS EVENTS
        private void OnRefreshTimerTick(object sender, EventArgs e)
        {
            RefreshAudioDevices();
        }

        private void RefreshDevices_Click(object sender, RoutedEventArgs e)
        {
            RefreshAudioDevices();
        }

        private void RefreshProcesses_Click(object sender, RoutedEventArgs e)
        {
            DeviceControl selectedDevice = Devices.FirstOrDefault(d => d.IsHighlighted);
            if (selectedDevice != null)
                RefreshDeviceProcesses(selectedDevice.Model);
        }

        private void Device_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            foreach (DeviceControl d in Devices)
                d.IsHighlighted = false;

            ((DeviceControl)sender).IsHighlighted = true;
            RefreshDeviceProcesses(((DeviceControl)sender).Model);
        }

        private void PathTBox_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SelectFileLocation();
        }
        #endregion

        #region RECORDING EVENTS
        private void RecordButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state != AppState.Idle) return;

            if (File.Exists(FullFilePath)) // aviod unintensional file rewriting
            {
                string message = "File with the same name is already exists, would you like to erase it?";
                if (MessageBox.Show(message, "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                {
                    SelectFileLocation();
                }
            }

            // chage app state and stop all other activity
            _state = AppState.Countdown;
            _refreshTimer.Stop();

            // enable countdown panel and update recording icon
            SetupCountdownPanel(true, true);
            SyncRecordingIconWithAppState();

            // setup UX update timer
            if (_recordingTimer == null)
            {
                _recordingTimer = new DispatcherTimer(DispatcherPriority.Render)
                {
                    Interval = TimeSpan.FromSeconds(1)
                };
            }
            _recordingTimer.Tick += UpdateCountdown;
            _recordingTimer.Start();
        }

        private void UpdateCountdown(object sender, EventArgs e)
        {
            --_countdownCounter;

            if (_countdownCounter > 0) // countdown is not finished yet
            {
                countdownTBlock.Text = _countdownCounter.ToString();
            }
            else // countdown is finished, prepare everything for recording
            {
                // reset counter
                _countdownCounter = 3;

                // switch UX update callback
                _recordingTimer.Stop();
                _recordingTimer.Tick -= UpdateCountdown;
                _recordingTimer.Tick += UpdateRecordingTimer;

                // setup recording time timer
                if (_recordingElapsedTimer == null)
                    _recordingElapsedTimer = new Stopwatch();

                // change app state and update recording icon
                _state = AppState.Recording;
                SetupCountdownPanel(true, false);
                SyncRecordingIconWithAppState();
                ToggleLogoAnimation(true);

                // start recording process
                _recordingElapsedTimer.Start();
                _recordingTimer.Start();
                StartRecording();
            }
        }

        private void UpdateRecordingTimer(object sender, EventArgs e)
        {
            recordingTimerTBlock.Text = $"Recording... {_recordingElapsedTimer.Elapsed:hh\\:mm\\:ss}";
        }

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            _waveWriter?.Write(e.Buffer, 0, e.BytesRecorded);
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                Dispatcher.Invoke(() =>
                {
                    StopRecording();

                    MessageBox.Show($"Recording was interrupted by an occured error: {e.Exception.Message}", "Recording error",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void StopRecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state == AppState.Countdown)
            {
                StopRecordingPreparation();
            }
            else if (_state == AppState.Recording)
            {
                StopRecording();
            }
        }
        #endregion


        #region STYLES CONTROLS
        private void SetupCountdownPanel(bool enablePanel, bool enableCounter)
        {
            if (enablePanel)
            {
                countdownPanel.Visibility = Visibility.Visible;
                countdownPanel.IsHitTestVisible = true;
            }
            else
            {
                countdownPanel.Visibility = Visibility.Collapsed;
                countdownPanel.IsHitTestVisible = false;
            }

            if (enableCounter)
                countdownTBlock.Text = "3";
            else
                countdownTBlock.Text = string.Empty;
        }

        private void SyncRecordingIconWithAppState()
        {
            SolidColorBrush brush;

            switch (_state)
            {
                case AppState.Idle:
                    recordingTimerTBlock.Text = "Ready";
                    brush = (SolidColorBrush)TryFindResource("Light1") ?? new SolidColorBrush(Colors.White);
                    recordingTimerTBlock.Foreground = brush;
                    recordingIcon1.Stroke = brush;
                    recordingIcon2.Fill = brush;
                    break;

                case AppState.Countdown:
                    recordingTimerTBlock.Text = "Prepairing...";
                    brush = (SolidColorBrush)TryFindResource("Light1") ?? new SolidColorBrush(Colors.White);
                    recordingTimerTBlock.Foreground = brush;
                    recordingIcon1.Stroke = brush;
                    recordingIcon2.Fill = brush;
                    break;

                case AppState.Recording:
                    recordingTimerTBlock.Text = "Recording... 00:00:00";
                    brush = new SolidColorBrush(Colors.IndianRed);
                    recordingTimerTBlock.Foreground = brush;
                    recordingIcon1.Stroke = brush;
                    recordingIcon2.Fill = brush;
                    break;
            }
        }

        private void ToggleLogoAnimation(bool enable)
        {
            if (enable)
            {
                Storyboard storyboard = (Storyboard)appLogo.Resources["BlinkStoryboard"];
                storyboard.Begin();
            }
            else
            {
                Storyboard storyboard = (Storyboard)appLogo.Resources["BlinkStoryboard"];
                storyboard.Stop();

                appLogo.Fill = new SolidColorBrush(Colors.White);
            }
        }
        #endregion

        #region REFRESH DATA
        private void RefreshAudioDevices()
        {
            List<MMDevice> devices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active).ToList();
            if (Devices.Count > 0) // update existing device controls
            {
                List<DeviceControl> oldDevices = new List<DeviceControl>(Devices);
                Devices.Clear();
                int foundDeviceId = -1;

                foreach (MMDevice newD in devices)
                {
                    // search for existing device control
                    foundDeviceId = -1;
                    for (int i = 0; i < oldDevices.Count; i++)
                    {
                        if (newD.ID == oldDevices[i].Model.ID)
                        {
                            foundDeviceId = i;
                            break;
                        }
                    }

                    if (foundDeviceId != -1) // existing device control found
                    {
                        // update all necessary properties
                        oldDevices[foundDeviceId].Model = newD;
                        oldDevices[foundDeviceId].Volume = Math.Floor(newD.AudioEndpointVolume.MasterVolumeLevelScalar * 100).ToString() + '%';
                        Devices.Add(oldDevices[foundDeviceId]);
                    }
                    else // existing device control is not found, create new
                    {
                        DeviceControl control = new DeviceControl()
                        {
                            IsSpeaker = (newD.DataFlow == DataFlow.Render),
                            DeviceName = newD.FriendlyName,
                            Volume = Math.Floor(newD.AudioEndpointVolume.MasterVolumeLevelScalar * 100).ToString() + '%',
                            Model = newD
                        };
                        control.Click += Device_MouseLeftButtonDown;
                        Devices.Add(control);
                    }
                }
            }
            else // construct new device controls
            {
                Devices.Clear();

                foreach (MMDevice d in devices)
                {
                    DeviceControl control = new DeviceControl()
                    {
                        IsSpeaker = (d.DataFlow == DataFlow.Render),
                        DeviceName = d.FriendlyName,
                        Volume = Math.Floor(d.AudioEndpointVolume.MasterVolumeLevelScalar * 100).ToString() + '%',
                        Model = d
                    };
                    control.Click += Device_MouseLeftButtonDown;
                    Devices.Add(control);
                }
            }

            // update device processes if any of them is selected
            DeviceControl selectedDevice = Devices.FirstOrDefault(d => d.IsHighlighted);
            if (selectedDevice != null)
                RefreshDeviceProcesses(selectedDevice.Model);
            
            else if (Devices.Count > 0) // just select first
            {
                Devices[0].IsHighlighted = true;
                RefreshDeviceProcesses(Devices[0].Model);
            }
        }

        private void RefreshDeviceProcesses(MMDevice device)
        {
            if (Processes.Count > 0) // update existing process controls
            {
                List<TextBlock> oldProcesses = new List<TextBlock>(Processes);
                Processes.Clear();
                int foundProcessId = -1;

                for (int i = 0; i < device.AudioSessionManager.Sessions.Count; i++)
                {
                    // search for existing process control
                    var newSession = device.AudioSessionManager.Sessions[i];
                    string newName = Process.GetProcessById((int)newSession.GetProcessID).ProcessName;
                    bool isSystem = newSession.IsSystemSoundsSession;
                    bool isActive = newSession.State == AudioSessionState.AudioSessionStateActive && newSession.SimpleAudioVolume.Volume > 0;
                    foundProcessId = -1;

                    for (int j = 0; j < oldProcesses.Count; j++)
                    {
                        if (newName == oldProcesses[j].Text)
                        {
                            foundProcessId = j;
                            break;
                        }
                    }

                    if (foundProcessId != -1) // existing process control found
                    {
                        TextBlock tb = oldProcesses[foundProcessId];

                        if (!isSystem)
                        {
                            if (isActive) tb.Foreground = (SolidColorBrush)TryFindResource("Light0") ?? new SolidColorBrush(Colors.White);
                            else tb.Foreground = (SolidColorBrush)TryFindResource("Light2") ?? new SolidColorBrush(Colors.White);
                        }

                        Processes.Add(tb);
                    }
                    else // existing process control is not found, create new
                    {
                        TextBlock tb = new TextBlock()
                        {
                            Text = newName,
                            FontWeight = FontWeights.Light,
                            FontSize = 10
                        };
                        if (isSystem) tb.Foreground = new SolidColorBrush(Colors.CornflowerBlue);
                        else if (isActive) tb.Foreground = (SolidColorBrush)TryFindResource("Light0") ?? new SolidColorBrush(Colors.White);
                        else tb.Foreground = (SolidColorBrush)TryFindResource("Light2") ?? new SolidColorBrush(Colors.White);

                        Processes.Add(tb);
                    }
                }
            }
            else // construct new process control
            {
                Processes.Clear();

                for (int i = 0; i < device.AudioSessionManager.Sessions.Count; i++)
                {
                    var session = device.AudioSessionManager.Sessions[i];
                    string name = Process.GetProcessById((int)session.GetProcessID).ProcessName;
                    bool isSystem = session.IsSystemSoundsSession;
                    bool isActive = session.State == AudioSessionState.AudioSessionStateActive && session.SimpleAudioVolume.Volume > 0;

                    TextBlock tb = new TextBlock()
                    {
                        Text = name,
                        FontWeight = FontWeights.Light,
                        FontSize = 10
                    };
                    if (isSystem) tb.Foreground = new SolidColorBrush(Colors.CornflowerBlue);
                    else if (isActive) tb.Foreground = (SolidColorBrush)TryFindResource("Light0") ?? new SolidColorBrush(Colors.White);
                    else tb.Foreground = (SolidColorBrush)TryFindResource("Light2") ?? new SolidColorBrush(Colors.White);

                    Processes.Add(tb);
                }
            }
        }
        #endregion

        #region RECORDING
        private void StartRecording()
        {
            try
            {
                DeviceControl selectedControl = Devices.FirstOrDefault(d => d.IsHighlighted);
                MMDevice selectedDevice = selectedControl?.Model;
                if (selectedDevice == null) return;

                _loopbackCapture = new WasapiLoopbackCapture(selectedDevice);
                _waveWriter = new WaveFileWriter(FullFilePath, _loopbackCapture.WaveFormat);

                _loopbackCapture.DataAvailable += OnDataAvailable;
                _loopbackCapture.RecordingStopped += OnRecordingStopped;
                _loopbackCapture.StartRecording();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to start recording: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                CleanupCapture();
            }
        }

        private void StopRecordingPreparation()
        {
            // reset counter
            _countdownCounter = 3;

            // stop all timers
            _recordingTimer.Stop();
            _recordingTimer.Tick -= UpdateCountdown;

            // change app state and update UX
            _state = AppState.Idle;
            SetupCountdownPanel(false, false);
            SyncRecordingIconWithAppState();
            _refreshTimer.Start();
        }

        private void StopRecording()
        {
            // stop all timers
            _recordingTimer.Stop();
            _recordingTimer.Tick -= UpdateRecordingTimer;
            _recordingElapsedTimer.Reset();

            // Stop recording process
            _loopbackCapture?.StopRecording();
            CleanupCapture();

            // change app state and update UX
            _state = AppState.Idle;
            SetupCountdownPanel(false, false);
            SyncRecordingIconWithAppState();
            ToggleLogoAnimation(false);
            _refreshTimer.Start();
        }

        private void CleanupCapture()
        {
            if (_loopbackCapture != null)
            {
                _loopbackCapture.DataAvailable -= OnDataAvailable;
                _loopbackCapture.RecordingStopped -= OnRecordingStopped;

                _loopbackCapture.Dispose();
                _loopbackCapture = null;
            }

            if (_waveWriter != null)
            {
                _waveWriter.Dispose();
                _waveWriter = null;
            }
        }

        private void FreeResources()
        {
            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer.Tick -= OnRefreshTimerTick;
                _refreshTimer = null;
            }

            if (_recordingTimer != null)
            {
                _recordingTimer.Stop();
                if (_state == AppState.Countdown)
                    _recordingTimer.Tick -= UpdateCountdown;
                else if (_state == AppState.Recording)
                    _recordingTimer.Tick -= UpdateRecordingTimer;
                _recordingTimer = null;
            }

            if (_recordingElapsedTimer != null)
            {
                _recordingElapsedTimer.Stop();
                _recordingElapsedTimer = null;
            }
            
            if (_loopbackCapture != null)
                _loopbackCapture.StopRecording();
            CleanupCapture();
        }
        #endregion

        #region OTHER FUNCTIONS
        private void SelectFileLocation()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Wave files (*.wav)|*.wav|MP3 files (*.mp3)|*.mp3",
                FilterIndex = 1,
                FileName = $"recording_{DateTime.Now:dd_MM_yyyy}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                // path without extension
                FilePath = Path.Combine(Path.GetDirectoryName(saveFileDialog.FileName),
                           Path.GetFileNameWithoutExtension(saveFileDialog.FileName));
            }
        }

        private string GetFileExtension()
        {
            switch (extensionCBox.SelectedIndex)
            {
                case 0:
                    return ".wav";

                case 1:
                    return ".mp3";

                default:
                    return string.Empty;
            }
        }
        #endregion
    }
}
