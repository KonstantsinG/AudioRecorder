using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Lame;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Resources;
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
        private LameMP3FileWriter _mp3Writer;
        private MMDeviceEnumerator _deviceEnumerator;
        private DispatcherTimer _refreshTimer;
        private DispatcherTimer _recordingTimer;
        private Stopwatch _recordingElapsedTimer;
        private CancellationTokenSource _conversionCancellation;
        private uint _countdownCounter = 3;
        private string _openExtension;
        private byte[] _coverImageData;
        private string _coverMimeType;
        private bool _usingDefaultCoverImage = true;

        public ObservableCollection<DeviceControl> Devices = new ObservableCollection<DeviceControl>();
        public ObservableCollection<TextBlock> Processes = new ObservableCollection<TextBlock>();

        private AppState _state = AppState.Idle;
        private enum AppState
        {
            Idle,
            Countdown,
            Recording,
            Converting
        }


        #region PROPS
        private static readonly string _defaultName = $"recording_{DateTime.Now:dd_MM_yyyy}";
        private static readonly string _defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                                                                   _defaultName);
        private static readonly string _defaultCoverImageUri = "pack://application:,,,/Icons/coverImage-default.jpg";

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

        private int _extensionIndex = 0;
        public int ExtensionIndex
        {
            get => _extensionIndex;
            set
            {
                if (_extensionIndex != value)
                {
                    _extensionIndex = value;
                    OnPropertyChanged(nameof(ExtensionIndex));
                }
            }
        }

        private int _conversionPercentage = 0;
        public int ConversionPercentage
        {
            get => _conversionPercentage;
            set
            {
                if (_conversionPercentage != value)
                {
                    _conversionPercentage = value;
                    OnPropertyChanged(nameof(ConversionPercentage));
                }
            }
        }

        private string _infoName;
        public string InfoName
        {
            get => _infoName;
            set
            {
                if (_infoName != value)
                {
                    _infoName = value;
                    OnPropertyChanged(nameof(InfoName));
                }
            }
        }

        private string _infoArtist;
        public string InfoArtist
        {
            get => _infoArtist;
            set
            {
                if (_infoArtist != value)
                {
                    _infoArtist = value;
                    OnPropertyChanged(nameof(InfoArtist));
                }
            }
        }

        private string _infoAlbum;
        public string InfoAlbum
        {
            get => _infoAlbum;
            set
            {
                if (_infoAlbum != value)
                {
                    _infoAlbum = value;
                    OnPropertyChanged(nameof(InfoAlbum));
                }
            }
        }

        private string _infoYear;
        public string InfoYear
        {
            get => _infoYear;
            set
            {
                if (_infoYear != value)
                {
                    _infoYear = value;
                    OnPropertyChanged(nameof(InfoYear));
                }
            }
        }

        private string _infoGenres;
        public string InfoGenres
        {
            get => _infoGenres;
            set
            {
                if (_infoGenres != value)
                {
                    _infoGenres = value;
                    OnPropertyChanged(nameof(InfoGenres));
                }
            }
        }

        private string _infoComment;
        public string InfoComment
        {
            get => _infoComment;
            set
            {
                if (_infoComment != value)
                {
                    _infoComment = value;
                    OnPropertyChanged(nameof(InfoComment));
                }
            }
        }
        #endregion


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

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            FreeRecordingResources();
            if (_state == AppState.Converting)
                AskBreakConversion(e);
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

        private void CoverImageControl_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SelectCoverImage();
        }

        private void CoverImageControl_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            LoadDefaultCoverImage();
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
            if (IsWavFormatSelected())
                _waveWriter?.Write(e.Buffer, 0, e.BytesRecorded);
            else
                _mp3Writer?.Write(e.Buffer, 0, e.BytesRecorded);
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

        #region METADATA EVENTS
        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state != AppState.Idle)
                return;

            OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = "Audio files (*.wav;*.mp3)|*.wav;*.mp3",
                Multiselect = false,
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                // save file path without extension
                FilePath = Path.Combine(Path.GetDirectoryName(dialog.FileName),
                           Path.GetFileNameWithoutExtension(dialog.FileName));

                // save and set file extension
                _openExtension = Path.GetExtension(dialog.FileName);
                if (GetFileExtension() != _openExtension)
                    SetFileExtension(_openExtension);

                // load file metadata
                ReadMetadata(dialog.FileName);
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_state != AppState.Idle)
                return;

            try
            {
                // check if user requested file conversion
                if (!CheckForFileConversionRequest()) // if not, write metadata immediately
                    WriteMetadata(FullFilePath);
                // otherwise we'll do it later after conversion
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to write metadata: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private void LoadDefaultCoverImage()
        {
            try
            {
                _usingDefaultCoverImage = true;
                var resourceUri = new Uri(_defaultCoverImageUri, UriKind.Absolute);
                StreamResourceInfo resourceInfo = Application.GetResourceStream(resourceUri);

                if (resourceInfo != null)
                {
                    using (var stream = resourceInfo.Stream)
                    {
                        _coverImageData = new byte[stream.Length];
                        stream.Read(_coverImageData, 0, _coverImageData.Length);
                    }

                    _coverMimeType = "image/jpeg";
                    DisplayCoverImage(_coverImageData);
                }
                else
                {
                    ClearCoverImage();
                }
            }
            catch (Exception)
            {
                ClearCoverImage();
            }
        }

        private void ClearCoverImage()
        {
            _usingDefaultCoverImage = true;
            infoCoverImageControl.Source = null;
            _coverImageData = null;
            _coverMimeType = null;
        }

        private void ToggleConversionPanel(bool enable)
        {
            if (enable)
            {
                conversionPanel.Visibility = Visibility.Visible;
                conversionPanel.IsHitTestVisible = true;
            }
            else
            {
                conversionPanel.Visibility = Visibility.Collapsed;
                conversionPanel.IsHitTestVisible = false;
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

                if (IsWavFormatSelected()) // recording .WAV
                {
                    _waveWriter = new WaveFileWriter(FullFilePath, _loopbackCapture.WaveFormat);
                    // add metadata
                }
                else // recording .MP3
                {
                    // add metadata
                    _mp3Writer = new LameMP3FileWriter(FullFilePath, _loopbackCapture.WaveFormat, LAMEPreset.STANDARD);
                }

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

            if (_mp3Writer != null)
            {
                _mp3Writer.Dispose();
                _mp3Writer = null;
            }
        }

        private void FreeRecordingResources()
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

        #region METADATA EDITING
        private bool CheckForFileConversionRequest()
        {
            bool conversionRequested = GetFileExtension() != _openExtension;

            if (conversionRequested)
            {
                _state = AppState.Converting;
                ToggleConversionPanel(true);

                if (_openExtension == ".wav") // .wav to .mp3
                    ConvertWav2Mp3Async(FilePath + _openExtension, FullFilePath);
                else if (_openExtension == ".mp3") // .mp3 to .wav
                    ConvertMp32WavAsync(FilePath + _openExtension, FullFilePath);
            }

            return conversionRequested;
        }

        private void WriteMetadata(string path)
        {
            using (var tagFile = TagLib.File.Create(path))
            {
                // name
                tagFile.Tag.Title = InfoName ?? string.Empty;
                tagFile.Tag.TitleSort = InfoName ?? string.Empty;

                // artist
                tagFile.Tag.AlbumArtists = new[] { InfoArtist ?? string.Empty };
                tagFile.Tag.AlbumArtistsSort = new[] { InfoArtist ?? string.Empty };
                tagFile.Tag.Performers = new[] { InfoArtist ?? string.Empty };
                tagFile.Tag.PerformersSort = new[] { InfoArtist ?? string.Empty };
                tagFile.Tag.Composers = new[] { InfoArtist ?? string.Empty };
                tagFile.Tag.ComposersSort = new[] { InfoArtist ?? string.Empty };

                // album
                tagFile.Tag.Album = InfoAlbum ?? string.Empty;
                tagFile.Tag.AlbumSort = InfoAlbum ?? string.Empty;

                // year, comment, genres
                if (uint.TryParse(InfoYear, out uint year)) tagFile.Tag.Year = year;
                tagFile.Tag.Comment = InfoComment ?? string.Empty;
                tagFile.Tag.Genres = string.IsNullOrWhiteSpace(InfoGenres) ? new string[] { } : // if entry is not empty - split it by the ',' character and trim spaces
                                     InfoGenres.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();

                // cover image
                if (!_usingDefaultCoverImage) // save cover image
                {
                    TagLib.Picture pic = new TagLib.Picture
                    {
                        Type = TagLib.PictureType.FrontCover,
                        Description = "Cover",
                        MimeType = _coverMimeType ?? "image/jpeg",
                        Data = new TagLib.ByteVector(_coverImageData)
                    };
                    tagFile.Tag.Pictures = new TagLib.IPicture[] { pic };
                }
                else // erase cover image
                {
                    tagFile.Tag.Pictures = new TagLib.IPicture[] { };
                }

                tagFile.Save();
            }
        }

        private void ReadMetadata(string path)
        {
            using (var tagFile = TagLib.File.Create(path))
            {
                // name
                if (!string.IsNullOrEmpty(tagFile.Tag.Title)) InfoName = tagFile.Tag.Title;
                else if (!string.IsNullOrEmpty(tagFile.Tag.TitleSort)) InfoName = tagFile.Tag.TitleSort;
                else InfoName = string.Empty;

                // artist
                if (tagFile.Tag.AlbumArtists.Length > 0) InfoArtist = tagFile.Tag.AlbumArtists[0];
                else if (tagFile.Tag.AlbumArtistsSort.Length > 0) InfoArtist = tagFile.Tag.AlbumArtistsSort[0];
                else if (tagFile.Tag.Performers.Length > 0) InfoArtist = tagFile.Tag.Performers[0];
                else if (tagFile.Tag.PerformersSort.Length > 0) InfoArtist = tagFile.Tag.PerformersSort[0];
                else if (tagFile.Tag.Composers.Length > 0) InfoArtist = tagFile.Tag.Composers[0];
                else if (tagFile.Tag.ComposersSort.Length > 0) InfoArtist = tagFile.Tag.ComposersSort[0];
                else InfoArtist = string.Empty;

                // album
                if (!string.IsNullOrEmpty(tagFile.Tag.Album)) InfoAlbum = tagFile.Tag.Album;
                else if (!string.IsNullOrEmpty(tagFile.Tag.AlbumSort)) InfoAlbum = tagFile.Tag.AlbumSort;
                else InfoAlbum = string.Empty;

                // year, comment, genres
                if (tagFile.Tag.Year != default) InfoYear = tagFile.Tag.Year.ToString();
                else InfoYear = string.Empty;

                if (!string.IsNullOrEmpty(tagFile.Tag.Comment)) InfoComment = tagFile.Tag.Comment;
                else InfoComment = string.Empty;

                if (tagFile.Tag.Genres.Length > 0) InfoGenres = string.Join(", ", tagFile.Tag.Genres);
                else InfoGenres = string.Empty;

                // cover image
                if (tagFile.Tag.Pictures.Length > 0)
                {
                    var picture = tagFile.Tag.Pictures[0];
                    _coverImageData = picture.Data.Data;
                    _coverMimeType = picture.MimeType;
                    DisplayCoverImage(_coverImageData);
                }
                else
                {
                    LoadDefaultCoverImage();
                }
            }
        }
        #endregion

        #region CONVERSIONS
        private async void ConvertWav2Mp3Async(string wavPath, string mp3Path, int bitRate = 192)
        {
            _conversionCancellation = new CancellationTokenSource();
            string tempPath = Path.GetTempFileName();

            try
            {
                await Task.Run(() =>
                {
                    using (var reader = new WaveFileReader(wavPath))
                    using (var writer = new LameMP3FileWriter(tempPath, reader.WaveFormat, bitRate))
                    {
                        // full data amount
                        long totalBytes = reader.Length;
                        long bytesProcessed = 0;
                        byte[] buffer = new byte[8192]; // 8 KB buffer

                        int bytesRead;
                        while ((bytesRead = reader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            _conversionCancellation.Token.ThrowIfCancellationRequested(); // throw exception if someone requested it
                            writer.Write(buffer, 0, bytesRead);
                            bytesProcessed += bytesRead;

                            // compute and report progress
                            ConversionPercentage = (int)((double)bytesProcessed / totalBytes * 100);
                        }
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // free cancellation token and erase temp file
                _conversionCancellation.Dispose();
                _conversionCancellation = null;
                EraseTempFile(tempPath);
                return;
            }

            // free cancellation token
            _conversionCancellation.Dispose();
            _conversionCancellation = null;

            // move converted file from temp to destination directory
            MoveTempFile(tempPath, mp3Path);

            // once we finished conversion, write metadata to a new file
            WriteMetadata(mp3Path);

            // switch the state and toggle the progress bar
            _state = AppState.Idle;
            ToggleConversionPanel(false);
        }

        private async void ConvertMp32WavAsync(string mp3Path, string wavPath)
        {
            _conversionCancellation = new CancellationTokenSource();
            string tempPath = Path.GetTempFileName();

            try
            {
                await Task.Run(() =>
                {
                    using (var reader = new MediaFoundationReader(mp3Path))
                    using (var writer = new WaveFileWriter(tempPath, reader.WaveFormat))
                    {
                        // full data amount
                        long totalBytes = reader.Length;
                        long bytesProcessed = 0;
                        byte[] buffer = new byte[8192]; // 8 KB buffer

                        int bytesRead;
                        while ((bytesRead = reader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            _conversionCancellation.Token.ThrowIfCancellationRequested(); // throw exception if someone requested it
                            writer.Write(buffer, 0, bytesRead);
                            bytesProcessed += bytesRead;

                            // compute and report progress
                            ConversionPercentage = (int)((double)bytesProcessed / totalBytes * 100);
                        }
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // free cancellation token and erase temp file
                _conversionCancellation.Dispose();
                _conversionCancellation = null;
                EraseTempFile(tempPath);
                return;
            }

            // free cancellation token
            _conversionCancellation.Dispose();
            _conversionCancellation = null;

            // move converted file from temp to destination directory
            MoveTempFile(tempPath, wavPath);

            // switch the state and toggle the progress bar
            _state = AppState.Idle;
            ToggleConversionPanel(false);

            // once we finished conversion, write metadata to a new file
            WriteMetadata(wavPath);
        }

        private void EraseTempFile(string tempPath)
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }

        private void MoveTempFile(string tempPath, string destinationPath)
        {
            if (File.Exists(destinationPath))
                File.Delete(destinationPath);

            File.Move(tempPath, destinationPath);
        }

        private async void AskBreakConversion(CancelEventArgs e)
        {
            var result = MessageBox.Show(
                "The file conversion is not yet complete. If you close the application now, the entire unfinished process will be lost.\n\n" +
                "Wait for the convertion to be completed?",
                "Warning",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.No) // no pressed -> erase conversion temporary file
            {
                _conversionCancellation?.Cancel();
                await Task.Delay(500);
            }
            else if (result == MessageBoxResult.Yes) // yes pressed -> wait until conversion is finished
            {
                e.Cancel = true; // temporary cancel shutdown

                while (_state == AppState.Converting)
                {
                    await Task.Delay(100);
                }

                Application.Current.Shutdown();
            }
        }
        #endregion

        #region OTHER FUNCTIONS
        private void SelectCoverImage()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*",
                Title = "Select cover image",
                Multiselect = false
            };

            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    // selected file must be less than 5 MB in size
                    var fileInfo = new FileInfo(openFileDialog.FileName);
                    if (fileInfo.Length > 5 * 1024 * 1024)
                    {
                        MessageBox.Show("Cover image must be less than 5 MB in size.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // select mime-type
                    string extension = Path.GetExtension(openFileDialog.FileName).ToLower();
                    switch (extension)
                    {
                        case ".jpg":
                            _coverMimeType = "image/jpeg";
                            break;

                        case ".jpeg":
                            _coverMimeType = "image/jpeg";
                            break;

                        case ".png":
                            _coverMimeType = "image/png";
                            break;

                        case ".bmp":
                            _coverMimeType = "image/bmp";
                            break;

                        default:
                            _coverMimeType = "image/jpeg";
                            break;
                    }

                    _usingDefaultCoverImage = false;
                    _coverImageData = File.ReadAllBytes(openFileDialog.FileName);
                    DisplayCoverImage(_coverImageData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Cover image loading error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LoadDefaultCoverImage();
            }
        }

        private void DisplayCoverImage(byte[] data)
        {
            if (data == null || data.Length == 0)
            {
                ClearCoverImage();
                return;
            }

            try
            {
                using (var ms = new MemoryStream(data))
                {
                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = ms;
                    bitmap.EndInit();
                    bitmap.Freeze(); // allow using image in different threads

                    infoCoverImageControl.Source = bitmap;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Image diplay error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearCoverImage();
            }
        }

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
            switch (ExtensionIndex)
            {
                case 0:
                    return ".wav";

                case 1:
                    return ".mp3";

                default:
                    return string.Empty;
            }
        }

        private void SetFileExtension(string extension)
        {
            switch (extension)
            {
                case ".wav":
                    ExtensionIndex = 0;
                    break;

                case ".mp3":
                    ExtensionIndex = 1;
                    break;
            }
        }

        private bool IsWavFormatSelected() => _extensionIndex == 0;
        #endregion
    }
}
