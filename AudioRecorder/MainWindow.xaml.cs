using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace AudioRecorder
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private WasapiLoopbackCapture _loopbackCapture;
        private WaveFileWriter _waveWriter;
        private MMDeviceEnumerator _deviceEnumerator;


        public MainWindow()
        {
            InitializeComponent();
            //LoadAudioDevices();
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




        private void LoadAudioDevices()
        {
            _deviceEnumerator = new MMDeviceEnumerator();
            var devices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            //cmbOutputDevices.ItemsSource = devices.ToList();

            try
            {
                var defaultDevice = _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
                //cmbOutputDevices.SelectedItem = devices.FirstOrDefault(d => d.ID == defaultDevice.ID);

                //if (cmbOutputDevices.SelectedItem == null && devices.Count > 0)
                    //cmbOutputDevices.SelectedIndex = 0;
            }
            catch
            {
                //if (cmbOutputDevices.Items.Count == 0)
                    //txtStatus.Text = "Output devices is not found...";
            }
        }

        private void btnRecord_Click(object sender, RoutedEventArgs e)
        {
            if (_loopbackCapture != null)
            {
                StopRecording();
                return;
            }

            //if (cmbOutputDevices.SelectedItem == null)
            {
                MessageBox.Show("Please select an output device for recording", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            //StartRecording();
        }

        private async void StartRecording()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Wave files (*.wav)|*.wav)",
                FileName = $"recording_{DateTime.Now:ddMMyyyy_HHmm}.wav"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                //var selectedDevice = (MMDevice)cmbOutputDevices.SelectedItem;
                //_loopbackCapture = new WasapiLoopbackCapture(selectedDevice);
                _waveWriter = new WaveFileWriter(saveFileDialog.FileName, _loopbackCapture.WaveFormat);

                _loopbackCapture.DataAvailable += OnDataAvailable;
                _loopbackCapture.RecordingStopped += OnRecordingStopped;

                _loopbackCapture.StartRecording();
                //btnRecord.Content = "Stop Recording";
                //txtStatus.Text = $"Recording {System.IO.Path.GetFileName(saveFileDialog.FileName)}...";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to start recording: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                CleanupCapture();
            }
        }

        private void StopRecording()
        {
            if (_loopbackCapture != null)
                _loopbackCapture.StopRecording();
        }

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            _waveWriter?.Write(e.Buffer, 0, e.BytesRecorded);
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                //btnRecord.Content = "Start Recording";
                //txtStatus.Text = "Ready";
            });

            if (e.Exception != null)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"Recording was interrupted by an occured error: {e.Exception.Message}", "Recording error",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
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

        protected override void OnClosed(EventArgs e)
        {
            CleanupCapture();
            base.OnClosed(e);
        }

        private void btnInfo_Click(object sender, RoutedEventArgs e)
        {
            var devices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            MMDevice d = devices.ToList()[0];
            var name = d.FriendlyName; // name
            var state = d.State; // device state
            var flow = d.DataFlow; // in/out

            var am = d.AudioEndpointVolume;
            var volume = am.MasterVolumeLevelScalar; // volume level 0-1

            var asc = d.AudioSessionManager.AudioSessionControl;
            var procName = Process.GetProcessById((int)asc.GetProcessID).ProcessName; // process name

            //tbInfo.Text = $"{name} | {state} | {flow} | {volume} | ";
            for (int i = 0; i < d.AudioSessionManager.Sessions.Count; i++)
            {
                var s = d.AudioSessionManager.Sessions[i];
                if (s.State != AudioSessionState.AudioSessionStateActive) continue;
                if (s.SimpleAudioVolume.Volume == 0) continue;
                
                //tbInfo.Text += Process.GetProcessById((int)s.GetProcessID).ProcessName + " | ";
            }
        }
    }
}
