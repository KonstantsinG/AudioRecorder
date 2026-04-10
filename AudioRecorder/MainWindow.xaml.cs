using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;

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
            LoadAudioDevices();
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
