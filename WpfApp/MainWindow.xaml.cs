using System;
using System.Windows;
using System.Windows.Media.Imaging;
using AForge.Video;
using AForge.Video.DirectShow;
using ZXing;
using NAudio.Wave;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.IO;

namespace WebcamApp
{
    public partial class MainWindow : Window
    {
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        private BarcodeReader barcodeReader;
        private WaveOutEvent waveOut;
        private AudioFileReader audioFile;
        private string hasScanned;
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string ApplicationName = "QR Code Scanner";
        private static readonly string SpreadsheetId = "1I8VpcrYF8aUOpyhiKm_vZ3roGKL50lCbiEiNsKNMdrQ";
        private static readonly string sheet = "Sheet1";
        private SheetsService service;

        public MainWindow()
        {
            InitializeComponent();
            barcodeReader = new BarcodeReader();
            InitializeAudio();
            InitializeGoogleSheetsService();
        }

        private void InitializeAudio()
        {
            waveOut = new WaveOutEvent();
            audioFile = new AudioFileReader("beep.wav");
            waveOut.Init(audioFile);
        }

        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (videoSource == null || !videoSource.IsRunning)
            {
                StartWebcam();
            }
            else
            {
                StopWebcam();
            }
        }

        private void StartWebcam()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count > 0)
            {
                videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
                videoSource.NewFrame += VideoSource_NewFrame;
                videoSource.Start();
                ToggleButton.Content = "Off Webcam";
            }
        }

        private void StopWebcam()
        {
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.SignalToStop();
                videoSource.NewFrame -= VideoSource_NewFrame;
                videoSource = null;
                WebcamFeed.Source = null;
                ToggleButton.Content = "On Webcam";
            }
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            BitmapImage bi;
            using (var bitmap = (System.Drawing.Bitmap)eventArgs.Frame.Clone())
            {
                bi = new BitmapImage();
                bi.BeginInit();
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                ms.Seek(0, System.IO.SeekOrigin.Begin);
                bi.StreamSource = ms;
                bi.EndInit();

                var result = barcodeReader.Decode(bitmap);
                if (result != null)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        //ResultText.Text = result.Text;
                        if (!string.IsNullOrEmpty(hasScanned) && hasScanned.Equals(result.Text)) return;
                        PlayBeep();
                        hasScanned = result.Text;
                        SaveToGoogleSheets(hasScanned);
                        MessageBoxResult messageBoxResult = MessageBox.Show(result.Text, "", MessageBoxButton.OK);
                        //MessageBox.Show(result.Text, "", MessageBoxButton.OK);
                        if (messageBoxResult.Equals(MessageBoxResult.OK)) hasScanned = "";
                        return;
                    }));
                }
            }
            bi.Freeze();
            Dispatcher.BeginInvoke(new Action(() => WebcamFeed.Source = bi));
        }

        private void PlayBeep()
        {
            audioFile.Position = 0;
            waveOut.Play();
        }

        protected override void OnClosed(EventArgs e)
        {
            StopWebcam();
            waveOut.Dispose();
            audioFile.Dispose();
            base.OnClosed(e);
        }

        private async void InitializeGoogleSheetsService()
        {
            UserCredential credential;
            using (var stream = new FileStream("D:\\Docm\\Dev\\.learning-process\\C#\\PRN221_B3W\\webcam\\Duongtddse172132_Webcam\\WpfApp\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { SheetsService.Scope.Spreadsheets },
                    "user",
                    CancellationToken.None
                );


                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
        }

        private async Task SaveToGoogleSheets(string qrData)
        {
            var range = $"{sheet}!A:A";
            var valueRange = new ValueRange();
            var objectList = new List<object>() { qrData, DateTime.Now.ToString() };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = await appendRequest.ExecuteAsync();
        }
    }
}