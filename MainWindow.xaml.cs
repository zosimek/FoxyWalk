//------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace Microsoft.Samples.Kinect.ColorBasics
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Windows;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Threading.Tasks;
    using System.Linq;
    using System.Windows.Forms;

    using Microsoft.Kinect;
    using System.Drawing;
    using Aspose.Imaging;
    using Aspose.Imaging.FileFormats.Gif;
    using Aspose.Imaging.FileFormats.Gif.Blocks;
    using NReco.VideoConverter;

    using ScreenRecorderNameSpace;
    using System.Collections.Generic;
    using System.Linq;
    using Accord.Video.FFMPEG;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        bool folderSelected = false;
        string tempPath= string.Empty;
        string outputPath = "";
        string videoName = "video.gif";
        private List<String> inputImageSequence = new List<string>();
        private ScreenRecorder screenRecorder = new ScreenRecorder(new System.Drawing.Rectangle(), "");
        private int fileCount = 1;


        private bool recording = false;

        /// <summary>
        /// Active Kinect sensor
        /// </summary>
        private KinectSensor sensor;

        /// <summary>
        /// Bitmap that will hold color information
        /// </summary>
        private WriteableBitmap colorBitmap;

        /// <summary>
        /// Intermediate storage for the color data received from the camera
        /// </summary>
        private byte[] colorPixels;

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Execute startup tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            // Look through all sensors and start the first connected one.
            // This requires that a Kinect is connected at the time of app startup.
            // To make your app robust against plug/unplug, 
            // it is recommended to use KinectSensorChooser provided in Microsoft.Kinect.Toolkit (See components in Toolkit Browser).
            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {                  
                    this.sensor = potentialSensor;
                    break;
                }
            }

            if (null != this.sensor)
            {

                this.lblStatus.Content = this.sensor.Status;
                this.lblConnectionId.Content = this.sensor.UniqueKinectId.ToString();

                // Turn on the color stream to receive color frames
                this.sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);

                // Allocate space to put the pixels we'll receive
                this.colorPixels = new byte[this.sensor.ColorStream.FramePixelDataLength];

                // This is the bitmap we'll display on-screen
                this.colorBitmap = new WriteableBitmap(this.sensor.ColorStream.FrameWidth, this.sensor.ColorStream.FrameHeight, 96.0, 96.0, PixelFormats.Bgr32, null);

                // Set the image we display to point to the bitmap where we'll put the image data
                this.Image.Source = this.colorBitmap;

                // Add an event handler to be called whenever there is new color frame data
                this.sensor.ColorFrameReady += this.SensorColorFrameReady;

                // Start the sensor!
                try
                {
                    this.sensor.Start();
                }
                catch (IOException)
                {
                    this.sensor = null;
                }
            }

            if (null == this.sensor)
            {
                this.lblStatus.Content = "Disconnected";
                this.lblConnectionId.Content = "-";
                this.statusBarText.Text = Properties.Resources.NoKinectReady;
                this.sensor.Stop();
            }
        }

        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (null != this.sensor)
            {
                this.sensor.Stop();
            }
        }

        /// <summary>
        /// Event handler for Kinect sensor's ColorFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
            {
                if (colorFrame != null)
                {
                    // Copy the pixel data from the image to a temporary array
                    colorFrame.CopyPixelDataTo(this.colorPixels);

                    // Write the pixel data into our bitmap
                    this.colorBitmap.WritePixels(
                        new Int32Rect(0, 0, this.colorBitmap.PixelWidth, this.colorBitmap.PixelHeight),
                        this.colorPixels,
                        this.colorBitmap.PixelWidth * sizeof(int),
                        0);
                }
            }

            if (recording)
            {
                SaveVideoFrames();
            }
        }

        /// <summary>
        /// Handles the user clicking on the screenshot button
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void ButtonScreenshotClick(object sender, RoutedEventArgs e)
        {
            if (null == this.sensor)
            {
                this.statusBarText.Text = Properties.Resources.ConnectDeviceFirst;
                return;
            }

            // create a png bitmap encoder which knows how to save a .png file
            BitmapEncoder encoder = new PngBitmapEncoder();

            // create frame from the writable bitmap and add to encoder
            encoder.Frames.Add(BitmapFrame.Create(this.colorBitmap));

            string time = System.DateTime.Now.ToString("hh'-'mm'-'ss", CultureInfo.CurrentUICulture.DateTimeFormat);

            string myPhotos = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);

            string path = Path.Combine(myPhotos, "KinectSnapshot-" + time + ".png");

            // write the new file to disk
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Create))
                {
                    encoder.Save(fs);
                }

                this.statusBarText.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", Properties.Resources.ScreenshotWriteSuccess, path);
            }
            catch (IOException)
            {
                this.statusBarText.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", Properties.Resources.ScreenshotWriteFailed, path);
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (folderSelected)
            {
                recording = true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("You must select output folder before recording", "Error");
            }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            recording = false;
            fileCount = 1;

            //SaveVideo();
            SaveMovie();

        }

        private void btnDirectory_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select an Output Folder";

            if(folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputPath = folderBrowserDialog.SelectedPath;
                folderSelected = true;

                // I've may messed up with rectangle
                System.Drawing.Rectangle bounds = new System.Drawing.Rectangle(0, 0, (int)SystemParameters.FullPrimaryScreenWidth, (int)SystemParameters.FullPrimaryScreenHeight);

                screenRecorder = new ScreenRecorder(bounds, outputPath);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select a folder", "Error");
            }
        }

        private void SaveVideoFrames()
        {
            if (null == this.sensor)
            {
                this.statusBarText.Text = Properties.Resources.ConnectDeviceFirst;
                return;
            }

            // create a png bitmap encoder which knows how to save a .png file
            BitmapEncoder encoder = new PngBitmapEncoder();

            // create frame from the writable bitmap and add to encoder
            encoder.Frames.Add(BitmapFrame.Create(this.colorBitmap));

            tempPath = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos) + "//tempScreenshot";

            string path = Path.Combine(tempPath, "KinectSnapshot-" + fileCount + ".png");
            inputImageSequence.Add(path);
            fileCount++;
            // write the new file to disk
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Create))
                {
                    encoder.Save(fs);
                }

                this.statusBarText.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", Properties.Resources.ScreenshotWriteSuccess, path);
            }
            catch (IOException)
            {
                this.statusBarText.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", Properties.Resources.ScreenshotWriteFailed, path);
            }
        }

        private void SaveVideo()
        {
            int width = 640;
            int heigh = 480;
            int frameRate = 30;

            VideoFileWriter writer = new VideoFileWriter();

            writer.Open("test.avi", width, heigh, frameRate, VideoCodec.MPEG4);

            foreach (var image in inputImageSequence)
            {
                Bitmap myBitmap = new Bitmap(image);
                writer.WriteVideoFrame(myBitmap);
            }
            writer.Close();
        }

        private void SaveMovie()
        {
            var videoConv = new FFMpegConverter();
            videoConv.ConcatMedia(inputImageSequence.ToArray(), Environment.GetFolderPath(Environment.SpecialFolder.MyVideos) + "//Kinect//" + videoName, Format.gif,
                new ConcatSettings()
                {
                    VideoFrameRate = inputImageSequence.Count * 10000,
                    VideoFrameSize = "640x480",
                    VideoFrameCount= inputImageSequence.Count,
                    ConcatAudioStream= false,
                });
        }
    }
}