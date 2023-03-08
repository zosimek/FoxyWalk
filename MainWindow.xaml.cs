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
    using Accord.Video.FFMPEG;
    using IronXL;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        List<float> left_wrist_X = new List<float>();
        List<float> left_wrist_Y = new List<float>();
        List<float> left_wrist_Z = new List<float>();
        List<float> right_wrist_X = new List<float>();
        List<float> right_wrist_Y = new List<float>();
        List<float> right_wrist_Z = new List<float>();

        List<float> left_elbow_X = new List<float>();
        List<float> left_elbow_Y = new List<float>();
        List<float> left_elbow_Z = new List<float>();
        List<float> right_elbow_X = new List<float>();
        List<float> right_elbow_Y = new List<float>();
        List<float> right_elbow_Z = new List<float>();

        List<float> left_shoulder_X = new List<float>();
        List<float> left_shoulder_Y = new List<float>();
        List<float> left_shoulder_Z = new List<float>();
        List<float> right_shoulder_X = new List<float>();
        List<float> right_shoulder_Y = new List<float>();
        List<float> right_shoulder_Z = new List<float>();

        List<float> left_ankle_X = new List<float>();
        List<float> left_ankle_Y = new List<float>();
        List<float> left_ankle_Z = new List<float>();
        List<float> right_ankle_X = new List<float>();
        List<float> right_ankle_Y = new List<float>();
        List<float> right_ankle_Z = new List<float>();

        List<float> left_knee_X = new List<float>();
        List<float> left_knee_Y = new List<float>();
        List<float> left_knee_Z = new List<float>();
        List<float> right_knee_X = new List<float>();
        List<float> right_knee_Y = new List<float>();
        List<float> right_knee_Z = new List<float>();

        List<float> left_hip_X = new List<float>();
        List<float> left_hip_Y = new List<float>();
        List<float> left_hip_Z = new List<float>();
        List<float> right_hip_X = new List<float>();
        List<float> right_hip_Y = new List<float>();
        List<float> right_hip_Z = new List<float>();

        List<float> center_hip_X = new List<float>();
        List<float> center_hip_Y = new List<float>();
        List<float> center_hip_Z = new List<float>();
        List<float> spine_X = new List<float>();
        List<float> spine_Y = new List<float>();
        List<float> spine_Z = new List<float>();
        List<float> center_shoulder_X = new List<float>();
        List<float> center_shoulder_Y = new List<float>();
        List<float> center_shoulder_Z = new List<float>();

        bool folderSelected = false;
        string tempPath= string.Empty;
        string outputPath = "";
        string videoName = "video.gif";
        string excelName = "excel.xlsx";
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


                /////////////////////////////   SKELETON /////////////////////////////////
                
                // Turn on the skeleton stream to receive skeleton frames
                this.sensor.SkeletonStream.Enable();
                // Add an event handler to be called whenever there is new color frame data
                this.sensor.SkeletonFrameReady += this.SensorSkeletonFrameReady;


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
               // this.sensor.Stop();
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
        /// Event handler for Kinect sensor's SkeletonFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            Skeleton[] skeletons = new Skeleton[0];

            //////////////////// JOINT DECLARATION //////////////////
            Joint left_wrist = new Joint();
            Joint right_wrist = new Joint();

            Joint left_elbow = new Joint();
            Joint right_elbow = new Joint();

            Joint left_shoulder = new Joint();
            Joint right_shoulder = new Joint();

            Joint left_ankle = new Joint();
            Joint right_ankle = new Joint();

            Joint left_knee = new Joint();
            Joint right_knee = new Joint();

            Joint left_hip = new Joint();
            Joint right_hip = new Joint();

            Joint center_hip = new Joint();
            Joint spine = new Joint();
            Joint center_shoulder= new Joint();

            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(skeletons);
                }
            }

            if (skeletons.Length != 0)
            {
                foreach (Skeleton skel in skeletons)
                {
                    if (skel.TrackingState == SkeletonTrackingState.Tracked)
                    {
                        left_ankle = skel.Joints[JointType.AnkleLeft];
                        left_ankle_X.Add(left_ankle.Position.X);
                        left_ankle_Y.Add(left_ankle.Position.Y);
                        left_ankle_Z.Add(left_ankle.Position.Z);

                        right_ankle = skel.Joints[JointType.AnkleRight];
                        right_ankle_X.Add(right_ankle.Position.X);
                        right_ankle_Y.Add(right_ankle.Position.Y);
                        right_ankle_Z.Add(right_ankle.Position.Z);

                        left_knee = skel.Joints[JointType.KneeLeft];
                        left_knee_X.Add(left_knee.Position.X);
                        left_knee_Y.Add(left_knee.Position.Y);
                        left_knee_Z.Add(left_knee.Position.Z);

                        right_knee = skel.Joints[JointType.KneeRight];
                        right_knee_X.Add(right_knee.Position.X);
                        right_knee_Y.Add(right_knee.Position.Y);
                        right_knee_Z.Add(right_knee.Position.Z);

                        left_hip = skel.Joints[JointType.HipLeft];
                        left_hip_X.Add(left_hip.Position.X);
                        left_hip_Y.Add(left_hip.Position.Y);
                        left_hip_Z.Add(left_hip.Position.Z);

                        right_hip = skel.Joints[JointType.HipRight];
                        right_hip_X.Add(right_hip.Position.X);
                        right_hip_Y.Add(right_hip.Position.Y);
                        right_hip_Z.Add(right_hip.Position.Z);

                        left_wrist = skel.Joints[JointType.WristLeft];
                        left_wrist_X.Add(left_wrist.Position.X);
                        left_wrist_Y.Add(left_wrist.Position.Y);
                        left_wrist_Z.Add(left_wrist.Position.Z);

                        right_wrist = skel.Joints[JointType.WristRight];
                        right_wrist_X.Add(right_wrist.Position.X);
                        right_wrist_Y.Add(right_wrist.Position.Y);
                        right_wrist_Z.Add(right_wrist.Position.Z);

                        left_elbow = skel.Joints[JointType.ElbowLeft];
                        left_elbow_X.Add(left_elbow.Position.X);
                        left_elbow_Y.Add(left_elbow.Position.Y);
                        left_elbow_Z.Add(left_elbow.Position.Z);

                        right_elbow = skel.Joints[JointType.ElbowRight];
                        right_elbow_X.Add(right_elbow.Position.X);
                        right_elbow_Y.Add(right_elbow.Position.Y);
                        right_elbow_Z.Add(right_elbow.Position.Z);

                        left_shoulder = skel.Joints[JointType.ShoulderLeft];
                        left_shoulder_X.Add(left_shoulder.Position.X);
                        left_shoulder_Y.Add(left_shoulder.Position.Y);
                        left_shoulder_Z.Add(left_shoulder.Position.Z);

                        right_shoulder = skel.Joints[JointType.ShoulderRight];
                        right_shoulder_X.Add(right_shoulder.Position.X);
                        right_shoulder_Y.Add(right_shoulder.Position.Y);
                        right_shoulder_Z.Add(right_shoulder.Position.Z);

                        center_hip = skel.Joints[JointType.HipCenter];
                        center_hip_X.Add(center_hip.Position.X);
                        center_hip_Y.Add(center_hip.Position.Y);
                        center_hip_Z.Add(center_hip.Position.Z);

                        spine = skel.Joints[JointType.Spine];
                        spine_X.Add(spine.Position.X);
                        spine_Y.Add(spine.Position.Y);
                        spine_Z.Add(spine.Position.Z);

                        center_shoulder = skel.Joints[JointType.ShoulderCenter];
                        center_shoulder_X.Add(center_shoulder.Position.X);
                        center_shoulder_Y.Add(center_shoulder.Position.Y);
                        center_shoulder_Z.Add(center_shoulder.Position.Z);


                    }
                }
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
            SaveExcel();

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

        private void SaveExcel()
        {
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var left_wrist_sheet = workbook.CreateWorkSheet("wrist_left");
            var right_wrist_sheet = workbook.CreateWorkSheet("wrist_right");
            var left_elbow_sheet = workbook.CreateWorkSheet("elbow_left");
            var right_elbow_sheet = workbook.CreateWorkSheet("elbow_right");
            var left_shoulder_sheet = workbook.CreateWorkSheet("shoulder_left");
            var right_shoulder_sheet = workbook.CreateWorkSheet("shoulder_right");
            var center_shoulder_sheet = workbook.CreateWorkSheet("shoulder_center");
            var left_ankle_sheet = workbook.CreateWorkSheet("ankle_left");
            var right_ankle_sheet = workbook.CreateWorkSheet("ankle_right");
            var left_knee_sheet = workbook.CreateWorkSheet("knee_left");
            var right_knee_sheet = workbook.CreateWorkSheet("knee_right");
            var left_hip_sheet = workbook.CreateWorkSheet("hip_left");
            var right_hip_sheet = workbook.CreateWorkSheet("hip_right");
            var center_hip_sheet = workbook.CreateWorkSheet("hip_center");
            var spine_sheet = workbook.CreateWorkSheet("spine");


            left_wrist_sheet["A1"].Value = "X";
            right_wrist_sheet["A1"].Value = "X";

            left_elbow_sheet["A1"].Value = "X";
            right_elbow_sheet["A1"].Value = "X";

            left_shoulder_sheet["A1"].Value = "X";
            right_shoulder_sheet["A1"].Value = "X";
            center_shoulder_sheet["A1"].Value = "X";

            left_ankle_sheet["A1"].Value = "X";
            right_ankle_sheet["A1"].Value = "X";

            left_knee_sheet["A1"].Value = "X";
            right_knee_sheet["A1"].Value = "X";

            left_hip_sheet["A1"].Value = "X";
            right_hip_sheet["A1"].Value = "X";
            center_hip_sheet["A1"].Value = "X";
            spine_sheet["A1"].Value = "X";


            /////////////////// Y  /////////////////////
            left_wrist_sheet["B1"].Value = "Y";
            right_wrist_sheet["B1"].Value = "Y";

            left_elbow_sheet["B1"].Value = "Y";
            right_elbow_sheet["B1"].Value = "Y";

            left_shoulder_sheet["B1"].Value = "Y";
            right_shoulder_sheet["B1"].Value = "Y";
            center_shoulder_sheet["B1"].Value = "Y";

            left_ankle_sheet["B1"].Value = "Y";
            right_ankle_sheet["B1"].Value = "Y";

            left_knee_sheet["B1"].Value = "Y";
            right_knee_sheet["B1"].Value = "Y";

            left_hip_sheet["B1"].Value = "Y";
            right_hip_sheet["B1"].Value = "Y";
            center_hip_sheet["B1"].Value = "Y";
            spine_sheet["B1"].Value = "Y";

            //////////////////  Z  //////////////////////

            left_wrist_sheet["C1"].Value = "Z";
            right_wrist_sheet["C1"].Value = "Z";

            left_elbow_sheet["C1"].Value = "Z";
            right_elbow_sheet["C1"].Value = "Z";

            left_shoulder_sheet["C1"].Value = "Z";
            right_shoulder_sheet["C1"].Value = "Z";
            center_shoulder_sheet["C1"].Value = "Z";

            left_ankle_sheet["C1"].Value = "Z";
            right_ankle_sheet["C1"].Value = "Z";

            left_knee_sheet["C1"].Value = "Z";
            right_knee_sheet["C1"].Value = "Z";

            left_hip_sheet["C1"].Value = "Z";
            right_hip_sheet["C1"].Value = "Z";
            center_hip_sheet["C1"].Value = "Z";
            spine_sheet["C1"].Value = "Z";

            var i = 2;

            //////////// left ankle  ////////////////
            foreach (float element in left_ankle_X)
            {
                left_ankle_sheet["A" + i.ToString()].Value= element;
                i++;
            }
            i = 2;
            foreach (float element in left_ankle_Y)
            {
                left_ankle_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_ankle_Z)
            {
                left_ankle_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right ankle  ////////////////
            i = 2;
            foreach (float element in right_ankle_X)
            {
                right_ankle_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_ankle_Y)
            {
                right_ankle_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_ankle_Z)
            {
                right_ankle_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// left knee  ////////////////
            i = 2;
            foreach (float element in left_knee_X)
            {
                left_knee_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_knee_Y)
            {
                left_knee_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_knee_Z)
            {
                left_knee_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right knee  ////////////////
            i = 2;
            foreach (float element in right_knee_X)
            {
                right_knee_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_knee_Y)
            {
                right_knee_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_knee_Z)
            {
                right_knee_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// left hip  ////////////////
            i = 2;
            foreach (float element in left_hip_X)
            {
                left_hip_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_hip_Y)
            {
                left_hip_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_hip_Z)
            {
                left_hip_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right hip  ////////////////
            i = 2;
            foreach (float element in right_hip_X)
            {
                right_hip_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_hip_Y)
            {
                right_hip_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_hip_Z)
            {
                right_hip_sheet["C" + i.ToString()].Value = element;
                i++;
            }
            //////////// center hip  ////////////////
            i = 2;
            foreach (float element in center_hip_X)
            {
                center_hip_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in center_hip_Y)
            {
                center_hip_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in center_hip_Z)
            {
                center_hip_sheet["C" + i.ToString()].Value = element;
                i++;
            }
            //////////// spine  ////////////////
            i = 2;
            foreach (float element in spine_X)
            {
                spine_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in spine_Y)
            {
                spine_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in spine_Z)
            {
                spine_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// left wrist  ////////////////
            i = 2;
            foreach (float element in left_wrist_X)
            {
                left_wrist_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_wrist_Y)
            {
                left_wrist_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_wrist_Z)
            {
                left_wrist_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right wrist  ////////////////
            i = 2;
            foreach (float element in right_wrist_X)
            {
                right_wrist_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_wrist_Y)
            {
                right_wrist_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_wrist_Z)
            {
                right_wrist_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// left elbow  ////////////////
            i = 2;
            foreach (float element in left_elbow_X)
            {
                left_elbow_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_elbow_Y)
            {
                left_elbow_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_elbow_Z)
            {
                left_elbow_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right elbow  ////////////////
            i = 2;
            foreach (float element in right_elbow_X)
            {
                right_elbow_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_elbow_Y)
            {
                right_elbow_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_elbow_Z)
            {
                right_elbow_sheet["C" + i.ToString()].Value = element;
                i++;
            }
            //////////// left shoulder  ////////////////
            i = 2;
            foreach (float element in left_shoulder_X)
            {
                left_shoulder_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_shoulder_Y)
            {
                left_shoulder_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in left_shoulder_Z)
            {
                left_shoulder_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// right shoulder  ////////////////
            i = 2;
            foreach (float element in right_shoulder_X)
            {
                right_shoulder_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_shoulder_Y)
            {
                right_shoulder_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in right_shoulder_Z)
            {
                right_shoulder_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            //////////// center shoulder  ////////////////
            i = 2;
            foreach (float element in center_shoulder_X)
            {
                center_shoulder_sheet["A" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in center_shoulder_Y)
            {
                center_shoulder_sheet["B" + i.ToString()].Value = element;
                i++;
            }
            i = 2;
            foreach (float element in center_shoulder_Z)
            {
                center_shoulder_sheet["C" + i.ToString()].Value = element;
                i++;
            }

            workbook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.MyVideos) + "//Kinect//" + excelName);
        }
    }
}