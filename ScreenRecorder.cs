using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using Accord.Video.FFMPEG;

namespace ScreenRecorderNameSpace
{
    class ScreenRecorder
    {
        // Video variables:
        private Rectangle bounds;
        private string outputPath = "";
        private string tempPath = "";
        private int fileCount = 1;
        private List<String> inputImageSequence = new List<string>();

        // File variables:
        private string audioName = "mic.wav";
        private string videoName = "video.mp4";
        private string finalName = "FinalVideo.mp4";


        // Audio varible:
        public static class NativeMethods
        {
            [DllImport("winmm.dll", EntryPoint = "mciSendStringA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            private static extern int mciSendString(string lpstrCommand, string lpstrReturnString, int uReturnLength, int hwndCallback);
        }

        public ScreenRecorder(Rectangle b, string outPath)
        {
            CreateTempFolder("tempScreenshot");

            bounds = b;
            outputPath = outPath;
        }

        private void CreateTempFolder(string name)
        {
            string pathName = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos) + "//" + name;
            Directory.CreateDirectory(pathName);
            tempPath = pathName;
        }

        private void DeletePath(string targetDir)
        {
            string[] files = Directory.GetFiles(targetDir);
            string[] dirs = Directory.GetDirectories(targetDir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                // DeletePath(dir);
            }
            Directory.Delete(targetDir, false);
        }

        private void DeleteFilesExcept(string targetFile, string exceptionFile)
        {
            string[] files = Directory.GetFiles(targetFile);

            foreach (string file in files)
            {
                if (file != exceptionFile)
                {
                    File.SetAttributes((string)file, FileAttributes.Normal);
                    File.Delete(file);
                }
            }
        }

        public void CleanUp()
        {
            if (Directory.Exists(tempPath))
            {
                //DeletePath(tempPath);
            }
        }

        public void RecordVideo()
        {
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
                }
                string name = tempPath + "//screenshot-" + fileCount + ".png";
                inputImageSequence.Add(name);
                fileCount++;

                bitmap.Dispose();
            }
        }

        private void SaveVideo(int width, int height, int frameRate)
        {
            using (VideoFileWriter videoFileWriter = new VideoFileWriter())
            {
                videoFileWriter.Open(outputPath + "//" + videoName, width, height, frameRate, VideoCodec.MPEG4);

                foreach (string imageLocation in inputImageSequence)
                {
                    Bitmap imageFrame = System.Drawing.Image.FromFile(imageLocation) as Bitmap;
                    videoFileWriter.WriteVideoFrame(imageFrame);
                }
                videoFileWriter.Close();
            }
        }

        public void Stop()
        {
            int width = bounds.Width;
            int height = bounds.Height;
            int frameRate = 10;

            //SaveVideo(width, height, frameRate);
            //DeletePath(tempPath);
            //DeleteFilesExcept(outputPath, outputPath + "//" + videoName);
        }
    }
}
