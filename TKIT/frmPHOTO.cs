using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;

namespace TKIT
{
    public partial class frmPHOTO : Form
    {
        public FilterInfoCollection USB_Webcams = null;//FilterInfoCollection類別實體化
        public VideoCaptureDevice Cam;//攝像頭的初始化

        public frmPHOTO()
        {
            InitializeComponent();
        }

        public void TAKE_OPEN()
        {
            USB_Webcams = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (USB_Webcams.Count > 0)  // The quantity of WebCam must be more than 0.
            {
                button1.Enabled = true;
                Cam = new VideoCaptureDevice(USB_Webcams[0].MonikerString);

                Cam.NewFrame += Cam_NewFrame;//Press Tab  to   create
            }
            else
            {
                button1.Enabled = false;
                MessageBox.Show("No video input device is connected.");
            }
        }

        public void TAKE_CLOSE()
        {
            if (Cam != null)
            {
                if (Cam.IsRunning)  // When Form1 closes itself, WebCam must stop, too.
                {
                    Cam.Stop();   // WebCam stops capturing images.
                }
            }
        }

        void Cam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            //throw new NotImplementedException();
            pictureBox1.Image = (Bitmap)eventArgs.Frame.Clone();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TAKE_OPEN();
            try
            {
                Cam.Start();   // WebCam starts capturing images.     
            }
            catch { }       
                  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TAKE_CLOSE();
            try
            {
                Cam.Stop();  // WebCam stops capturing images.
            }
            catch { }
        }
    }
}
