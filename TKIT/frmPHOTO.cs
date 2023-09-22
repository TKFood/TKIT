using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;
using AForge.Video;
using AForge.Video.DirectShow;
using System.IO;



namespace TKIT
{
    public partial class frmPHOTO : Form
    {
        StringBuilder sbSql = new StringBuilder();
        SqlConnection sqlConn = new SqlConnection();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;

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

        //保存图片
        private delegate void SaveImage();
        private void SaveImageHH(string ImagePath)
        {
            if (this.pictureBox1.InvokeRequired)
            {
                SaveImage saveimage = delegate { this.pictureBox1.Image.Save(ImagePath); };
                this.pictureBox1.Invoke(saveimage);
            }
            else
            {
                this.pictureBox1.Image.Save(ImagePath);
            }

        }

        // 將 PictureBox 中的圖片轉換為位元組數組
        private byte[] ImageToByteArray(Image image)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg); // 或者使用其他圖像格式
                return ms.ToArray();
            }
        }

        // 將位元組數組插入到資料庫的 BLOB 欄位中
        private void InsertImageIntoDatabase(string NO,string CTIMES, byte[] imageBytes)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                     INSERT INTO [TKWAREHOUSE].[dbo].[PACKAGEBOXSPHOTO]
                                    ([NO], [CTIMES], [PHOTOS])
                                    VALUES
                                    (@NO, @CTIMES, @PHOTOS)
                                    "
                                    );

                cmd.Parameters.AddWithValue("@NO", NO);
                cmd.Parameters.AddWithValue("@CTIMES", CTIMES);
                cmd.Parameters.AddWithValue("@PHOTOS", imageBytes);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    //MessageBox.Show("圖片已成功存儲到資料庫。");

                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        // 將 PictureBox 中的圖片存儲到資料庫
        private void SaveImageToDatabase()
        {
            // 替換為您的 PictureBox 控制項名稱
            Image image = pictureBox1.Image;

            if (image != null)
            {
                byte[] imageBytes = ImageToByteArray(image);
                InsertImageIntoDatabase(DateTime.Now.ToString("yyyyMMdd"), DateTime.Now.ToString("yyyyMMdd HH:MM:ss"), imageBytes);
              
            }
            else
            {
              
            }
        }

        private void DisplayImageFromFolder(string folderPath)
        {
            // 檢查資料夾是否存在
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("資料夾不存在。");
                return;
            }

            // 獲取資料夾中的所有圖片檔案
            string[] imageFiles = Directory.GetFiles(folderPath, "*.jpg"); // 只顯示 .jpg 檔案，您可以根據需要更改擴展名

            if (imageFiles.Length > 0)
            {
                // 選擇第一張圖片顯示
                string imagePath = imageFiles[0];

                // 顯示圖片在 PictureBox 控制項上
                pictureBox1.Image = Image.FromFile(imagePath);
            }
            else
            {
                // 如果沒有圖片，清除 PictureBox
                pictureBox1.Image = null;
                MessageBox.Show("沒有找到圖片。");
            }
        }


        //

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

        private void button3_Click(object sender, EventArgs e)
        {
            //string imagePath = System.Environment.CurrentDirectory;
            string imagePath = Path.Combine(Environment.CurrentDirectory, "Images",DateTime.Now.ToString("yyyy"));
            if (!Directory.Exists(imagePath))
            {
                Directory.CreateDirectory(imagePath);
            }
            SaveImageHH(imagePath+"\\"+ DateTime.Now.ToString("yyyyMMddHHmmss") + ".jpg");
            SaveImageToDatabase();

            MessageBox.Show("OK");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string imagePath = Path.Combine(Environment.CurrentDirectory, "Images", DateTime.Now.ToString("yyyy"));
            DisplayImageFromFolder(imagePath);
        }

        
}
}
