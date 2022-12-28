using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using FastReport;
using FastReport.Data;
using System.Net.Mail;//<-基本上發mail就用這個class
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Diagnostics;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using TKITDLL;
using System.Net.Http;
using System.Net;

namespace TKIT
{
    public partial class FrmUOFSENDMESSAGE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds1 = new DataSet();


        FileInfo info;
        string[] tempFile;
        string tFileName = "";

        public FrmUOFSENDMESSAGE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETPIC()
        {
            //Image O_Image = Image.FromStream(WebRequest.Create("https://eip.tkfood.com.tw/UOF/Common/FileCenter/V3/Handler/FileControlHandler.ashx?id=0f7e7008-971e-49dd-a83b-987300f69baf").GetResponse().GetResponseStream());
            ////将获取的图片赋给图片框
            //pictureBox1.Image = O_Image;
            //pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        public void SEARCH(string NAME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    
                                    SELECT 
                                    [TB_EIP_ALBUM_CLASS].[CLASS_NAME] AS '分類'
                                    ,[TB_EIP_ALBUM].[ALBUM_TOPIC] AS '主題'
                                    ,[PHOTO_TOPIC] AS '照片名稱'
                                    ,[TB_EIP_ALBUM_CLASS].[CLASS_GUID]
                                    ,[TB_EIP_ALBUM].[ALBUM_GUID]
                                    ,[PHOTO_GUID]
                                    ,[FILE_ID]
                                    ,[THUMBNAIL_FILE_ID]
                                    ,[PHOTO_DESC]
                                    ,[FRONT_COVER]
                                    ,[COMMEND_COUNT]
                                    ,[RESIZE_FILE_ID]
                                    FROM [UOF].[dbo].[TB_EIP_ALBUM_CLASS], [UOF].[dbo].[TB_EIP_ALBUM],[UOF].[dbo].[TB_EIP_ALBUM_PHOTO]
                                    WHERE 1=1
                                    AND [TB_EIP_ALBUM_CLASS].CLASS_GUID=[TB_EIP_ALBUM].CLASS_GUID
                                    AND [TB_EIP_ALBUM].ALBUM_GUID=[TB_EIP_ALBUM_PHOTO].ALBUM_GUID
                                    AND [CLASS_NAME] LIKE '%賀圖區%'
                                    AND [PHOTO_TOPIC] LIKE '%{0}%'
                                    ORDER BY [PHOTO_TOPIC]

                                    ", NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds1.Tables["TEMPds1"];

                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView1.Columns["分類"].Width = 100;
                    dataGridView1.Columns["主題"].Width = 100;
                    dataGridView1.Columns["照片名稱"].Width = 100;

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

            try
            {
                if (dataGridView1.CurrentRow != null)
                {
                    int rowindex = dataGridView1.CurrentRow.Index;

                    if (rowindex >= 0)
                    {
                        DataGridViewRow row = dataGridView1.Rows[rowindex];

                        string RESIZE_FILE_ID = row.Cells["RESIZE_FILE_ID"].Value.ToString();

                        Image O_Image = Image.FromStream(WebRequest.Create("https://eip.tkfood.com.tw/UOF/Common/FileCenter/V3/Handler/FileControlHandler.ashx?id="+RESIZE_FILE_ID+"").GetResponse().GetResponseStream());
                        //将获取的图片赋给图片框
                        pictureBox1.Image = O_Image;
                        pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

                    }
                    else
                    {


                    }
                }
            }
            catch
            {

            }
            
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //SETPIC();

            SEARCH(textBox1.Text);
        }
        #endregion

    
    }
}
