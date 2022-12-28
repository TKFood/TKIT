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
            Image O_Image = Image.FromStream(WebRequest.Create("https://eip.tkfood.com.tw/UOF/Common/FileCenter/V3/Handler/FileControlHandler.ashx?id=0f7e7008-971e-49dd-a83b-987300f69baf").GetResponse().GetResponseStream());
            //将获取的图片赋给图片框
            pictureBox1.Image = O_Image;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETPIC();
        }
        #endregion
    }
}
