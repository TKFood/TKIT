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
    public partial class FrmSETFORMFLOW : Form
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

        public FrmSETFORMFLOW()
        {
            InitializeComponent();
        }


        #region FUNCTION

        public void SEARCH(string FORM_NAME)
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
                                    [CATEGORY_NAME] AS '表單分類'
                                    ,[FORM_NAME] AS '表單名稱'
                                    ,[FORM_ID]

                                    FROM [UOF].[dbo].[TB_WKF_FORM],[UOF].[dbo].[TB_WKF_FORM_CATEGORY]
                                    WHERE  [TB_WKF_FORM].CATEGORY_ID=[TB_WKF_FORM_CATEGORY].CATEGORY_ID
                                    AND [FORM_CTL]=0
                                    AND [FORM_NAME] LIKE '%{0}%'
                                    ORDER BY [CATEGORY_NAME],[FORM_NAME]
                                    ", FORM_NAME);

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
                    dataGridView1.Columns["表單分類"].Width = 100;
                    dataGridView1.Columns["表單名稱"].Width = 200;


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
            textBox2.Text = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                   
                    textBox2.Text = row.Cells["表單名稱"].Value.ToString();
                    textBox3.Text = row.Cells["表單名稱"].Value.ToString();

                    SEARCH2(textBox2.Text);
                    SEARCH3(textBox3.Text);

                }
                else
                {
                    textBox2.Text = null;
                    textBox3.Text = null;

                }
            }
        }

        public void SEARCH2(string FORM_NAME)
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
                                    [UOF_FORM_NAME] AS '表單名稱'
                                    ,[RANKS] AS '預設簽核職級'
                                    ,[TITLE_NAME] AS '預設簽核職級名稱'
                                    ,[ID]
                                    FROM [UOF].[dbo].[Z_UOF_FORM_DEFALUT_SINGERS]
                                    WHERE [UOF_FORM_NAME]='{0}'
                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView2.Columns["表單名稱"].Width = 200;
                    dataGridView2.Columns["預設簽核職級"].Width =200;
                    dataGridView2.Columns["預設簽核職級名稱"].Width = 200;

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCH3(string FORM_NAME)
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
                                     [ID]
                                    ,[UOF_FORM_NAME]
                                    ,[APPLY_GROUP_ID]
                                    ,[APPLY_GROUP_NAME]
                                    ,[APPLY_RANKS_OPERATOR]
                                    ,[APPLY_RANKS]
                                    ,[APPLY_TITLE_NAME]
                                    ,[APPLY_FILEDS1]
                                    ,[APPLY_FILEDS_OPERATOR1]
                                    ,[APPLY_FILEDS_VALUES1]
                                    ,[APPLY_FILEDS2]
                                    ,[APPLY_FILEDS_OPERATOR2]
                                    ,[APPLY_FILEDS_VALUES2]
                                    ,[APPLY_FILEDS3]
                                    ,[APPLY_FILEDS_OPERATOR3]
                                    ,[APPLY_FILEDS_VALUES3]
                                    ,[APPLY_FILEDS4]
                                    ,[APPLY_FILEDS_OPERATOR4]
                                    ,[APPLY_FILEDS_VALUES4]
                                    ,[APPLY_FILEDS5]
                                    ,[APPLY_FILEDS_OPERATOR5]
                                    ,[APPLY_FILEDS_VALUES5]
                                    ,[SET_FLOW_RANKS]
                                    ,[SET_FLOW_TITLE_NAME]
                                    ,[PRIORITYS]
                                    ,[ISUSED]
                                    ,[COMMENTS]
                                    FROM [UOF].[dbo].[Z_UOF_FROM_CONDITIONS]
                                    WHERE [UOF_FORM_NAME]='{0}'
                                    ORDER BY [PRIORITYS]
                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds1.Tables["TEMPds1"];
                 

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox1.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH2(textBox2.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH3(textBox3.Text);
        }
        #endregion


    }
}
