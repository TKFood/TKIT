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

            comboBox1load();
            textBox4.Text = "1";
        }


        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT  [RANK] ,[TITLE_NAME] FROM [UOF].[dbo].[TB_EB_JOB_TITLE] ORDER BY [RANK] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("RANK", typeof(string));
            dt.Columns.Add("TITLE_NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "RANK";
            comboBox1.DisplayMember = "TITLE_NAME";
            sqlConn.Close();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
            {
                textBox4.Text = comboBox1.SelectedValue.ToString();
            }
        }

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

        public void SEARCH2(string UOF_FORM_NAME)
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
                                    ", UOF_FORM_NAME);

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

        public void SEARCH3(string UOF_FORM_NAME)
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
                                    ,[APPLY_GROUP_ID] AS '限定申請部門代號'
                                    ,[APPLY_GROUP_NAME] AS '限定申請部門'
                                    ,[APPLY_RANKS_OPERATOR] AS '限定職級比較'
                                    ,[APPLY_RANKS] AS '限定申請職級'
                                    ,[APPLY_TITLE_NAME] AS '限定申請職級'
                                    ,[APPLY_FILEDS1] AS '限定申請欄位1'
                                    ,[APPLY_FILEDS_OPERATOR1] AS '限定申請欄位比較1'
                                    ,[APPLY_FILEDS_VALUES1] AS '限定申請欄位值1'
                                    ,[APPLY_FILEDS2] AS '限定申請欄位2'
                                    ,[APPLY_FILEDS_OPERATOR2] AS '限定申請欄位比較2'
                                    ,[APPLY_FILEDS_VALUES2] AS '限定申請欄位值2'
                                    ,[APPLY_FILEDS3] AS '限定申請欄位3'
                                    ,[APPLY_FILEDS_OPERATOR3] AS '限定申請欄位比較3'
                                    ,[APPLY_FILEDS_VALUES3] AS '限定申請欄位值3'
                                    ,[APPLY_FILEDS4] AS '限定申請欄位4'
                                    ,[APPLY_FILEDS_OPERATOR4] AS '限定申請欄位比較4'
                                    ,[APPLY_FILEDS_VALUES4] AS '限定申請欄位值4'
                                    ,[APPLY_FILEDS5] AS '限定申請欄位5'
                                    ,[APPLY_FILEDS_OPERATOR5] AS '限定申請欄位比較5'
                                    ,[APPLY_FILEDS_VALUES5] AS '限定申請欄位值5'
                                    ,[SET_FLOW_RANKS] AS '表單簽核的職級代號'
                                    ,[SET_FLOW_TITLE_NAME] AS '表單簽核的職'
                                    ,[PRIORITYS] AS '條件優先權'
                                    ,[ISUSED] AS '是否使用'
                                    ,[COMMENTS] AS '備註'
                                    ,[ID]
                                    FROM [UOF].[dbo].[Z_UOF_FROM_CONDITIONS]
                                    WHERE [UOF_FORM_NAME]='{0}'
                                    ORDER BY [PRIORITYS]
                                    ", UOF_FORM_NAME);

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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            textBox5.Text = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox5.Text = row.Cells["ID"].Value.ToString();


                }
                else
                {
                    textBox5.Text = null;
                  

                }
            }
        }

        public void ADD_UOF_Z_UOF_FORM_DEFALUT_SINGERS(string UOF_FORM_NAME,string RANKS,string TITLE_NAME)
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(UOF_FORM_NAME))
                    {
                        StringBuilder SBSQL = new StringBuilder();
                        SBSQL.AppendFormat(@"   
                                            INSERT INTO  [UOF].[dbo].[Z_UOF_FORM_DEFALUT_SINGERS]
                                            (
                                            [UOF_FORM_NAME]
                                            ,[RANKS]
                                            ,[TITLE_NAME]
                                            )
                                            VALUES
                                            (
                                            @UOF_FORM_NAME
                                            ,@RANKS
                                            ,@TITLE_NAME
                                            )
                                       
                                            ");

                        string sql = SBSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@UOF_FORM_NAME", UOF_FORM_NAME);
                            cmd.Parameters.AddWithValue("@RANKS", RANKS);
                            cmd.Parameters.AddWithValue("@TITLE_NAME", TITLE_NAME);


                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("失敗");
            }
        }
        public void DELETE_UOF_Z_UOF_FORM_DEFALUT_SINGERS(string UOF_FORM_NAME)
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(UOF_FORM_NAME))
                    {
                        StringBuilder SBSQL = new StringBuilder();
                        SBSQL.AppendFormat(@"   
                                            DELETE  [UOF].[dbo].[Z_UOF_FORM_DEFALUT_SINGERS]
                                            WHERE [UOF_FORM_NAME]=@UOF_FORM_NAME
                                       
                                             ");

                        string sql = SBSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@UOF_FORM_NAME", UOF_FORM_NAME);




                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("失敗");
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
        private void button4_Click(object sender, EventArgs e)
        {
            ADD_UOF_Z_UOF_FORM_DEFALUT_SINGERS(textBox2.Text, textBox4.Text,comboBox1.Text.ToString());

            SEARCH2(textBox2.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
           
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETE_UOF_Z_UOF_FORM_DEFALUT_SINGERS(textBox3.Text);

                SEARCH2(textBox5.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        #endregion

       
    }
}
