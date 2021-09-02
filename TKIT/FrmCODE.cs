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

namespace TKIT
{
    public partial class FrmCODE : Form
    {
        public FrmCODE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public string Encryption(string PlainText)
        {
            using (Aes aesAlg = Aes.Create())
            {
                //加密金鑰(32 Byte)
                aesAlg.Key = Encoding.Unicode.GetBytes("老楊加密老楊加密老楊加密老楊加密");
                //初始向量(Initial Vector, iv) 類似雜湊演算法中的加密鹽(16 Byte)
                aesAlg.IV = Encoding.Unicode.GetBytes("加密加密加密加密");
                //加密器
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                //執行加密
                byte[] encrypted = encryptor.TransformFinalBlock(Encoding.Unicode.GetBytes(PlainText), 0,
        Encoding.Unicode.GetBytes(PlainText).Length);

                return Convert.ToBase64String(encrypted);
            }
        }

        public string Decryption(string CipherText)
        {
            using (Aes aesAlg = Aes.Create())
            {
                //加密金鑰(32 Byte)
                aesAlg.Key = Encoding.Unicode.GetBytes("老楊加密老楊加密老楊加密老楊加密");
                //初始向量(Initial Vector, iv) 類似雜湊演算法中的加密鹽(16 Byte)
                aesAlg.IV = Encoding.Unicode.GetBytes("加密加密加密加密");
                //加密器
                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                //執行加密
                byte[] decrypted = decryptor.TransformFinalBlock(Convert.FromBase64String(CipherText), 0, Convert.FromBase64String(CipherText).Length);
                return Encoding.Unicode.GetString(decrypted);
            }
        }

        public void TestSqlConnectionStringFromConfigAndPasswordEncrpted()
        {

            //連接字串產生器
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbTKITTEST"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = Decryption(sqlsb.Password);
            sqlsb.UserID= Decryption(sqlsb.UserID);

            //簡單連線資料庫查詢SQL Server版本資料
            using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("select @@version", conn))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = cmd;
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    Console.WriteLine($"{table.Rows[0][0]}");
                }
            }
        }

        public void SETTEXTBOX()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            textBox2.Text = Encryption(textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox4.Text = Decryption(textBox3.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETTEXTBOX();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TestSqlConnectionStringFromConfigAndPasswordEncrpted();
        }
        #endregion


    }
}
