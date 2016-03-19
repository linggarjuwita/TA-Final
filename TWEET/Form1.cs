using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TWEET
{
    public partial class Form1 : Form
    {
        DataTable dtTweet = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                OleDbConnection connExcel = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName + ";Extended Properties='Excel 8.0;HDR=Yes'");
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter adtExcel = new OleDbDataAdapter();
                cmdExcel.Connection = connExcel;

                connExcel.Open();
                DataTable schemaExcel = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                cmdExcel.CommandText = "SELECT tweet_id, name, gender, text From [" + schemaExcel.Rows[0]["TABLE_NAME"].ToString() + "]";
                adtExcel.SelectCommand = cmdExcel;
                adtExcel.Fill(dtTweet);
                connExcel.Close();

                textBox1.Text = openFileDialog1.FileName;
                dataGridView1.DataSource = dtTweet;

            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dtTweet.Rows.Count; i++)
            {
                int inDir = i + 1;
                string text = dtTweet.Rows[i][3].ToString();
                string cleanUp = Property.CleanUpTweets(text);
                string tweet = Stopword.RemoveStopwords(cleanUp);
                StreamWriter file = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "extraction\\" + inDir + ".txt");
                file.WriteLine(tweet);
                file.Close();
            }

            MessageBox.Show("Done!");
        }
    }
}
