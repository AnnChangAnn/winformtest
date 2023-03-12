using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data.SQLite;

namespace MISSampleApp
{
    public partial class Form1 : Form
    {
        public SQLiteConnection sqliteconn;

        public Form1()
        {
            InitializeComponent();
            sqliteconn = DBConnection();
        }

        public SQLiteConnection DBConnection()
        {
            //建立DB connection
            SQLiteConnection sqliteconn = new SQLiteConnection("Data Source=..\\..\\..\\database.db;Version=3;New=False;Compress=True;");
            sqliteconn.Open();
            return sqliteconn;
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            // Configure open file dialog box
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox_FilePath.Text = openFileDialog.FileName;
            }
        }

        private void buttonGetData_Click(object sender, EventArgs e)
        {
            // Import Excel to DataGridview 
            //建立OLEDB連線
            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0;", textBox_FilePath.Text.Trim());

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

                OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", sheetName), conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
        }        

        
        private void button_InsertDB_Click(object sender, EventArgs e)
        {
            // Insert DataGridview data to SQLite DB

            //滑鼠等待
            Cursor.Current = Cursors.WaitCursor;

            //清空table資料 for testing
            SQLiteCommand ResetTable = new SQLiteCommand("DELETE FROM SampleTable", sqliteconn);
            ResetTable.ExecuteNonQuery();

            //逐筆資料insert into table
            for (int RowCount = 0; RowCount < dataGridView1.Rows.Count - 1; RowCount++)
            {
                string myinsertSQL = "INSERT INTO SampleTable(Office, Item, Company, SaleQ4, SaleQ3, SaleQ2, SaleQ1, ReqDate, OrderNo,Currenry, Price, MSP) " +
                    "VALUES (@Office, @Item, @Company, @SaleQ4, @SaleQ3, @SaleQ2, @SaleQ1, @ReqDate, @OrderNo, @Currenry, @Price, @MSP)";
                SQLiteCommand insertSQL = new SQLiteCommand(myinsertSQL, sqliteconn);
                insertSQL.Parameters.AddWithValue("@Office", dataGridView1.Rows[RowCount].Cells[0].Value.ToString());
                insertSQL.Parameters.AddWithValue("@Item", dataGridView1.Rows[RowCount].Cells[1].Value.ToString());
                insertSQL.Parameters.AddWithValue("@Company", dataGridView1.Rows[RowCount].Cells[2].Value.ToString());
                insertSQL.Parameters.AddWithValue("@SaleQ4", dataGridView1.Rows[RowCount].Cells[3].Value.ToString());
                insertSQL.Parameters.AddWithValue("@SaleQ3", dataGridView1.Rows[RowCount].Cells[4].Value.ToString());
                insertSQL.Parameters.AddWithValue("@SaleQ2", dataGridView1.Rows[RowCount].Cells[5].Value.ToString());
                insertSQL.Parameters.AddWithValue("@SaleQ1", dataGridView1.Rows[RowCount].Cells[6].Value.ToString());
                insertSQL.Parameters.AddWithValue("@ReqDate", dataGridView1.Rows[RowCount].Cells[7].Value.ToString());
                insertSQL.Parameters.AddWithValue("@OrderNo", dataGridView1.Rows[RowCount].Cells[8].Value.ToString());
                insertSQL.Parameters.AddWithValue("@Currenry", dataGridView1.Rows[RowCount].Cells[9].Value.ToString());
                insertSQL.Parameters.AddWithValue("@Price", dataGridView1.Rows[RowCount].Cells[10].Value.ToString());
                insertSQL.Parameters.AddWithValue("@MSP", dataGridView1.Rows[RowCount].Cells[11].Value.ToString());

                insertSQL.ExecuteNonQuery();
            }

            //滑鼠恢復
            Cursor.Current = Cursors.Default;
        }

        private void button_SQLQuery_Click(object sender, EventArgs e)
        {
            // SQLQuery       
            string myQuerySQL = "SELECT Office, Item ,OrderNo, Price, MSP FROM SampleTable a WHERE a.MSP > a.Price order by MSP";
            SQLiteCommand QuerySQL = new SQLiteCommand(sqliteconn);

            QuerySQL.CommandText = myQuerySQL;
            SQLiteDataReader SQLreader = QuerySQL.ExecuteReader();

            //建立DataTable來接收資料
            DataTable UpdateQuery = new DataTable();
            UpdateQuery.Load(SQLreader);

            //將DataTable中的資料更新到gridview1內
            dataGridView1.DataSource = UpdateQuery;
        }
    }
}
