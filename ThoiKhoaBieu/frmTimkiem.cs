using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ThoiKhoaBieu
{
    public partial class frmTimkiem : Form
    {
        public frmTimkiem()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel File|*.xls";
            if (of.ShowDialog() == DialogResult.OK)
            {
                ReadExcelContents(of.FileName);
                duongdan = of.FileName;
            }
        }

        string duongdan;
        public DataTable ReadExcelContents(string fileName)
        {
            try
            {
                OleDbConnection connection = new OleDbConnection();
                connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + fileName); //Excel 97-2003, .xls
                string excelQuery = @"Select * FROM [Sheet1$]";                   
                connection.Open();
                OleDbCommand cmd = new OleDbCommand(excelQuery, connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = cmd;
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                DataTable dt = ds.Tables[0];
                dataGridViewTKB.DataSource = dt.DefaultView;
                connection.Close();
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Chương trình không thực hiện được " + ex.Message, "Xin chú ý", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtTimkiem.Text == "")
                {
                    MessageBox.Show("Bạn chưa nhập nội dung tìm kiếm", "Thông báo lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    OleDbConnection connection = new OleDbConnection();
                    connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + duongdan); //Excel 97-2003, .xls

                    string excelQuery = @"Select * FROM [Sheet1$] where F2 like'%" + txtTimkiem.Text + "%' ";
                    //string excelQuery = @"Select * FROM [Sheet1$] where F2 like'%" + txtTimkiem.Text + "%'  AND F3 IN(SELECT F3 FROM[Sheet1$] WHERE F2 = like'%" + txtTimkiem.Text + "%'";
                    connection.Open();
                    OleDbCommand cmd = new OleDbCommand(excelQuery, connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = cmd;
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    DataTable dt = ds.Tables[0];                    
                    dataGridViewTKB.DataSource = dt.DefaultView;

                }
            }
            catch (Exception ex) { }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection();
            connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + duongdan); //Excel 97-2003, .xls            
            string excelQuery = @"Select * FROM [Sheet1$]";        
            connection.Open();
            OleDbCommand cmd = new OleDbCommand(excelQuery, connection);
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            adapter.SelectCommand = cmd;
            DataSet ds = new DataSet();
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];            
            dataGridViewTKB.DataSource = dt.DefaultView;
        }
    }
}
