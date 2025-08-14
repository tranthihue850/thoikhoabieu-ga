using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ThoiKhoaBieu
{
    public partial class frmMain : Form
    {
        public RWDataExcel data;
        public ClassQuanThe lichSang;
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnOpenData_Click(object sender, EventArgs e)
        {
            OpenFileDialog save = new OpenFileDialog();
            save.Title = "Chose file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                data = new RWDataExcel();
                data.ReadAllDataInExcelFile(save.FileName);
                data.Dispose();
                this.txtsolop.Text = data.soLop.ToString();
                this.txtsomon.Text = data.soMon.ToString();
                this.txtsogv.Text = data.soGv.ToString();
                ReadFile(save.FileName);
                //Properti_Draw();
            }
            else
            {

            }
        }

        private void ReadFile(string pathFile)
        {
            data = new RWDataExcel();
            data.ReadAllDataInExcelFile(pathFile);
            int rong = 776; // bar5.Size.Width - 20; ////lay kich thuoc cua bar
            this.lblMain.Text = pathFile;
            // Properti_Draw();
            // xoa sach checklistbox
            while (checkedListBox1.Items.Count > 0)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; ++i)
                    checkedListBox1.Items.RemoveAt(i);
            }
            for (int i = 0; i < data.soLop; ++i)
                this.checkedListBox1.Items.Add(data.lop[i]);

            #region  hien thi class
            // xoa sach nhung j nhin thay
            DeleteDataView(dataClass);
            //// nhap du lieu vao dataGridView
            // tieu de
            DataGridViewColumn newCol = new DataGridViewColumn(); // add a column to the grid         
            DataGridViewCell newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Tên lớp"; //"Class";
            newCol.Visible = true;
            dataClass.Columns.Add(newCol);
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Mã lớp";//" Class Code";
            newCol.Visible = true;
            dataClass.Columns.Add(newCol);
            // du lieu
            for (int i = 0; i < data.soLop; ++i)
            {
                dataClass.Rows.Add();// them dong moi                   
                dataClass[1, i].Value = (i).ToString();
                dataClass[0, i].Value = data.lop[i];
            }
            #endregion

            //#region  hien thi mon
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataSub);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Môn học"; // "Subjects";
            newCol.Visible = true;
            dataSub.Columns.Add(newCol);
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Mã môn học"; // " Subjects Code";
            newCol.Visible = true;
            dataSub.Columns.Add(newCol);
            // du lieu
            for (int i = 0; i < data.soMon; ++i)
            {
                dataSub.Rows.Add();// them dong moi                   
                dataSub[1, i].Value = (i).ToString();
                dataSub[0, i].Value = data.mon[i];
            }
            //#endregion

            //#region  hien thi giao vien
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataTea);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Giáo viên"; //"Teachers";
            newCol.Visible = true;
            dataTea.Columns.Add(newCol);
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Mã giáo viên";  // " Teachers Code";
            newCol.Visible = true;
            dataTea.Columns.Add(newCol);
            // du lieu
            for (int i = 0; i < data.soGv; ++i)
            {
                dataTea.Rows.Add();// them dong moi                   
                dataTea[1, i].Value = (i).ToString();
                dataTea[0, i].Value = data.giangVien[i];
            }
            //#endregion

            //#region  hien thi Lop - Mon
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataCS);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Cla- Sub ";
            newCol.Visible = true;
            dataCS.Columns.Add(newCol);
            for (int i = 0; i < data.soMon; ++i)
            {
                newCol = new DataGridViewColumn(); // add a column to the grid         
                newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
                newCol.CellTemplate = newCell;
                newCol.HeaderText = data.mon[i];
                newCol.Visible = true;
                dataCS.Columns.Add(newCol);
            }
            // du lieu
            for (int i = 0; i < data.soLop; ++i)
            {
                dataCS.Rows.Add();// them dong moi
                dataCS[0, i].Value = data.lop[i];
                for (int j = 0; j < data.soMon; ++j)
                {
                    dataCS[j + 1, i].Value = data.qhLopMon[i, j].ToString();
                }
            }
            //#endregion

            //#region  hien thi Lop - Giang vien
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataCT);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = "Cla- Tea ";
            newCol.Visible = true;
            dataCT.Columns.Add(newCol);
            for (int i = 0; i < data.soGv; ++i)
            {
                newCol = new DataGridViewColumn(); // add a column to the grid         
                newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
                newCol.CellTemplate = newCell;
                newCol.HeaderText = data.giangVien[i];
                newCol.Visible = true;
                dataCT.Columns.Add(newCol);
            }
            // du lieu
            for (int i = 0; i < data.soLop; ++i)
            {
                dataCT.Rows.Add();// them dong moi
                dataCT[0, i].Value = data.lop[i];
                for (int j = 0; j < data.soGv; ++j)
                {
                    dataCT[j + 1, i].Value = data.qhLopGiangVien[i, j].ToString();
                }
            }
            //#endregion

            //#region  hien thi Giang vien- Mon
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataTS);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = " Tea - Sub ";
            newCol.Visible = true;
            dataTS.Columns.Add(newCol);
            for (int i = 0; i < data.soMon; ++i)
            {
                newCol = new DataGridViewColumn(); // add a column to the grid         
                newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
                newCol.CellTemplate = newCell;
                newCol.HeaderText = data.mon[i];
                newCol.Visible = true;
                dataTS.Columns.Add(newCol);
            }
            // du lieu
            for (int i = 0; i < data.soGv; ++i)
            {
                dataTS.Rows.Add();// them dong moi
                dataTS[0, i].Value = data.giangVien[i];
                for (int j = 0; j < data.soMon; ++j)
                {
                    dataTS[j + 1, i].Value = data.qhGiangVienMon[i, j].ToString();
                }
            }
            //#endregion

            //#region  hien thi lịch bận giảng viên
            //// xoa sach nhung j nhin thay
            DeleteDataView(dataBusy);
            //// nhap du lieu vao dataGridView
            // tieu de
            newCol = new DataGridViewColumn(); // add a column to the grid         
            newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
            newCol.CellTemplate = newCell;
            newCol.HeaderText = " Tea -Day ";
            newCol.Visible = true;
            dataBusy.Columns.Add(newCol);
            for (int i = 0; i < data.soNgay; ++i)
            {
                newCol = new DataGridViewColumn(); // add a column to the grid         
                newCell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column  
                newCol.CellTemplate = newCell;
                newCol.HeaderText = "Day" + (i + 1).ToString();
                newCol.Visible = true;
                dataBusy.Columns.Add(newCol);
            }
            // du lieu
            for (int i = 0; i < data.soGv; ++i)
            {
                dataBusy.Rows.Add();// them dong moi
                dataBusy[0, i].Value = data.giangVien[i];
                for (int j = 0; j < data.soNgay; ++j)
                {
                    dataBusy[j + 1, i].Value = data.lichBanGiangVien[i, j];
                }
            }
            //#endregion

        }

        public void DeleteDataView(DataGridView a)
        {
            int row = a.Rows.Count;
            if (row > 0)
                for (int i = row - 2; i >= 0; --i)
                    a.Rows.RemoveAt(i);
            int col = a.Columns.Count;
            if (col > 0)
                for (int i = col - 1; i >= 0; --i)
                    a.Columns.RemoveAt(i);
        }

        private void btnCreateData_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Path save file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                RWDataExcel a = new RWDataExcel(save.FileName);
                a.CreateFile(Convert.ToInt32(this.txtsolop.Text), Convert.ToInt32(this.txtsomon.Text), Convert.ToInt32(this.txtsogv.Text), Convert.ToInt32(this.txtsongay.Text));
                a.Show();
                MessageBox.Show("Tiếp Tục Nhập Dữ Liệu", "Thông Báo");
                a.Dispose();
                a.ReadSimpleFileAndCreteFullTitleData();
                a.Show();
                MessageBox.Show("Kết Thúc Nhập Dữ Liệu", "Thông Báo");
                a.Dispose();
            }
            else
            {

            }
        }

        private void btnRunAll_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Path save file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                KhoiTaoQuanThe();
                lichSang.TienHoaKetXuatDuLieu(Convert.ToInt32(this.textBox11.Text), save.FileName);
            }
        }

        public void KhoiTaoQuanThe()
        {
            int[] sang;
            int lopSang = 0;
            foreach (Object cb in checkedListBox1.CheckedItems)
                ++lopSang;
            sang = new int[lopSang];
            if (lopSang < 1)
            {
                MessageBox.Show("Chưa chọn lớp để lập lịch !");
            }
            else
            {
                bool[] chon = new bool[data.soLop];
                int k = -1;
                foreach (Object cb in checkedListBox1.CheckedItems)
                {
                    ++k;
                    for (int i = 0; i < data.soLop; ++i)
                    {
                        if (data.lop[i] == cb.ToString())
                        {
                            sang[k] = i;
                            chon[i] = true;
                        }
                    }
                }

                // lay cac tham so
                int size = Convert.ToInt32(textBox12.Text); // txt12
                int soNgayTrongTuan = Convert.ToInt32(txtsongay.Text);
                int soGioTrongBuoi = Convert.ToInt32(MaxGiongay.Text);
                int soGioTrongTiet = Convert.ToInt32(MaxGioMon.Text);
                double lai = Convert.ToDouble(textBox13.Text); // txt13
                double dotBien = Convert.ToDouble(textBox14.Text); // txt14
                lichSang = new ClassQuanThe(size, sang, data, soNgayTrongTuan, soGioTrongBuoi, soGioTrongTiet, lai, dotBien);
            }
        }

        private void btnDoiCuoi_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Path save file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                this.KhoiTaoQuanThe();
                try
                {
                    lichSang.TienHoaKetXuatDuLieuDoiCuoi(Convert.ToInt32(this.textBox11.Text), save.FileName);
                }
                catch (Exception et)
                {
                    MessageBox.Show(et.ToString(), "Lỗi chưa khởi tạo tham số !");
                }
            }
        }

        private void btnRunHoanHao_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Path save file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                this.KhoiTaoQuanThe();
                lichSang.TienHoaKetXuatDuLieuHoanHao(save.FileName);
            }
        }

        private void btnKetqua_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Path save file Excel";
            save.Filter = "Excel(2003)  (*.xls)|*.xls|Excel(2007) (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                lichSang.KetXuatKetQuaCuoiCung(save.FileName);
            }
        }
    }
}
