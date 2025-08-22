using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.AccessControl;
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
        public void Properti_Draw()
        {
            int w = 1500; // Convert.ToInt32(textBox2.Text);
            int h = 1500; // Convert.ToInt32(textBox2.Text);
            double gv = 0.25; // //Convert.ToDouble(txtsogv.Text);
            double l = 0.5; // Convert.ToDouble(textBox7.Text);
            double m = 0.75; // Convert.ToDouble(textBox8.Text);
            DrawAll draw = new DrawAll(w, h, data, gv, l, m);
            draw.CreateTitle("Mô tả dữ liệu đầu vào", 32, Color.Red);
            draw.DrawAllArrow();
            draw.DrawAllPoint();
            int lgv = 300; // Convert.ToInt32(textBox6.Text);
            int ll = -15;// Convert.ToInt32(textBox4.Text);
            int lm = -100; // Convert.ToInt32(textBox5.Text);
            draw.DrawAllLabel(lgv, Color.Red, ll, Color.Blue, lm, Color.Violet);
            this.pictureBox1.Image = draw.Return();
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
                Properti_Draw();
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

    class DrawAll
    {
        #region Khai báo biến
        // bien chua du lieu can ve
        RWDataExcel data;
        // mang cac Point , moi point chua mot doi tuong can the hien
        Point[] giangVien, lop, mon;
        //cac du lieu dinh dang thuoc tinh cho mot bitmap
        public int width, height, radius;
        //cac doi tuong dung de ve        
        public Bitmap bm;
        public Graphics g;
        //---
        //   public static int length;
        //  public static Double[,] maTrix;
        public Random ran = new Random();
        //  public static int[] hoanhDo, tungDo;
        //   public static double chiPhi;
        public static string chuTrinh = "";
        #endregion

        #region Hàm hỗ trợ vẽ đồ thị

        #region Cac ham khoi tao du lieu va return toan cuc
        public DrawAll(int w, int h, RWDataExcel d, double gv, double l, double m)
        {
            this.width = w;
            this.height = h;
            this.data = d;
            this.radius = (w + h) / 2 * 2 / 100;//lay 1% kich co cua anh de tao diem
            bm = new Bitmap(width, height);
            g = Graphics.FromImage(bm);
            // khoi tao cac point cho cac  du lieu dau vao lay tu data the hien cho bai toan lap lich
            this.giangVien = new Point[data.soGv];
            this.lop = new Point[data.soLop];
            this.mon = new Point[data.soMon];
            // khoi tao vi tri cho cac point
            int x;
            x = Convert.ToInt32(width * gv);
            int tg = Convert.ToInt32((this.height * 0.8) / giangVien.Length);
            int y = Convert.ToInt32(this.height * 0.14);
            for (int i = 0; i < giangVien.Length; ++i)
            {
                giangVien[i] = new Point(x, y);
                y += tg;
            }
            // khoi tao  cho mon
            x = Convert.ToInt32(width * m);
            tg = Convert.ToInt32((this.height * 0.8) / mon.Length);
            y = Convert.ToInt32(this.height * 0.14);
            for (int i = 0; i < mon.Length; ++i)
            {
                mon[i] = new Point(x, y);
                y += tg;
            }
            // khoi tao cho lop
            x = Convert.ToInt32(width * l);
            tg = Convert.ToInt32((this.height * 0.8) / lop.Length);
            y = Convert.ToInt32(this.height * 0.17);
            for (int i = 0; i < lop.Length; ++i)
            {
                lop[i] = new Point(x, y);
                y += tg;
            }
            //MAU NEN

            g.FillRectangle(new Pen(Color.White).Brush, new Rectangle(0, 0, bm.Width, bm.Height));
            g.DrawRectangle(new Pen(Color.Red), new Rectangle(5, 5, bm.Width - 10, bm.Height - 10));


        }
        public Bitmap Return()
        {
            return this.bm;
        }
        #endregion

        public void CreateTitle(string s, int size, Color c)// ve tieu de cho mot buc anh
        {
            // ve tieu de 5% chieu cao va 35% chieu rong
            FontFamily fontFamily = new FontFamily("Arial");
            g.DrawString(s, new Font(fontFamily, size, FontStyle.Bold, GraphicsUnit.Pixel), new Pen(c).Brush, new Point(Convert.ToInt32(this.width * 0.35), Convert.ToInt32(this.height * 0.05)));
        }
        private void DrawPoint(Point p, string sc)//VE MOT DIEM
        {
            int r = this.radius;
            Rectangle rec1 = new Rectangle(p.X - r, p.Y - r, r + r, r + r);
            r -= 3;
            Rectangle rec2 = new Rectangle(p.X - r, p.Y - r, r + r, r + r);
            g.FillEllipse(new Pen(Color.Violet).Brush, rec1);
            g.FillEllipse(new Pen(Color.Cyan).Brush, rec2);
            // ve ten dinh
            FontFamily fontFamily = new FontFamily("Arial");
            g.DrawString(sc, new Font(fontFamily, r + r - 2, FontStyle.Regular, GraphicsUnit.Pixel), new Pen(Color.Green).Brush, new Point(p.X - r, p.Y - r));
        }
        private void DrawLabel(Point p, string sc, int change, Color c)//VE MOT label cho mot diem
        {
            int r = this.radius;
            Rectangle rec1 = new Rectangle(p.X - r, p.Y - r, r + r, r + r);
            r = Convert.ToInt32(r * 0.7);

            // ve ten dinh
            FontFamily fontFamily = new FontFamily("Arial");
            g.DrawString(sc, new Font(fontFamily, r + r - 2, FontStyle.Underline, GraphicsUnit.Pixel), new Pen(c).Brush, new Point(p.X - r - change, p.Y - r));
        }
        private void DrawLabelStyle(Point p, string sc, int change, Color c)//VE MOT label cho mot diem
        {
            int r = this.radius;
            Rectangle rec1 = new Rectangle(p.X - r, p.Y - r, r + r, r + r);
            r = Convert.ToInt32(r * 0.7);

            // ve ten dinh
            FontFamily fontFamily = new FontFamily("Arial");
            g.DrawString(sc, new Font(fontFamily, r + r - 2, FontStyle.Underline, GraphicsUnit.Pixel), new Pen(c).Brush, new Point(p.X - r - change, p.Y + 2 * r));
        }
        public void DrawAllPoint()
        {

            for (int i = 0; i < this.giangVien.Length; ++i)
                DrawPoint(giangVien[i], i.ToString());

            for (int i = 0; i < this.lop.Length; ++i)
                DrawPoint(lop[i], i.ToString());

            for (int i = 0; i < this.mon.Length; ++i)
                DrawPoint(mon[i], i.ToString());

        }
        public void DrawAllLabel(int change1, Color c1, int change2, Color c2, int change3, Color c3)
        {

            for (int i = 0; i < this.giangVien.Length; ++i)
                DrawLabel(giangVien[i], data.giangVien[i], change1, c1);

            for (int i = 0; i < this.lop.Length; ++i)
                DrawLabelStyle(lop[i], data.lop[i], change2, c2);

            for (int i = 0; i < this.mon.Length; ++i)
                DrawLabel(mon[i], data.mon[i], change3, c3);

        }
        public void DrawArrow(Point a, Point b, string s)
        {

            g.DrawLine(new Pen(Color.Green), a, b);
            // ve trong so canh
            /*    FontFamily fontFamily = new FontFamily("Arial");
                g.DrawString(s, new Font(fontFamily, radius, FontStyle.Regular, GraphicsUnit.Pixel), new Pen(Color.Green).Brush, new Point((b.X+(a.X + b.X) / 2)/2, (b.Y+(a.Y + b.Y) /2)/2));
    // ve mui ten tu a toi b
                Ten moi = new Ten(a.X, a.Y, (a.X + (a.X + b.X) / 2) / 2, (a.Y + (a.Y + b.Y) / 2) / 2, radius / 2);
                g.DrawLine(new Pen(Color.Red), moi.a1, moi.b1, (a.X + (a.X + b.X) / 2) / 2, (a.Y + (a.Y + b.Y) / 2) / 2);
                g.DrawLine(new Pen(Color.Red), moi.a2, moi.b2, (a.X + (a.X + b.X) / 2) / 2, (a.Y + (a.Y + b.Y) / 2) / 2);
           */
        }
        public void DrawAllArrow()
        {

            for (int i = 0; i < this.giangVien.Length; ++i)
                for (int j = 0; j < this.lop.Length; ++j)
                    if (data.qhLopGiangVien[j, i])
                        DrawArrow(giangVien[i], lop[j], "");

            for (int i = 0; i < this.mon.Length; ++i)
                for (int j = 0; j < this.lop.Length; ++j)
                    if (data.qhLopMon[j, i] > 0)
                        DrawArrow(lop[j], mon[i], data.qhLopMon[j, i].ToString());

            /*   for (int i = 0; i < this.giangVien.Length; ++i)
                   for (int j = 0; j < this.mon .Length; ++j)
                       if (data.qhGiangVienMon[i, j])
                           DrawArrow(giangVien[i],mon[j], "");
             * */


        }
        /*
        public void DrawArrow(int[] huong)
        {
            FontFamily fontFamily = new FontFamily("Arial");
            for (int i = 1; i < length; ++i)
            {
                g.DrawLine(new Pen(Color.Red), hoanhDo[huong[i - 1]], tungDo[huong[i - 1]], hoanhDo[huong[i]], tungDo[huong[i]]);
                Ten moi = new Ten(hoanhDo[huong[i - 1]], tungDo[huong[i - 1]], (hoanhDo[huong[i - 1]] + hoanhDo[huong[i]]) / 2, (tungDo[huong[i - 1]] + tungDo[huong[i]]) / 2, radius / 2);
                g.DrawLine(new Pen(Color.Red), moi.a1, moi.b1, (DrawAll.hoanhDo[huong[i - 1]] + DrawAll.hoanhDo[huong[i]]) / 2, (DrawAll.tungDo[huong[i - 1]] + DrawAll.tungDo[huong[i]]) / 2);
                g.DrawLine(new Pen(Color.Red), moi.a2, moi.b2, (hoanhDo[huong[i - 1]] + hoanhDo[huong[i]]) / 2, (tungDo[huong[i - 1]] + tungDo[huong[i]]) / 2);
                // ve trong so canh          
                g.DrawString(maTrix[huong[i - 1], huong[i]].ToString(), new Font(fontFamily, radius, FontStyle.Regular, GraphicsUnit.Pixel), new Pen(Color.Blue).Brush, new Point((hoanhDo[huong[i - 1]] + hoanhDo[huong[i]]) / 2, (tungDo[huong[i - 1]] + tungDo[huong[i]]) / 2));
            }
            int c = length - 1;
            g.DrawLine(new Pen(Color.Red), hoanhDo[huong[c]], tungDo[huong[c]], hoanhDo[huong[0]], tungDo[huong[0]]);
            Ten mi = new Ten(hoanhDo[huong[c]], tungDo[huong[c]], (hoanhDo[huong[c]] + hoanhDo[huong[0]]) / 2, (tungDo[huong[c]] + tungDo[huong[0]]) / 2, radius / 2);
            g.DrawLine(new Pen(Color.Red), mi.a1, mi.b1, (hoanhDo[huong[c]] + hoanhDo[huong[0]]) / 2, (tungDo[huong[c]] + tungDo[huong[0]]) / 2);
            g.DrawLine(new Pen(Color.Red), mi.a2, mi.b2, (hoanhDo[huong[c]] + hoanhDo[huong[0]]) / 2, (tungDo[huong[c]] + tungDo[huong[0]]) / 2);
            // ve trong so canh       
            g.DrawString(maTrix[huong[c], huong[0]].ToString(), new Font(fontFamily, radius, FontStyle.Regular, GraphicsUnit.Pixel), new Pen(Color.Blue).Brush, new Point((hoanhDo[huong[c]] + hoanhDo[huong[0]]) / 2, (tungDo[huong[c]] + tungDo[huong[0]]) / 2));
            // tinh tong chi phi
            Double count = maTrix[huong[length - 1], huong[0]];
            for (int i = 1; i < length; ++i)
                count += maTrix[huong[i - 1], huong[i]];
            // g.DrawString("Chi phi=" + count.ToString(), new Font(fontFamily, radius, FontStyle.Regular, GraphicsUnit.Pixel), new Pen(Color.Blue).Brush, new Point(width / 2, height - radius * 2));
            chiPhi = count;
            chuTrinh = (huong[0] + 1).ToString();
            for (int i = 1; i < length; ++i) chuTrinh += " ->" + (huong[i] + 1).ToString();
            chuTrinh += " ->" + (huong[0] + 1).ToString();

        }
       */
        #endregion

    }
}
