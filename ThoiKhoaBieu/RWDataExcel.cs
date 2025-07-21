using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ThoiKhoaBieu
{
    public class RWDataExcel
    {
        #region Data-Properties
        private Excel.Application xApp;
        private Excel.Workbook xBook;
        private Excel.Worksheet xSheet;
        private Range range;
        private String path;
        // dữ liệu cho bài toán lập lịch
        public String[] lop, mon, giangVien;
        public int soLop, soMon, soGv, soNgay;
        public int[,] qhLopMon;
        public bool[,] qhLopGiangVien, qhGiangVienMon, TKBGiaovien;
        public String[,] lichBanLop, lichBanGiangVien;
        // dữ liệu phục vụ kết xuất dữ liệu
        private int conTro;// chỉ ra vị trí cuối cùng trong file Excel đang kết suất
        #endregion

        public RWDataExcel()
        { }
        public RWDataExcel(string st)
        {
            this.path = st;
            xApp = new Excel.Application();
            //Create a workbook object 
            xBook = (Excel.Workbook)xApp.Workbooks.Add(1);
            //Assign the active worksheet of the workbook 
            //object to a worksheet object 
            xSheet = (Excel.Worksheet)xBook.ActiveSheet;
            //Excel.XlFileFormat.xlXMLSpreadsheet 
            xBook.SaveAs(st, Excel.XlFileFormat.xlWorkbookNormal, "", "", false, false, 0, "", 0, "", "", "");
            xApp.Visible = false;
        }

        #region  hàm hỗ trợ
        public void Show()
        {
            xApp.Visible = true;
        }
        public void Dispose()
        {
            xApp.Quit();
        }
        public String ChangeIntToStringCel(int a)
        {
            String re = "";
            int tg;
            while (a > 0)
            {
                tg = a % 26;
                a = a / 26;
                if (tg == 0)
                {
                    re = "Z" + re;
                    --a;
                }
                else re = Convert.ToChar(tg + 64).ToString() + re;
            }

            return re;
        }
        #endregion

        #region Các hàm hỗ trợ nhập đọc dữ liệu cho bài toán lập lich

        public void CreateFile(int soLopHoc, int soMonHoc, int soGV, int ngayHoc)
        {
            this.soLop = soLopHoc; this.soMon = soMonHoc; this.soGv = soGV; this.soNgay = ngayHoc;
            String infor = soLop.ToString() + "," + soMon.ToString() + "," + soGv.ToString() + "," + soNgay.ToString();

            #region Khởi tạo tiêu đề, ghi thông số vào file excel
            range = xSheet.get_Range("E1", "H1");
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 22;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Dữ Liệu Lập Lịch");
            range.set_Item(1, 100, infor);
            #endregion

            #region phần nhập dữ liệu cho từng lớp
            range = xSheet.get_Range("A3", "A4");
            range.Interior.ColorIndex = 27;//22 40
            // range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 12;
            range.set_Item(1, 1, "Mã Lớp");
            for (int i = 1; i <= soLop; ++i)
                range.set_Item(1, i + 1, i.ToString());
            range.set_Item(2, 1, "Tên Lớp");
            range.Columns.AutoFit();
            //tô màu vùng nhập dữ liệu
            for(int i=2;i<=soLop+1;++i)
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "3", this.ChangeIntToStringCel(i) + "4");
                    range.Interior.ColorIndex = 34;//22 40
                    range.Font.Color = 12;
                }else

               {
                range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "3", this.ChangeIntToStringCel(i) + "4");
                range.Interior.ColorIndex = 37;
                range.Font.Color = 12;
               }

            #endregion

            #region phần nhập dữ liệu cho từng môn học
            range = xSheet.get_Range("A6", "A7");
            range.Interior.ColorIndex = 45;//22 40
            // range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 12;
            range.set_Item(1, 1, "Mã Môn");
            for (int i = 1; i <= soMon; ++i)
                range.set_Item(1, i + 1, i.ToString());
            range.set_Item(2, 1, "Tên Môn");
            range.Columns.AutoFit();
            //tô màu vùng nhập dữ liệu
            for (int i = 2; i <= soMon + 1; ++i)
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "6", this.ChangeIntToStringCel(i) + "7");
                    range.Interior.ColorIndex = 34;//22 40
                    range.Font.Color = 12;
                }
                else
                {
                    range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "6", this.ChangeIntToStringCel(i) + "7");
                    range.Interior.ColorIndex =37;
                    range.Font.Color = 12;
                }
              

            #endregion

            #region phần nhập dữ liệu cho từng giảng viên
            range = xSheet.get_Range("A9", "A10");
            range.Interior.ColorIndex = 13;//22 40
            // range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Mã GV");
            for (int i = 1; i <= soGV; ++i)
                range.set_Item(1, i + 1, i.ToString());
            range.set_Item(2, 1, "Tên GV");
            range.Columns.AutoFit();
            //tô màu vùng nhập dữ liệu
            for (int i = 2; i <= soGv + 1; ++i)
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "9", this.ChangeIntToStringCel(i) + "10");
                    range.Interior.ColorIndex = 34;
                    range.Font.Color = 12;
                }
                else
                {
                    range = xSheet.get_Range(this.ChangeIntToStringCel(i) + "9", this.ChangeIntToStringCel(i) + "10");
                    range.Interior.ColorIndex = 37;
                    range.Font.Color = 12;
                }

            #endregion

        }
        public void ReadSimpleFileAndCreteFullTitleData()
        {
            #region Khởi tạo phần đọc dữ  liệu cũ
            xApp = new Excel.Application();
            string workbookPath = path;
            xBook = xApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, true, 0, true, true, true);
            xSheet = (Excel.Worksheet)xBook.ActiveSheet;
            // bắt đầu đọc tên lớp tên môn và tên giảng viên
            Excel.Range range;
            Object valueOfRang;
            String tg;
            int key = 2;// key=2 tuong ung voi B
            #endregion

            #region  đọc tên lớp bắt đầu từ B4
            this.lop = new String[soLop];
            for (int i = 0; i < this.soLop; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "4";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    lop[i] = valueOfRang.ToString();
                else lop[i] = "";
                ++key;
            }
            #endregion

            #region  đọc tên môn bắt đầu từ B7
            key = 2;
            this.mon = new String[soMon];
            for (int i = 0; i < this.soMon; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "7";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    mon[i] = valueOfRang.ToString();
                else mon[i] = "";
                ++key;
            }
            #endregion

            #region  đọc tên giảng viên bắt đầu từ B10
            key = 2;
            this.giangVien = new String[soGv];
            for (int i = 0; i < this.soGv; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "10";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    this.giangVien[i] = valueOfRang.ToString();
                else this.giangVien[i] = "";
                ++key;
            }
            #endregion

            #region Khởi tạo phần nhập dữ liệu mới
            int viTri = 13;// con tro trong excel chi ra vi tri hien hanh se tao du lieu

            #region phần nhập dữ liệu cho quan hệ lớp-môn học
            #region Khởi tạo tiêu đê cho phần lớp môn học
            String cel1 = "A" + viTri.ToString();
            String cel2 = "C" + viTri.ToString();
            range = xSheet.get_Range(cel1, cel2);
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Quan hệ Lớp-Môn ");
            viTri += 1;
            #endregion

            range = xSheet.get_Range("A" + viTri.ToString(), ChangeIntToStringCel(soMon + 1) + (viTri + soLop).ToString());
            range.Interior.ColorIndex = 40;//22 40          
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Lớp/Môn");
            for (int i = 1; i <= soMon; ++i)
                range.set_Item(1, i + 1, mon[i - 1]);
            for (int i = 1; i <= soLop; ++i)
                range.set_Item(i + 1, 1, lop[i - 1]);
            range.Columns.AutoFit();
            // tô màu cho vùng nhập dữ liệu
            for (int i = viTri + 1; i <= viTri + soLop; ++i)
                if (i%2==0)
            {
                range = xSheet.get_Range("B" + (i).ToString(), ChangeIntToStringCel(soMon + 1) + (i).ToString());

                range.Interior.ColorIndex = 34;//22 40          
                range.Font.Italic = true;
                range.Font.Size = 14;
                range.Font.Name = "Times New Roman";
                range.Font.Color = 19;
            }
            else
            { range = xSheet.get_Range("B" + (i).ToString(), ChangeIntToStringCel(soMon + 1) + (i).ToString());

                range.Interior.ColorIndex = 37;        
                range.Font.Italic = true;
                range.Font.Size = 14;
                range.Font.Name = "Times New Roman";
                range.Font.Color = 19;
            }
            #endregion

            viTri += soLop + 3;
            // vi tri tang len soLop+4 sau phần khởi tạo quan hệ lớp-môn học ở trên

            #region phần nhập dữ liệu cho quan hệ lớp-giảng viên

            #region Khởi tạo tiêu đê cho phần lớp -giảng viên
            cel1 = "A" + viTri.ToString();
            cel2 = "C" + viTri.ToString();
            range = xSheet.get_Range(cel1, cel2);
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Quan hệ Lớp-Giảng Viên ");
            viTri += 1;
            #endregion

            range = xSheet.get_Range("A" + viTri.ToString(), ChangeIntToStringCel(soGv + 1) + (viTri + soLop).ToString());
            range.Interior.ColorIndex = 40;//22 40          
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Lớp/Gv");
            for (int i = 1; i <= soGv; ++i)
                range.set_Item(1, i + 1, this.giangVien[i - 1]);
            for (int i = 1; i <= soLop; ++i)
                range.set_Item(i + 1, 1, lop[i - 1]);
            range.Columns.AutoFit();
            // tô màu cho vùng nhập dữ liệu
            for (int i = viTri + 1; i <= viTri + soLop;++i )
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soGv + 1) + i.ToString());
                    range.Interior.ColorIndex = 34;//22 40          
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }
                else
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soGv + 1) + i.ToString());
                    range.Interior.ColorIndex = 37;
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }
            #endregion

            viTri += soLop + 3;
            // vi tri tang len soLop+4 sau phần khởi tạo quan hệ lớp-giảng viên ở trên

            #region phần nhập dữ liệu cho quan hệ giảng viên-môn học
            #region Khởi tạo tiêu đê cho phần lớp môn học
            cel1 = "A" + viTri.ToString();
            cel2 = "C" + viTri.ToString();
            range = xSheet.get_Range(cel1, cel2);
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Quan hệ Gv-Môn ");
            viTri += 1;
            #endregion

            range = xSheet.get_Range("A" + viTri.ToString(), ChangeIntToStringCel(soMon + 1) + (viTri + soGv).ToString());
            range.Interior.ColorIndex = 40;//22 40          
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Gv/Môn");
            for (int i = 1; i <= soMon; ++i)
                range.set_Item(1, i + 1, mon[i - 1]);
            for (int i = 1; i <= soGv; ++i)
                range.set_Item(i + 1, 1, giangVien[i - 1]);
            range.Columns.AutoFit();
            // tô màu cho vùng nhập dữ liệu
            for (int i = viTri + 1; i <= viTri + soGv; ++i)
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soMon + 1) + i.ToString());
                    range.Interior.ColorIndex = 34;//22 40          
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }
                else
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soMon + 1) + i.ToString());
                    range.Interior.ColorIndex = 37;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }
            #endregion

            viTri += soGv + 3;
            // vi tri tang len soLop+4 sau phần khởi tạo quan hệ giảng viên-môn học ở trên
          #region phần nhập dữ liệu cho lịch bận của lớp
            #region Khởi tạo tiêu đê cho phần lịch bận của lớp
            /*   cel1 = "A" + viTri.ToString();
            cel2 = "C" + viTri.ToString();
            range = xSheet.get_Range(cel1, cel2);
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Lịch bận của lớp ");
            viTri += 1;
            #endregion

            range = xSheet.get_Range("A" + viTri.ToString(), ChangeIntToStringCel(soNgay + 1) + (viTri + soLop).ToString());
            range.Interior.ColorIndex = 40;//22 40          
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Lớp/Ngày");
            for (int i = 2; i <= soNgay + 1; ++i)
                range.set_Item(1, i, "Thứ " + i.ToString());
            for (int i = 1; i <= soLop; ++i)
                range.set_Item(i + 1, 1, lop[i - 1]);
            range.Columns.AutoFit();
            // tô màu cho vùng nhập dữ liệu
            for (int i = viTri + 1; i <= viTri + soLop; ++i)
                if (i % 2 == 0)
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soNgay + 1) + i.ToString());

                    range.Interior.ColorIndex = 34;//22 40          
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }
                else
                {
                    range = xSheet.get_Range("B" + i.ToString(), ChangeIntToStringCel(soNgay + 1) + i.ToString());

                    range.Interior.ColorIndex = 27;
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;

                }
        */
          #endregion

            //     viTri += soLop + 3;
            // vi tri tang len soLop+4 sau phần khởi tạo lịch bận của lớp ở trên
            #region phần nhập dữ liệu cho lịch bận của giảng viên
            #region Khởi tạo tiêu đê cho phần lịch bận của giảng viên
            cel1 = "A" + viTri.ToString();
            cel2 = "C" + viTri.ToString();
            range = xSheet.get_Range(cel1, cel2);
            range.Interior.ColorIndex = 36;//22 40
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Lịch bận của giảng viên ");
            viTri += 1;
            #endregion

            range = xSheet.get_Range("A" + viTri.ToString(), ChangeIntToStringCel(soNgay + 1) + (viTri + soGv).ToString());
            range.Interior.ColorIndex = 40;//22 40          
            range.Font.Italic = true;
            range.Font.Size = 14;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 19;
            range.set_Item(1, 1, "Gv/Ngày");
            for (int i = 2; i <= soNgay + 1; ++i)
                range.set_Item(1, i, "Thứ " + i.ToString());
            for (int i = 1; i <= soGv; ++i)
                range.set_Item(i + 1, 1, giangVien[i - 1]);
            range.Columns.AutoFit();
            // tô màu cho vùng nhập dữ liệu
            for (int i = viTri + 1; i <= viTri + soGv ; ++i)
                if (i % 2 == 0)
                {

                    range = xSheet.get_Range("B" + (i).ToString(), ChangeIntToStringCel(soNgay + 1) + (i).ToString());

                    range.Interior.ColorIndex = 34;//22 40          
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }else
                {

                    range = xSheet.get_Range("B" + (i).ToString(), ChangeIntToStringCel(soNgay + 1) + (i).ToString());

                    range.Interior.ColorIndex =37;
                    range.Font.Italic = true;
                    range.Font.Size = 14;
                    range.Font.Name = "Times New Roman";
                    range.Font.Color = 19;
                }

            #endregion

            #endregion
        #endregion
            #region  doc TKB cua tung giao vien

            #endregion
        }
        public void ReadAllDataInExcelFile(string pa)
        {
            this.path = pa;
            // String pt = "";
            #region Khởi tạo phần đọc dữ  liệu
            xApp = new Excel.Application();
            string workbookPath = path;
            xBook = xApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, true, 0, true, true, true);
            xSheet = (Excel.Worksheet)xBook.ActiveSheet;
            // bắt đầu đọc tên lớp tên môn và tên giảng viên
            Excel.Range range;
            Object valueOfRang;
            String tg;
            int key = 2;// key=2 tuong ung voi B
            #endregion

            #region Đọc thông số ban đầu
            String st = "CZ1";// mã excel của ô 1,100 duoc chọn để ghi dữ liệu
            range = xSheet.get_Range(st, st);
            valueOfRang = range.Value2;
            String[] thongSo = valueOfRang.ToString().Split(',');
            this.soLop = Convert.ToInt32(thongSo[0]);
            this.soMon = Convert.ToInt32(thongSo[1]);
            this.soGv = Convert.ToInt32(thongSo[2]);
            this.soNgay = Convert.ToInt32(thongSo[3]);
            #endregion

            #region  đọc tên lớp bắt đầu từ B4
            this.lop = new String[soLop];
            for (int i = 0; i < this.soLop; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "4";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    lop[i] = valueOfRang.ToString();
                else lop[i] = "";
                ++key;
            }

            #endregion

            #region  đọc tên môn bắt đầu từ B7
            key = 2;
            this.mon = new String[soMon];
            for (int i = 0; i < this.soMon; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "7";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    mon[i] = valueOfRang.ToString();
                else mon[i] = "";
                ++key;
            }
            #endregion

            #region  đọc tên giảng viên bắt đầu từ B10
            key = 2;
            this.giangVien = new String[soGv];
            for (int i = 0; i < this.soGv; ++i)
            {
                tg = this.ChangeIntToStringCel(key) + "10";
                range = xSheet.get_Range(tg, tg);
                valueOfRang = range.Value2;
                if (valueOfRang != null)
                    this.giangVien[i] = valueOfRang.ToString();
                else this.giangVien[i] = "";
                ++key;

            }
            #endregion
            int viTri = 15;// khởi tạo vị trí bắt đầu đọc dữ liệu 

            #region Đọc quan hệ lớp - môn
            this.qhLopMon = new int[soLop, soMon];
            for (int i = 0; i < soLop; ++i)
            {
                key = 2;
                for (int j = 0; j < soMon; ++j)
                {
                    st = ChangeIntToStringCel(key) + viTri.ToString();
                    range = xSheet.get_Range(st, st);
                    valueOfRang = range.Value2;
                    ++key;
                    if (valueOfRang != null)
                        qhLopMon[i, j] = Convert.ToInt32(valueOfRang.ToString());
                    else
                        qhLopMon[i, j] = 0;
                }
                ++viTri;
            }

            #endregion

            viTri += 4;

            #region Đọc quan hệ lớp - giảng viên
            this.qhLopGiangVien = new bool[soLop, soGv];
            for (int i = 0; i < soLop; ++i)
            {
                key = 2;
                for (int j = 0; j < soGv; ++j)
                {
                    st = ChangeIntToStringCel(key) + viTri.ToString();
                    range = xSheet.get_Range(st, st);
                    valueOfRang = range.Value2;
                    ++key;
                    if (valueOfRang != null)
                        qhLopGiangVien[i, j] = true;
                    else
                        qhLopGiangVien[i, j] = false;
                }
                ++viTri;
            }

            #endregion

            viTri += 4;

            #region Đọc quan hệ giảng viên - môn
            this.qhGiangVienMon = new bool[soGv, soMon];
            for (int i = 0; i < soGv; ++i)
            {
                key = 2;
                for (int j = 0; j < soMon; ++j)
                {
                    st = ChangeIntToStringCel(key) + viTri.ToString();
                    range = xSheet.get_Range(st, st);
                    valueOfRang = range.Value2;
                    ++key;
                    if (valueOfRang != null)
                        this.qhGiangVienMon[i, j] = true;
                    else
                        this.qhGiangVienMon[i, j] = false;
                }
                ++viTri;
            }

            #endregion

         //   viTri += 4;

            #region Đọc lịch bận của Lớp
          /*  this.lichBanLop = new String[soLop, soNgay];
            for (int i = 0; i < soLop; ++i)
            {
                key = 2;
                for (int j = 0; j < soNgay; ++j)
                {
                    st = ChangeIntToStringCel(key) + viTri.ToString();
                    range = xSheet.get_Range(st, st);
                    valueOfRang = range.Value2;
                    ++key;
                    if (valueOfRang != null)
                        lichBanLop[i, j] = valueOfRang.ToString();
                    else
                        lichBanLop[i, j] = "";
                }
                ++viTri;
            }
            */
            #endregion

           
            viTri += 4;

            #region Đọc lịch bận của giang vien
            lichBanGiangVien = new String[soGv, soNgay];
            for (int i = 0; i < soGv; ++i)
            {
                key = 2;
                for (int j = 0; j < soNgay; ++j)
                {
                    st = ChangeIntToStringCel(key) + viTri.ToString();
                    range = xSheet.get_Range(st, st);
                    valueOfRang = range.Value2;
                    ++key;
                    if (valueOfRang != null)
                        lichBanGiangVien[i, j] = valueOfRang.ToString();
                    else
                        lichBanGiangVien[i, j] = "";
                }
                ++viTri;
            }

            #endregion

        }
        #endregion

        #region Kết xuất dữ liệu từ thuật giải di truyền ra ExcelFile

        public void CreteTitleForOutPutExcelFile()
        {
            this.conTro = 1;
            #region Khởi tạo tiêu đề, ghi thông số vào file excel
            range = xSheet.get_Range("E1", "F1");
            range.Interior.ColorIndex = 6;//22 40
            range.Font.Bold = true;
            range.Font.Underline = true;
            range.Font.Size = 22;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 60;
            range.Columns.AutoFit();
            range.set_Item(1, 1, "Quá Trình Tiến Hóa");
            #endregion
        }
        public void CreateTitleForOneStepInExcelOutPut(int doi)
        {
            conTro += 2;
            #region Khởi tạo tiêu đề cho tung doi
            range = xSheet.get_Range("A" + conTro.ToString(), "A" + conTro.ToString());
            range.Interior.ColorIndex = 4;//22 40
            range.Font.Bold = true;
            range.Font.Size = 19;
            range.Font.Underline = true;
            range.Font.Italic = true;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 67;
            range.set_Item(1, 1, "Đời " + doi.ToString() + ":");
            range.Columns.AutoFit();
            #endregion
            conTro += 1;

            #region Khởi tạo tiêu đề trong bang cho tung doi
            range = xSheet.get_Range("B" + conTro.ToString(), "J" + conTro.ToString());
            range.Interior.ColorIndex = 19;//22 40
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Size = 15;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 23;
            range.set_Item(1, 1, "STT"); range.Columns.AutoFit();
            range.set_Item(1, 2, " Cấu trúc NST"); range.Columns.AutoFit();
            range.set_Item(1, 3, "ID NST"); range.Columns.AutoFit();
            range.set_Item(1, 4, "Độ thích nghi"); range.Columns.AutoFit();
            range.set_Item(1, 5, "Hàm mục tiêu 1"); range.Columns.AutoFit();
            range.set_Item(1, 6, "Hàm mục tiêu 2"); range.Columns.AutoFit();
            range.set_Item(1, 7, "Hàm mục tiêu 3"); range.Columns.AutoFit();
            range.set_Item(1, 8, " XS sống "); range.Columns.AutoFit();
            range.set_Item(1, 9, " Đặc tính "); range.Columns.AutoFit();

            #endregion

            conTro += 1;
        }
        public void OutPutDataForOneStep(int stt, int ID, string name, double doThichNghi, string mucTieu1, string mucTieu2,string mucTieu3, double xsSong, string nguonGoc)
        {
            range = xSheet.get_Range("B" + conTro.ToString(), "J" + conTro.ToString());
            range.Interior.ColorIndex = 34;//22 40
            range.Font.Italic = true;
            range.Font.Size = 12;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 1;
            range.set_Item(1, 1, stt.ToString());
            range.set_Item(1, 2, name);
            range.set_Item(1, 3, ID.ToString());
            range.set_Item(1, 4, Convert.ToInt64(doThichNghi).ToString());
            range.set_Item(1, 5, mucTieu1);
            range.set_Item(1, 6, mucTieu2);
            range.set_Item(1, 7, mucTieu3);
            range.set_Item(1, 8, Convert.ToInt32(xsSong * 100).ToString() + "%");
            range.set_Item(1, 9, nguonGoc);

            ++conTro;
        }
        public void OutPutTitleEndOfData(NST a, RWDataExcel data)
        {
            this.conTro = 1;
            #region Khởi tạo tiêu đề, ghi thông số vào file excel
            range = xSheet.get_Range("E1", "M1");
            range.Interior.ColorIndex = 8;//22 40
            range.Font.Bold = true;
            range.Font.Underline = true;
            range.Font.Size = 20;
            range.Font.Name = "Times New Roman";
            range.Font.Color = 60;
            range.Columns.AutoFit();
            range.set_Item(1, 1, "THỜI KHÓA BIỂU TOÀN KHÓA - ID : " + a.ID.ToString());
            range = xSheet.get_Range("A3", "E3");
            range.Interior.ColorIndex = 4;//22 40
            range.Font.Bold = true;
            range.Font.Underline = true;
            range.Font.Italic = true;
            range.Font.Size = 17;
            range.Font.Name = "Times New Roman";
            range.set_Item(1, 1, "THỜI KHÓA BIỂU CÁC LỚP");
            #endregion
            conTro = 4;

            for (int i = 0; i < a.gen.Length; ++i)
            {
                range = xSheet.get_Range("B" + conTro.ToString(), "B" + conTro.ToString());
                range.Interior.ColorIndex = 6;//22 40
                range.Font.Bold = true;
                range.Font.Size = 12;
                range.Font.Name = "Times New Roman";
                range.set_Item(1, 1, "Lớp: " + data.lop[a.gen[i].maLop]);
                ++conTro;
                range = xSheet.get_Range("C" + conTro.ToString(), ChangeIntToStringCel(2 + Gen.soNgay) + conTro.ToString());
                for (int j = 0; j < Gen.soNgay; ++j)
                {
                    range.Interior.ColorIndex = 8;//22 40
                    range.Font.Bold = true;
                    range.Font.Underline = true;

                    range.Font.Size = 12;
                    range.Font.Name = "Times New Roman";
                    range.set_Item(1, j + 1, "Thứ " + (j + 2).ToString());
                }
                ++conTro;
                range = xSheet.get_Range("C" + conTro.ToString(), "C" + conTro.ToString());
                for (int j = 0; j < Gen.soNgay; ++j)
                {
                    for (int k = 0; k < a.gen[i].ngay[j].tiet.Length; ++k)
                        if (a.gen[i].ngay[j].tiet[k] != null)
                        {
                            for (int tc = 0; tc < data.giangVien.Length; ++tc)
                                if (a.gen[i].ngay[j].tiet[k].maGiangVien == tc)
                                    range.set_Item(k + 1, j + 1, data.mon[a.gen[i].ngay[j].tiet[k].maMon] + "(" + a.gen[i].ngay[j].tiet[k].soGio.ToString() + " - " + data.giangVien[tc]+")");
                        }
                    //range.set_Item(k + 1, j + 1, data.mon[a.gen[i].ngay[j].tiet[k].maMon] + "(" + a.gen[i].ngay[j].tiet[k].soGio.ToString() + ")");
                }
                conTro += 4;
            }
            //==========================================
            range = xSheet.get_Range("A" + conTro.ToString(), "F" + conTro.ToString());
            range.Interior.ColorIndex = 4;//22 40
            range.Font.Bold = true;
            range.Font.Underline = true;
            range.Font.Italic = true;
            range.Font.Size = 17;
            range.Font.Name = "Times New Roman";
            range.set_Item(1, 1, "THỜI KHÓA BIỂU CÁC GIẢNG VIÊN");
            ++conTro;
            //-----
            for (int i = 0; i < data.giangVien.Length; ++i)
            {
                range = xSheet.get_Range("B" + conTro.ToString(), "D" + conTro.ToString());
                range.Interior.ColorIndex = 6;//22 40
                range.Font.Bold = true;
                range.Font.Size = 12;
                range.Font.Name = "Times New Roman";
                range.set_Item(1, 1, "Giảng viên: " + data.giangVien[i]);
                ++conTro;
                range = xSheet.get_Range("C" + conTro.ToString(), ChangeIntToStringCel(2 + Gen.soNgay) + conTro.ToString());
                for (int j = 0; j < Gen.soNgay; ++j)
                {
                    range.Interior.ColorIndex = 8;//22 40
                    range.Font.Bold = true;
                    range.Font.Underline = true;

                    range.Font.Size = 12;
                    range.Font.Name = "Times New Roman";
                    range.set_Item(1, j + 1, "Thứ " + (j + 2).ToString());
                }
                ++conTro;
                range = xSheet.get_Range("C" + conTro.ToString(), "C" + conTro.ToString());
                for (int j = 0; j < Gen.soNgay; ++j)
                {
                    int tg = 0;
                    for (int k = 0; k < a.gen.Length; ++k)
                    {
                        int batDau = 1;
                        for (int p = 0; p < a.gen[k].ngay[j].tiet.Length; ++p)
                        {
                            if (a.gen[k].ngay[j].tiet[p] != null)
                            {
                                if (a.gen[k].ngay[j].tiet[p].maGiangVien == i)
                                {
                                    ++tg;
                                    range.set_Item(tg, j + 1, data.lop[k] + "-" + data.mon[a.gen[k].ngay[j].tiet[p].maMon] + "(" + batDau.ToString() + "-" + a.gen[k].ngay[j].tiet[p].soGio + ")");
                                }
                                batDau += a.gen[k].ngay[j].tiet[p].soGio;
                            }
                            else break;
                        }
                    }
                }
                conTro += 4;
            }
          

        }
        #endregion
        public void Quit()
        {
            xApp.Quit();
        }
    }
}

