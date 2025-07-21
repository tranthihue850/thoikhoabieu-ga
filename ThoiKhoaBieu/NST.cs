using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ThoiKhoaBieu
{
    public class NST //đại diện cho một cá thể độc lập. 1 cá thể là một lời giải cho bài toán
    {
        //Data- Property
        #region Data
        public static Random ran = new Random();// bàn tay của tự nhiên;
        public static double maxDoThichNghi = 1000;
        public double doThichNghi = 0, xacSuatSong = 0, mucTieu, mucTieu2,mucTieu3=0;
        public string nguonGoc;
        // thông số kiểm soát cá thể
        public static int Count;// tổng số cá thể tồn tại trong toàn bộ quá trình phát triển
        public int ID;
        // Cau truc NST
        public RWDataExcel dataNST;
        public int[] cacLop;// danh sach cac lop can dc lap thoi khoa bieu
        public Gen[] gen;
        #endregion
        // Method
        #region Method

        #region Các hàm khởi tạo
        public NST()
        {
        }
        public NST(int[] lop, RWDataExcel data)// khởi tạo ngẫu nhiên một NST
        {
            ++NST.Count;
            this.ID = NST.Count;
            this.nguonGoc = "NST Gốc";
            // khoi tao 
            this.dataNST = data;
            this.cacLop = new int[lop.Length];
            for (int i = 0; i < cacLop.Length; ++i)
                cacLop[i] = lop[i];
            this.gen = new Gen[cacLop.Length];
            for (int i = 0; i < cacLop.Length; ++i)
                gen[i] = new Gen(cacLop[i], data);

        }
        // sao chép NST
        public NST(NST bo)
        {
            ++NST.Count;
            this.ID = NST.Count;
            this.nguonGoc = "Bố :" + bo.ID.ToString();
            //--------copy du lieu tu bo sang con
            dataNST = bo.dataNST;
            this.cacLop = new int[bo.cacLop.Length];
            this.gen = new Gen[bo.gen.Length];
            for (int i = 0; i < cacLop.Length; ++i)
                cacLop[i] = bo.cacLop[i];
            for (int i = 0; i < gen.Length; ++i)
                gen[i] = new Gen(bo.gen[i]);
        }
        #endregion

        #region Hàm thuộc tính di truyền
        public static NST operator +(NST bo, NST me)//phép lai
        {
            NST con = new NST(bo);
            con.nguonGoc = "Lai: " + bo.ID.ToString() + "+" + me.ID.ToString();
            //===============================================
            // khoi tao 2 dien lai           
            int cuoi = NST.ran.Next(bo.gen.Length);
            int dau = NST.ran.Next(cuoi);
            for (int i = dau; i <= cuoi; ++i)
                con.gen[i] = new Gen(me.gen[i]);
            return con;
        }
        //-----------------------------------
        public static NST operator !(NST bo)//phép đột biến
        {
            NST con = new NST(bo);
            con.nguonGoc = "Đột Biến: " + bo.ID.ToString();
            //===============================================
            // chon 1 gen de dot bien
            int db = NST.ran.Next(bo.gen.Length);
            con.gen[db] = new Gen(con.gen[db].maLop, bo.dataNST);
            return con;
        }
        #endregion

        #region Các hàm hỗ trợ
        public void SetDoThichNghi(double ts, int hocQua1Lan,int trungLichBan)
        {
            this.mucTieu = ts;
            this.mucTieu2 = hocQua1Lan;
            mucTieu3 = trungLichBan;
            this.doThichNghi = NST.maxDoThichNghi;
            this.doThichNghi -= ts * 25;
            this.doThichNghi -= hocQua1Lan * 5;
            doThichNghi -= trungLichBan * 10;
            if (doThichNghi < 0) doThichNghi = 1;
        }
        public void SetXacSuatSong(double ts)
        {
            this.xacSuatSong = ts;
        }
        public static bool operator &(NST a, NST b)
        {
            bool ok = true;
            if (a.cacLop.Length != b.cacLop.Length) ok = false;
            else
            {
                for (int i = 0; i < a.cacLop.Length; ++i)
                    if ((a.gen[i] & b.gen[i]) == false) { ok = false; break; }
            }
            return ok;

        }
        // kiểm tra trong 1 NST số lần các lớp học 2 tiết liên tục cùng 1 môn
        public int TrungLap()
        {
            int trungNhau = 0;
            for (int i = 0; i < cacLop.Length - 1; ++i)
                for (int j = i + 1; j < cacLop.Length; ++j)
                    trungNhau += gen[i] | gen[j];
            return trungNhau;

        }
        public int HocQua1Lan()
        {
            int trung = 0;
            for (int i = 0; i < cacLop.Length; ++i)
                trung += gen[i].HocQua1Lan();
            return trung;
        }
        public int TrungLichBan(Busy busy)
        {
            int trung = 0;
            for (int i = 0; i < cacLop.Length; ++i)
                trung += busy.TrungLichBan(gen[i]);
            return trung;
        }
        public String Show()
        {
            string s = "'";
            for (int i = 0; i < cacLop.Length; ++i)
            {
                s += "Lop:" + cacLop[i].ToString() + "\r\n";
                s += gen[i].Show();
                s += "====================================================\r\n";
            }

            return s;
        }
        public string ShortShow()
        {
            string s = "";
            if (gen[0] != null)
            {
                // string s = "";

                try
                {
                    // Kiểm tra xem gen[0], ngay[0], tiet[0] và tiet[1] có hợp lệ không
                    if (gen.Length > 0 && gen[0].ngay.Length > 0)
                    {
                        if (gen[0].ngay[0].tiet[0] != null)
                        {
                            if (gen[0].ngay[0].tiet.Length > 1) // Đảm bảo có ít nhất 2 phần tử trong tiet
                            {
                                s += gen[0].ngay[0].tiet[0].ShowAll() + "  " + gen[0].ngay[0].tiet[1].ShowAll() + ".......";
                            }
                            else
                            {
                                s += "Không đủ tiet để hiển thị.";
                            }
                        }
                        else
                        {
                            s += "Null Không đủ tiet để hiển thị.";
                        }    
                    }
                    else
                    {
                        s += "Không có dữ liệu ngày.";
                    }
                }
                catch (IndexOutOfRangeException)
                {
                    // Nếu có exception do chỉ số không hợp lệ
                    s += "Lỗi: chỉ số không hợp lệ.";
                }
                catch (NullReferenceException)
                {
                    // Nếu gen hoặc các phần tử không tồn tại
                    s += "Lỗi: dữ liệu không tồn tại.";
                }
                catch (Exception ex)
                {
                    // Bắt các loại exception khác nếu cần thiết
                    s += "Lỗi không xác định: " + ex.Message;
                }
            }
            // s += gen[0].ngay[0].tiet[0].ShowAll() + "  " + gen[0].ngay[0].tiet[1].ShowAll() + ".......";
            return s;
        }
        #endregion

        #endregion

    }
}
