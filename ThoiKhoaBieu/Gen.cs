using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ThoiKhoaBieu
{
    public class Gen
    // một gen tương ứng  với một thời khóa biểu của một lớp 
    {
        //Data- Property
        #region Data
        public static int soNgay;// số ngày học trong tuần         
        public Ngay[] ngay;
        public int maLop;
        public int[] mon;// số h phải học của từng môn trong tuần (mã môn chính là index số h phải học của từng môn trong tuần)

        #endregion
        // Method
        #region Method
        #region Các hàm khởi tạo
        public Gen(RWDataExcel data,int maGv)
        {
            ngay = new Ngay[Gen.soNgay];
            for (int i=0;i<Gen .soNgay ;++i)
                ngay [i]=new Ngay (data .lichBanGiangVien [maGv ,i],maGv );
        }
        public Gen(int malop, RWDataExcel data)// khoi tao  ngau nhien mot gen
        {           
            this.maLop = malop;
            this.ngay = new Ngay[Gen.soNgay];
            for (int i = 0; i < Gen.soNgay; ++i)
                ngay[i] = new Ngay();
            this.mon = new int[data.soMon];
            for (int i = 0; i < mon.Length; ++i)
                mon[i] = data.qhLopMon[malop, i];
            // khoi tao ngong nhien mot thoi khoa bieu cho mot lop
            //mang cac tiet phai hoc trong tuan
            int tongSoTiet = -1;
            Tiet[] phaiHoc = new Tiet[Ngay.maxGio * Gen.soNgay*2];
            // MessageBox .Show ((Ngay.maxGio * Gen.soNgay).ToString ());
            // bat dau tim so tiet phai hoc trong mot tuan cua lop
            bool ok;
            for (int i = 0; i < mon.Length; ++i)
            {
                ok = false;
                if (data.qhLopMon[malop,i] >= 1)
                {
                    for (int j = 0; j < data.soGv; ++j)
                        if (data.qhLopGiangVien[malop, j] && data.qhGiangVienMon[j, i])
                        {
                            ok = true;
                            // khoi tao cac tiet phai hoc
                            int[] cacTiet = Split(mon[i], Tiet.maxGio);
                            for (int k = 0; k < cacTiet.Length; ++k)
                            {
                                ++tongSoTiet;
                                phaiHoc[tongSoTiet] = new Tiet(i, j, cacTiet[k]);
                            }
                            break;
                        }
                }
                else ok = true;
                if (ok == false)
                {
                    MessageBox.Show("Lớp " + data.lop[malop] + " không có giảng viên dạy môn " + data.mon[i], "Thông báo");
                    break;
                }
            }

            // khoi tao tung ngay tao thanh mot lich hoc ngau nhien
            // kiem tra xem co hoc het tat ca cac mon trong mot tuan khong
            int tongPhaiHocTrongTuan = 0;
            for (int i = 0; i < this.mon.Length; ++i)
                tongPhaiHocTrongTuan += mon[i];
            if (tongPhaiHocTrongTuan > (Gen.soNgay * Ngay.maxGio))
                MessageBox.Show("Số lượng các giời phải học là quá lớn để học trong 1 tuần", "Thông báo");

            // tao 1 thu tu cac mon ngau nhien phai hoc
            int[] arraySort = new int[tongSoTiet + 1];
            for (int i = 0; i < arraySort.Length; ++i) arraySort[i] = NST.ran.Next();
            Array.Sort(arraySort, phaiHoc);

            // dua cac mon theo thu tu vao thoi khoa bieu
            int daHoc = -1;
          
                daHoc = -1;
                for (int i = 0; i < Gen.soNgay; ++i)
                {
                    ngay[i] = new Ngay();
                    while ((daHoc< tongSoTiet) && (ngay[i].Add(phaiHoc[daHoc + 1])))  // điều kiện để dừng nếu ko thêm số lần dạy vào từng ngày trong tkb
                        ++daHoc;
                }
             int newStart = tongSoTiet+1;
             int newEnd = tongSoTiet;
             // khoi tao cac tiet phai hoc them lan nua
             for (int i = daHoc + 1; i <= tongSoTiet; ++i)
             {
                 int[] cacTiet = Split(phaiHoc [i].soGio, Tiet.maxGio-1);
                 int newMon = phaiHoc[i].maMon;
                 int newGv = phaiHoc[i].maGiangVien;
                 for (int k = 0; k < cacTiet.Length; ++k)
                 {
                     ++newEnd;
                     phaiHoc[newEnd] = new Tiet(newMon  ,newGv, cacTiet[k]);
                 }
             }
            //xep them lan nua
             daHoc = tongSoTiet;
             for (int i = 0; i < Gen.soNgay; ++i)
             {                
                 while ((daHoc < newEnd) && (ngay[i].Add(phaiHoc[daHoc + 1])))  // điều kiện để dừng nếu ko thêm số lần dạy vào từng ngày trong tkb
                     ++daHoc;
             }
             if (daHoc <newEnd)   MessageBox.Show("Không thể xếp lịch với quy định đã đưa vào, hãy điều chỉnh tham số đầu vào ", "Thông báo");
             int[] array = new int[Gen.soNgay ];
             for (int i = 0; i < array.Length; ++i) array[i] = NST.ran.Next();
             Array.Sort(array, ngay);
             // tim mon chao co vut len dau:
             if (ngay[0].tiet[0] != null)
             {
                 int firstMaMon = ngay[0].tiet[0].maMon;
                 int firstMaGv = ngay[0].tiet[0].maGiangVien;
                 for (int i = 0; i < Gen.soNgay; ++i)
                 {
                     ngay[i].SetRealLength();
                     for (int j = 0; j < ngay[i].realLength; ++j)
                     {
                         string result = "";

                         try
                         {
                            // Lấy phần mở rộng của maMon
                            //var checkMa = data.mon[ngay[i].tiet[j].maMon].Split('.');
                            if (ngay[i].tiet[j] != null)
                            {
                                string[] parts = data.mon[ngay[i].tiet[j].maMon].Split('.');
                                result = parts[parts.Length - 1]; // Lấy phần tử cuối cùng
                            }
                            else
                            {
                                result = "";
                            }    
                         }
                         catch (IndexOutOfRangeException)
                         {
                             // Nếu có exception, gán giá trị rỗng
                             result = "";
                         }
                         catch (Exception ex)
                         {
                             // Bắt các loại exception khác nếu cần thiết
                             result = "";
                         }

                         // var l6 = data.mon[ngay[i].tiet[j].maMon].Split('.')[data.mon[ngay[i].tiet[j].maMon].Split('.').Length - 1];
                         if (result == "L6")
                         {
                             ngay[0].ChangaeOneFirst(ngay[i].tiet[j].maMon, ngay[i].tiet[j].maGiangVien);
                             ngay[i].tiet[j] = new Tiet(firstMaMon, firstMaGv, 1);
                         }
                     }
                 }
             }
             Ngay tg;
             tg = ngay[0];
             ngay[0] = ngay[4];
             ngay[4] = tg;
             int end = 0;
             for (int i = 0; i < Ngay.maxGio; ++i)
                 if (ngay[4].tiet[i] != null) end = i;
                 else break;
             Tiet tgt;
             tgt = ngay[4].tiet[0];
             ngay[4].tiet[0] = ngay[4].tiet[end];
             ngay[4].tiet[end] = tgt;
            //-------------------------
             if (ngay[0].tiet[0] != null)
             {
                 int firstMaMon = ngay[0].tiet[0].maMon;
                 int firstMaGv = ngay[0].tiet[0].maGiangVien;
                 for (int i = 0; i < Gen.soNgay; ++i)
                 {
                     ngay[i].SetRealLength();
                     for (int j = 0; j < ngay[i].realLength; ++j)
                     {
                         // var abc = data.mon[ngay[i].tiet[j].maMon].Split('.')[data.mon[ngay[i].tiet[j].maMon].Split('.').Length - 1];
                         string result = "";

                         try
                         {
                            // Lấy phần mở rộng của maMon
                            // var par1 = data.mon[ngay[i].tiet[j].maMon]?.Split('.');
                            if (ngay[i].tiet[j] != null)
                            {
                                string[] parts = data.mon[ngay[i].tiet[j].maMon]?.Split('.');

                                // Kiểm tra xem mảng có phần tử hay không
                                if (parts.Length > 0)
                                {
                                    result = parts[parts.Length - 1]; // Lấy phần tử cuối cùng
                                }
                                else
                                {
                                    result = ""; // Nếu không có phần tử nào
                                }
                            }
                            else
                            {
                                result = "";
                            }    
                         }
                         catch (IndexOutOfRangeException)
                         {
                             // Nếu có exception do chỉ số không hợp lệ, gán giá trị rỗng
                             result = "";
                         }
                         catch (NullReferenceException)
                         {
                             // Nếu data.mon hoặc phần tử không tồn tại, gán giá trị rỗng
                             result = "";
                         }
                         catch (Exception ex)
                         {
                             // Bắt các loại exception khác nếu cần thiết
                             result = "";
                         }

                         //if (data.mon[ngay[i].tiet[j].maMon].Split('.')[data.mon[ngay[i].tiet[j].maMon].Split('.').Length - 1] == "F2")
                         if (result == "F2")
                         {
                             ngay[0].ChangaeOneFirst(ngay[i].tiet[j].maMon, ngay[i].tiet[j].maGiangVien);
                             ngay[i].tiet[j] = new Tiet(firstMaMon, firstMaGv, 1);
                         }
                     }
                 }
             }
          
        }
        public Gen(Gen bo)// ham sao chep gen
        {
            this.maLop = bo.maLop;
            this.ngay = new Ngay[Gen.soNgay];
            this.mon = new int[bo.mon.Length];
            for (int i = 0; i < mon.Length; ++i)
                mon[i] = bo.mon[i];
            for (int i = 0; i < Gen.soNgay; ++i)
            {
                ngay[i] = new Ngay();
                for (int j = 0; j < bo.ngay[i].tiet.Length; ++j)
                    if (bo.ngay[i].tiet[j] != null) ngay[i].Add(bo.ngay[i].tiet[j]);
            }
        }
        #endregion

        #region Hàm thuộc tính di truyền
        public static void SetSoNgayHocTrongTuan(int songay)
        {
            soNgay = songay;

        }
        public void DotBien()
        {

        }
        #endregion

        #region Các hàm hỗ trợ
        public string Show()
        {
            string st = "";
            for (int i = 0; i < Gen.soNgay; ++i)
            {
                st += "Thu " + (i + 2).ToString() + " : ";

                st += ngay[i].Show() + "\r\n";
            }
            return st;
        }
        public int[] Split(int a, int b)//chia mot Mon voi so gio phai hoc trong mot tuan thanh cac Tiet dam bao dieu kien so h phai hoc nho hon quy dinh
        {
           // MessageBox.Show(b.ToString());
            int tg = a / b;
            if (a % b > 0) ++tg;
            int[] re = new int[tg];
            for (int i = 0; i < re.Length; ++i) re[i] = 0;
            while (a > 0)
            {
                for (int i = 0; i < re.Length; ++i)
                {
                    ++re[i];
                    --a;
                    if (a <= 0) break;
                }

            }

            return re;
        }

        #region Kiểm tra sự trùng lặp giữa các lớp
        public static bool operator &(Gen a, Gen b)
        {
            bool ok = true;
            for (int i = 0; i < Gen.soNgay; ++i)
                if ((a.ngay[i] & b.ngay[i]) == false) { ok = false; break; }
            return ok;

        }
        public static int operator |(Gen a, Gen b)// tra lai so h trung lap giua cac loop
        {
            int trungNhau = 0;
            for (int i = 0; i < Gen.soNgay; ++i)
                trungNhau += a.ngay[i] | b.ngay[i];
            return trungNhau;

        }
        public static int operator ^(Gen a, Gen b)// tra lai su trung lap lich ban
        {
            int trungNhau = 0;
            for (int i = 0; i < Gen.soNgay; ++i)
                trungNhau += a.ngay[i] ^ b.ngay[i];
            return trungNhau;

        }
        // kiểm tra xem có ngày nào học 1 môn 2 lần ko
        public int HocQua1Lan()
        {
            int trung = 0;
            for (int i = 0; i < Gen.soNgay; ++i)
                trung += this.ngay[i].HocQua1Lan();
            return trung;
        }
        #endregion

        #endregion

        #endregion

    }
}
