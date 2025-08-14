using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ThoiKhoaBieu
{
    public class ClassTietHoc
    {
        //Data- Property
        #region Data
        // mỗi một tiết học gồm 2 phần 1: Mã môn học cho tiết đó 2: số giờ phải học
        public int maMon = 0, soGio = 0, maGiangVien = 0;
        public static int maxGio;
        #endregion
        // Method
        #region Method
        #region Các hàm khởi tạo
        public ClassTietHoc(int mon, int giangvien, int gio)// khoi tao  ngau nhien mot gen
        {
            this.maMon = mon;
            this.maGiangVien = giangvien;
            this.soGio = gio;
        }
        public static void SetMaxGio(int gio)
        {
            maxGio = gio;
        }
        #endregion



        #region Các hàm hỗ trợ
        public static bool operator &(ClassTietHoc a, ClassTietHoc b)//so sanh 2 doi tuong cua lop Tiet
        {

            return (a.maMon == b.maMon) && (a.maGiangVien == b.maGiangVien) && (a.soGio == b.soGio);
        }
        public static bool operator |(ClassTietHoc a, ClassTietHoc b)// so sanh bo qua so gio
        {
            return (a.maMon == b.maMon) && (a.maGiangVien == b.maGiangVien);
        }
        public static bool operator ^(ClassTietHoc a,ClassTietHoc b) // chi so sanh giang vien 
        {
            return (a.maGiangVien ==b.maGiangVien );
        }
        #endregion
        public string Show()
        {
            string st = "M:" + this.maMon.ToString() + ",GV:" + this.maGiangVien.ToString();
            return st;
        }
        public string ShowAll()
        {
            string st = "M:" + this.maMon.ToString() + ",GV:" + this.maGiangVien.ToString() + ",SG:" + this.soGio.ToString();
            return st;
        }
        #endregion
    }
}
