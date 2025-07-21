using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ThoiKhoaBieu
{
    public class Ngay
    {
        //Data- Property
        #region Data
        public static int maxGio=6;
        public Tiet[] tiet;
        public Tiet[] ssTiet;// dung de so sanh cac tiet trong ngay;
        public int realLength = 0;
        #endregion
        // Method
        #region Method
        #region Các hàm khởi tạo
        public Ngay(string ban,int gv)
        {
            this.tiet = new Tiet[Ngay.maxGio];
            for (int i = 0; i < Ngay.maxGio; ++i)
                tiet[i] = null;
            this.ssTiet = new Tiet[Ngay.maxGio];
            string[] st = ban.Split(',');
            int[] vitri = new int[st.Length];           
            for (int i = 0; i < vitri.Length; ++i)
                try
                {
                    vitri[i] = Convert.ToInt32(st[i]);
                }
                catch (Exception ett) { ett.ToString(); }
            Array.Sort(vitri);
            bool[] ok = new bool[Ngay.maxGio];
            for (int i=0;i<vitri .Length ;++i)
                try
                {
                    ok[vitri[i]-1] = true;
                }
                catch (Exception e) { e.ToString(); }
            for (int i = 0; i < Ngay.maxGio; ++i)
            {
                if (ok[i]) Add(new Tiet(-1, gv, 1));
                else Add (new Tiet (-1,-1,1));
            }

        }
        public Ngay()
        {
            this.tiet = new Tiet[Ngay.maxGio];
            for (int i = 0; i < Ngay.maxGio; ++i)
                tiet[i] = null;
            this.ssTiet = new Tiet[Ngay.maxGio];

        }
        public bool Add(Tiet them)//them 1 Tiet vao mot ngay, neu vuot qua so h toi da phai hoc trong 1 ngay thi khong them dc
        {
            bool ok = false;
            int tongGio = 0;
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (tiet[i] != null) tongGio += tiet[i].soGio;
                else
                {
                    if ((tongGio + them.soGio) <= Ngay.maxGio)
                    {
                        tiet[i] = them;
                        ok = true;
                        break;
                    }
                }
            return ok;// neu them thanh cong tra lai true nguoc lai tra lai false 

        }
        public void ChangaeOneFirst(int maMonDB,int maGVDB)
        {
            int end = 0;
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (tiet[i] == null) { end = i; break; }
            if (end > 0)
            {
                if (tiet[0].soGio > 1)
                {
                    tiet[0].soGio -= 1;
                    for (int i = end; i > 0; --i)
                        tiet[i] = tiet[i - 1];
                    tiet[0] = new Tiet(maMonDB, maGVDB, 1);
                }
                else
                {
                    tiet[0].maGiangVien = maGVDB;
                    tiet[0].maMon = maMonDB;
                }
            }
        }
        public void ChangaeOneLast(Tiet sh)
        {
            int end = 0;
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (tiet[i] == null) { end = i; break; }
            if (end > 0)
            {
                if (tiet[end - 1].soGio > 1)
                {
                    tiet[end - 1].soGio -= 1;
                    Add(sh);
                }
                else if (tiet[end - 1].soGio == 1) tiet[end - 1] = sh;
            }
            else Add(sh);
            
        }
        public void SetRealLength()
        {
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (tiet[i] != null) this.realLength = i;
                else break;
            ++this.realLength;
            
        }
        public void CreateCompareTiet() // ham duoi ra de so sanh
        {
            int j = -1;
            for (int i = 0; i < Ngay.maxGio; ++i)
            {
                if (tiet[i] != null)
                {
                    for (int k = 0; k < tiet[i].soGio; ++k)
                    {
                        ++j;
                        ssTiet[j] = tiet[i];
                    }
                }
                else break;
            }
        }
        public static void SetMaxGio(int t)
        {
            maxGio = t;
        }
        public string Show()
        {
            string st = "";
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (tiet[i] != null) st += tiet[i].ShowAll() + ", ";
                else break;
            st += "\n";
            for (int i = 0; i < Ngay.maxGio; ++i)
                if (ssTiet[i] != null) st += ssTiet[i].Show() + ", ";
                else break;
            return st;

        }
        public static bool operator &(Ngay a, Ngay b) // so sanh 3 tham so
        {
            bool ok = true;
            if (a.tiet.Length != b.tiet.Length) ok = false;
            for (int i = 0; i < a.tiet.Length; ++i)
                if (a.tiet[i] != null)
                {
                    if (b.tiet[i] == null) ok = false;
                    else
                    {
                        if ((a.tiet[i] & b.tiet[i]) == false)
                        {
                            ok = false;
                            break;
                        }
                    }
                }
                else break;
            return ok;
        }
        public static int operator |(Ngay a, Ngay b) // so sanh 2 tham so
        {
            int trungNhau = 0;
            a.CreateCompareTiet();
            b.CreateCompareTiet();
            for (int i = 0; i < a.ssTiet.Length; ++i)
                if ((a.ssTiet[i] != null) && (b.ssTiet[i] != null))
                {
                    if (a.ssTiet[i] | b.ssTiet[i]) ++trungNhau;
                }
                else break;
         
            return trungNhau;
        }
        public static int operator ^(Ngay a, Ngay b) // so sanh 1 tham so
        {
            int trungNhau = 0;
            a.CreateCompareTiet();
            b.CreateCompareTiet();
            for (int i = 0; i < a.ssTiet.Length; ++i)
                if ((a.ssTiet[i] != null) && (b.ssTiet[i] != null))
                {
                    if (a.ssTiet[i] ^ b.ssTiet[i]) ++trungNhau;
                }
                else break;
            return trungNhau;
        }
        public int HocQua1Lan() // kiem tra 1 mon co hoc qua' 1 lan/ngay ko ? de toi uu ham muc tieu 2
        {
            int trung = 0, k = 0;
            for (k = 0; k < this.tiet.Length; ++k)
                if (tiet[k] == null) break;
            for (int i = 0; i < k - 1; ++i)
                for (int j = i + 1; j < k; ++j)
                    if ((tiet[i] | tiet[j])&&((tiet[i].soGio +tiet[j].soGio )>Tiet.maxGio)) ++trung;
            return trung;
        }
        #endregion



        #endregion
    }
}
