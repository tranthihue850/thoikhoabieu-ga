using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace ThoiKhoaBieu
{
    public class ClassQuanThe
    {
        //Data- Property
        #region Data
        // dữ liệu của quần thể
        private int size, realSize;
        private ClassNhiemSacThe[] quanThe;
        private double[] Ps;
        private int[] Ps2;
        private int tsToPs2 = 10000;
        //tham số GA
        private double Pc, Pm;// xác suất lai,đột biến
        // tham số chỉ ra đang ở thế hệ thứ mây
        private int doi;
        // đường dẫn kết suất dữ liệu
        private string path;
        //đối tượng kết suất
        private RWDataExcel outData, excelData;
        // lich ban giang vien
        public ClassBusy busy;
        #endregion
        //----------------------------------------------------------------------------
        // Method
        #region Method

        #region Các hàm khởi tạo
        public ClassQuanThe()
        { }
        public ClassQuanThe(int Size, int[] lop, RWDataExcel data, int soNgayHocTrongTuan, int soGioMotBuoi, int soGioMotTiet, double sxLai, double sxDotBien)// sử dụng phép toán khởi tạo ngỗng nhiên
        {
            // đặt giá trị Count số đếm cá thể trong quần thể=0
            ClassNhiemSacThe.Count = 0;
            // khởi tạo vùng nhớ cho quần thể
            this.size = Size;
            this.realSize = Size;
            this.quanThe = new ClassNhiemSacThe[this.size * 3];// số cá thể trong quần thể bùm nổ tối đa gấp 3 lần ban đầu khi xs lai=đột biến=100%
            this.Pc = sxLai;
            this.Pm = sxDotBien;
            this.Ps = new double[this.size];
            this.Ps2 = new int[this.size * 3];
            this.doi = 0;
            this.excelData = data;            
            //khởi tạo từng cá thể 
            // khoi tao nhung tham so static 
            ClassTietHoc.SetMaxGio(soGioMotTiet);
            ClassNgay.SetMaxGio(soGioMotBuoi);
            ClassGen.SetSoNgayHocTrongTuan(soNgayHocTrongTuan);

            //khoi tao lich ban sau khi da khoi tao cac tham so khac
            CreateTs();
            busy = new ClassBusy(data);
            for (int i = 0; i < this.realSize; ++i)
                quanThe[i] = new ClassNhiemSacThe(lop, data);
            // gọi hàm mục tiêu tính độ thích nghi  cho từng cá thể trong quần thể
            HamMucTieu();

        }
        private void HamMucTieu()// hàm mục tiêu cho bài toán
        {

            // tính lại độ thích nghi cho toan bộ cá thể trong quần thể phục vụ cho phép chọn
            for (int i = 0; i < this.realSize; ++i)
            {
                MucTieu(i);
            }
            //sử dụng phương pháp loại bỏ những cá thể giống nhau ra khỏi quần thể
            for (int i = 0; i < this.realSize - 1; ++i)
                for (int j = i + 1; j < this.realSize; ++j)
                    if (quanThe[i] & quanThe[j]) quanThe[j].doThichNghi = 0;

            // tính tổng độ thích nghi toàn quần thể ban đầu
            double tong = 0;
            for (int i = 0; i < this.realSize; ++i)
                tong += quanThe[i].doThichNghi;
            if (tong == 0) tong = 0.1;
            // thiết đặt sx sinh tồn cho từng cá thể trong quần thể
            for (int i = 0; i < this.realSize; ++i)
                quanThe[i].SetXacSuatSong(quanThe[i].doThichNghi / tong);

            // thiết đặt  sx sinh tồn cho từng cá thể trong quần thể voi he so nhan tsToPs2

            for (int i = 0; i < this.realSize; ++i)
                Ps2[i] = Convert.ToInt32(quanThe[i].xacSuatSong * tsToPs2);
        }
        private void SetXacSuatSong()// thiết đặt khả năng sông của mỗi NST
        {
            double tong = 0;
            for (int i = 0; i < this.size; ++i)
                tong += quanThe[i].doThichNghi;
            if (tong == 0) tong = 0.1;
            // thiết đặt sx sống cho từng cá thể trong quần thể
            for (int i = 0; i < this.size; ++i)
                quanThe[i].SetXacSuatSong(quanThe[i].doThichNghi / tong);

            // thiết đặt cộng dồn sx sống cho từng cá thể trong quần thể
            Ps[0] = quanThe[0].xacSuatSong;
            for (int i = 1; i < this.size; ++i)
                Ps[i] = Ps[i - 1] + quanThe[i].xacSuatSong;

        }
        #endregion

        #region Hàm tiến hóa
        public void ChonLocTuNhien()
        {
            HamMucTieu();
            //   ++doi; KetXuatDuLieuTungDoi();
            /*  int tongCon = tsToPs2;
              int chon,tong,j;
              NST tg = new NST();// bien trung gian ho cho su doi cho
              for(int i=0;i<this.size ;++i)
              {
                  chon = NST.ran.Next(tongCon);
                  j = i;
                  tong=Ps2[j];
                  while ((tong < chon)&&(j<(realSize-1) ))
                  {
                      ++j;
                      tong += Ps2[j];
                  }
                  //doi cho
                  tg = quanThe[i];
                  quanThe[i] = quanThe[j];
                  quanThe[j] = tg;
                  // dat lai tonCon va Ps2
                  tongCon -= Ps2[j];
                  Ps2[j] = Ps2[i];                
              }*/
            double[] tg = new double[realSize];
            for (int i = 0; i < realSize; ++i)
                tg[i] = -quanThe[i].doThichNghi;
            Array.Sort(tg, quanThe);
            this.realSize = this.size;

        }
        public void DotBienQuanThe()
        {
            // tính số cá thể sẽ phải tạo ra= phép đột biến
            int countDB = Convert.ToInt32(size * Pm);
            // tạo ta conntDB con đột biến đưa vào quân thể
            int viTriBo = 0;
            for (int i = 0; i < countDB; ++i)
            {
                // tìm bố
                //viTriBo = Array.BinarySearch(Ps, NST.ran.NextDouble());
                //if (viTriBo < 0) viTriBo = Math.Abs(viTriBo) - 1;               
                // Đột biến  
                viTriBo = ClassNhiemSacThe.ran.Next(this.size);
                quanThe[realSize] = !quanThe[viTriBo];
                ++realSize;
            }
        }
        public void LaiTaoQuanThe()
        {
            // tính số cá thể sẽ phải tạo ra= phép lai
            int countLai = Convert.ToInt32(size * Pc);
            // tạo ta conntlai con lai đưa vào quân thể
            int viTriBo = 0, viTriMe = 0;
            for (int i = 0; i < countLai; ++i)
            {
                // tìm bố
                //viTriBo = Array.BinarySearch(Ps, NST.ran.NextDouble());
                //if (viTriBo < 0) viTriBo = Math.Abs(viTriBo) - 1;
                viTriBo = ClassNhiemSacThe.ran.Next(this.size);
                // tìm mẹ
                do
                {
                    //viTriMe = Array.BinarySearch(Ps, NST.ran.NextDouble());
                    //if (viTriMe < 0) viTriMe = Math.Abs(viTriMe) - 1;
                    viTriMe = ClassNhiemSacThe.ran.Next(this.size);
                } while (viTriBo == viTriMe);
                // lai tạo                 
                quanThe[realSize] = quanThe[viTriBo] + quanThe[viTriMe];
                ++realSize;
            }
        }
        public void MucTieu(int i)// goi muc tieu cho mot ca the
        {
            // đóng vai trò môi  trường đánh giá khả năng thích nghi của một cá thể    
            quanThe[i].SetDoThichNghi(quanThe[i].TrungLap(), quanThe[i].HocQua1Lan(),quanThe [i].TrungLichBan(busy ));          

        }

        #region các phương pháp thực hiện tiến hóa
        public void TienHoa(int soLan)// tien hóa quần thể sau k lần
        {
            this.doi = 0;
            for (int k = 0; k < soLan; ++k)
            {
                ++doi;
                SetXacSuatSong();
                LaiTaoQuanThe();
                DotBienQuanThe();
                ChonLocTuNhien();
            }
        }
        public void TienHoaTuBuoc()
        {
            ++doi;
            SetXacSuatSong();
            LaiTaoQuanThe();
            DotBienQuanThe();
            ChonLocTuNhien();
        }
        public void TienHoaKetXuatDuLieu(int soLan, string pa)
        {
            this.doi = 0;
            this.path = pa;
            ++doi;
            KetXuatDuLieuTungDoi();
            for (int k = 0; k < soLan; ++k)
            {
                //++doi;                KetXuatDuLieuTungDoi();
                SetXacSuatSong();
                LaiTaoQuanThe();
                //++doi;                KetXuatDuLieuTungDoi();
                DotBienQuanThe();
                HamMucTieu();
                ChonLocTuNhien();
                ++doi;
                KetXuatDuLieuTungDoi();
            }
            outData.Dispose();
        }
        public void TienHoaKetXuatDuLieuHoanHao(string pa)
        {
            this.doi = 0;
            this.path = pa;
            ++doi;
            KetXuatDuLieuTungDoi();
            while (quanThe[0].doThichNghi < ClassNhiemSacThe.maxDoThichNghi)
            {
                //++doi;                KetXuatDuLieuTungDoi();
                SetXacSuatSong();
                LaiTaoQuanThe();
                //++doi;                KetXuatDuLieuTungDoi();
                DotBienQuanThe();
                HamMucTieu();
                ChonLocTuNhien();
                ++doi;
                //  KetXuatDuLieuTungDoi();
            }
            KetXuatDuLieuTungDoi();
            outData.Dispose();
        }
        public void CreateTs()
        {
            this.Pm *= 10;
            if (Pm > 1) Pm = 1;
        }

        public void TienHoaKetXuatDuLieuDoiCuoi(int soLan, string pa)
        {
            this.doi = 0;
            this.path = pa;
            ++doi;
            KetXuatDuLieuTungDoi();
            for (int k = 0; k < soLan; ++k)
            {
                LaiTaoQuanThe();
                DotBienQuanThe();
                ChonLocTuNhien();
            }
            doi = soLan;
            KetXuatDuLieuTungDoi();
            outData.Dispose();
        }
        #endregion

        #endregion

        #region Các hàm hỗ trợ
        public void KetXuatDuLieuTungDoi()
        {
            if (this.doi == 1)
            {
                //nếu  là đời đầu tiên thì tạo tiêu đề cho nó và khởi tạo đối tượng kết xuất
                outData = new RWDataExcel(this.path);
                outData.CreteTitleForOutPutExcelFile();
                outData.Show();
            }
            // thông tin cho từng đời
            outData.CreateTitleForOneStepInExcelOutPut(this.doi);
            for (int i = 0; i < this.realSize; ++i)
                outData.OutPutDataForOneStep(i, quanThe[i].ID, quanThe[i].ShortShow(), quanThe[i].doThichNghi, quanThe[i].mucTieu.ToString(), quanThe[i].mucTieu2.ToString(),quanThe [i].mucTieu3.ToString(), quanThe[i].xacSuatSong, quanThe[i].nguonGoc);
        }
        public void KetXuatKetQuaCuoiCung(string pa)
        {
            RWDataExcel moi = new RWDataExcel(pa);
            moi.Show();
            moi.OutPutTitleEndOfData(quanThe[0], excelData);
            moi.Dispose();
        }

        #endregion

        #endregion
    }
}
