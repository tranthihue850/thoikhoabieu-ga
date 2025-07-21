using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ThoiKhoaBieu
{
    public class Busy
    {
        public int soGv;
        public Gen[] lichBan;
        public Busy(RWDataExcel data)
        {
            soGv = data.soGv;
            lichBan = new Gen[soGv];
            for (int i = 0; i < soGv; ++i) lichBan[i] = new Gen(data,i);
        }
        public int TrungLichBan(Gen a)
        {
            int trung = 0;
            for (int i = 0; i < soGv; ++i)
                trung += a ^ lichBan[i];
            return trung;
        }
    }
}
