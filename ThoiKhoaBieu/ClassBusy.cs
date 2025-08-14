using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ThoiKhoaBieu
{
    public class ClassBusy
    {
        public int soGv;
        public ClassGen[] lichBan;
        public ClassBusy(RWDataExcel data)
        {
            soGv = data.soGv;
            lichBan = new ClassGen[soGv];
            for (int i = 0; i < soGv; ++i) lichBan[i] = new ClassGen(data,i);
        }
        public int TrungLichBan(ClassGen a)
        {
            int trung = 0;
            for (int i = 0; i < soGv; ++i)
                trung += a ^ lichBan[i];
            return trung;
        }
    }
}
