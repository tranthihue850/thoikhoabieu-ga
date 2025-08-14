using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace ThoiKhoaBieu
{
    public class ClassCaThe
    {
        private int bitCount;
        public int BitCount
        {
            get { return bitCount; }
            set { bitCount = value; }
        }

        private string propertiesGA="Gốc";

        public string PropertiesGA
        {
            get { return propertiesGA; }
            set { propertiesGA = value; }
        }

        private int decimalValue = 0;
        public int DecimalValue
        {
            get { return decimalValue; }
            set { decimalValue = value; }
        }

        private string binaryValue;
        public string BinaryValue
        {
            get { return binaryValue; }
            set { binaryValue = value; }
        }

        private double fx;

        public double Fx
        {
            get { return fx; }
            set { fx = value; }
        }
        public ClassCaThe(int bitCount)
        {
            this.bitCount = bitCount;
            getBinary();
            getDecimal();
        }

        public ClassCaThe(int bitCount,int dValue)
        {
            this.bitCount = bitCount;
            this.decimalValue = dValue;
            binaryValue = convertToBirary(decimalValue);
        }
        public ClassCaThe(int bitCount,string strB)
        {
            this.bitCount = bitCount;
            this.decimalValue = convertToDecimal(strB);
            this.binaryValue = strB;
        }
        private void getBinary()
        {
            string temp = "";
            while (temp.IndexOf("1") < 0)
            {
                temp = "";
                for (int i = 0; i < BitCount; i++)
                {
                    Random x = new Random();
                    temp += x.Next(0, 2).ToString() + " ";
                    Thread.Sleep(120);
                }
            }
            BinaryValue = temp.Trim();
        }
        private string convertToBirary(int d)
        {
            string strB = "";
            while (d != 0)
            {
                strB += d % 2 + "";
                d /= 2;
            }
            string strResult = "";
            for (int i = strB.Length-1; i >= 0; i--)
                strResult += strB[i];
            //gan them bit cho du n bits
            strB = "";
            for (int i = 1; i <= BitCount - strResult.Length; i++)
                strB += "0";
            strResult = strB + strResult;
          return strResult;
        }
        private int convertToDecimal(string strB)
        {
            int nD = 0;
            for (int i = 0; i < strB.Length; i++)
            {
                int intVal =strB[i]=='1'?1:0;
                nD += intVal * (int)Math.Pow(2, strB.Length - 1 - i);
            }
            return nD;
        }
        private void getDecimal()
        {
            string[] strArrTemp = BinaryValue.Split(' ');
            for (int i = 0; i < strArrTemp.Length; i++)
            { 
                int intVal=int.Parse(strArrTemp[i]);
                DecimalValue += intVal * (int)Math.Pow(2, strArrTemp.Length - 1 - i);
            }
        }
    }
}
