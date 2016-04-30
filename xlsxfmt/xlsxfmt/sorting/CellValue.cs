using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt.sorting
{
    public class CellValue : IComparable
    {
        public string strValue { get; set; }
        public double numValue { get; set; }

        public DateTime dateValue { get; set; }

        public bool isString { get; set; }

        public bool isNumeric { get; set; }

        public bool isDate { get; set; }

        public CellValue(string value)
        {
            strValue = value;
            isString = true;
            isNumeric = false;
            isDate = false;
        }

        public CellValue(double value)
        {
            numValue = value;
            isString = false;
            isNumeric = true;
            isDate = false;
        }

        public CellValue(DateTime value)
        {
            dateValue = value;
            isString = false;
            isNumeric = false;
            isDate = true;
        }
        public int CompareTo(object obj)
        {
            if (obj is CellValue)
            {
                if (isString)
                    return strValue.CompareTo(((CellValue)obj).strValue);
                else if (isNumeric)
                    return numValue.CompareTo(((CellValue)obj).numValue);
                else if (isDate)
                    return dateValue.CompareTo(((CellValue)obj).dateValue);
                else
                {
                    return -1;
                }
            }
            else
                throw new NotImplementedException();
        }
    }
}
