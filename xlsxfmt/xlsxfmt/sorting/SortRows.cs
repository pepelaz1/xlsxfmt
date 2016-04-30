using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt.sorting
{
    public class SortRows
    {
        static int maxRows = 1000000;
        public SortRow[] rows { get; set; }

        public SortRows(int capacity)
        {
            rows = new SortRow[capacity];
        }

        public void sort(int colNumber, bool isAscending)
        {
            if (rows != null && rows.Length > colNumber ){
                if (isAscending)
                    Array.Sort(rows, delegate(SortRow x, SortRow y) { return x.cellValues[colNumber].CompareTo(y.cellValues[colNumber]); });
                else
                    Array.Sort(rows, delegate(SortRow x, SortRow y) { return y.cellValues[colNumber].CompareTo(x.cellValues[colNumber]); });
			    
            }
        }
    }
}
