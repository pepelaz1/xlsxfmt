using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlsxfmt.sorting;

namespace xlsxfmt
{
    public class SortRow
    {
        public int index { get; set; }
        public List<CellValue> cellValues { get; set; } 
    }
}
