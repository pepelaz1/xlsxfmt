using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt
{
    public class RowsRange
    {
        public int startRow { get; set; }
        public int endRow { get; set; }

        public RowsRange(int start, int end)
        {
            startRow = start;
            endRow = end;
        }
    }
}
