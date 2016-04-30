using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt
{
    public class SortKeyLevel
    {
        public String key { get; set; }
        public int grouplevel { get; set; }

        public SortKeyLevel(String k, int lvl)
        {
            key = k;
            grouplevel = lvl;
        }
    }
}
