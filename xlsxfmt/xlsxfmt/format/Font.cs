using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt.format
{
    public class Font
    {
        public string Family { get; set; }
        public string Size { get; set; }
        public string Style { get; set; }
        public Header Header { get; set; }
        public Data Data { get; set; }
        public Footer Footer { get; set; }
    }
}
