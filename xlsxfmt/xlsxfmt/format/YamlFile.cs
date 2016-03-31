using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt.format
{
    public class YamlFile
    {
        public Format Format { get; set; }
        public Defaults Defaults { get; set; }
        public List<Sheet> Sheet { get; set; }
    }
}
