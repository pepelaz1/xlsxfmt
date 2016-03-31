using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class Subtotal
    {
        public string Group { get; set; }
        [YamlMember(Alias = "total-row-bgcolor")]
        public string TotalRowBgcolor { get; set; }
        public string Function { get; set; }
        
    }
}
