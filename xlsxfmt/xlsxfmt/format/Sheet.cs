using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class Sheet
    {
        public string Name { get; set; }
        public string Source { get; set; }
        [YamlMember(Alias = "freeze-on-cell")]
        public string FreezeOnCell { get; set; }   
        [YamlMember(Alias = "header-row-bgcolor")]
        public string HeaderRowBgcolor { get; set; } 
        [YamlMember(Alias = "grand-total-row-bgcolor")]
        public string GrandTotalRowBgcolor { get; set; }
        public List<Sort> Sort { get; set; }
        public string Hidden { get; set; }
        [YamlMember(Alias = "include-logo")]
        public string IncludeLogo { get; set; }
        public List<Column> Column { get; set; }
        public Font Font { get; set; }
    }
}
