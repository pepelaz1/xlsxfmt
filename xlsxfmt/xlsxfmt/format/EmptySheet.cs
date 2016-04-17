using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class EmptySheet
    {
        public String exclude { get; set; }

        [YamlMember(Alias = "default-value")]
        public String defaultValue { get; set; }
    }
}
