using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class StopValue
    {
         [YamlMember(Alias = "stop-value")]
        public String stopValue { get; set; }
    }
}
