using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class RequiredValue
    {
        [YamlMember(Alias = "required-value")]
        public String requiredValue { get; set; }
    }
}
