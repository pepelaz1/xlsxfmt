using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class Format
    {
       public string Name { get; set; }
       public string Description { get; set; }
       public string Version { get; set; }
       [YamlMember(Alias = "logo-path")]
       public string LogoPath { get; set; }
       [YamlMember(Alias = "output-filename-base")]
       public string OutputFilenameBase { get; set; }
    }
}
