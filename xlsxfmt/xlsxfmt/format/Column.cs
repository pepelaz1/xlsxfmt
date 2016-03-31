using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace xlsxfmt.format
{
    public class Column
    {
        public string Name { get; set; }
        public int Number { get; set; }
        public string Source { get; set; }
        public string Width { get; set; }
        [YamlMember(Alias = "format-type")]
        public string FormatType { get; set; }
        [YamlMember(Alias = "decimal-places")]
        public string DecimalPlaces { get; set; }
        [YamlMember(Alias = "date-format")]
        public string DateFormat { get; set; }
        public Font Font { get; set; }
        [YamlMember(Alias = "conditional-formatting")]
        public ConditionalFormatting conditionalFormatting { get; set; }
        public Subtotal Subtotal { get; set; }
    }
}
