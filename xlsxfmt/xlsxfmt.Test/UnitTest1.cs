using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace xlsxfmt.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var xf = new XlsxFormatter( new string[] {
                @"examples\test input1.xlsx",
                @"examples\testFormat.yaml",
                @"--output-filename-prefix=_outpref_",
                @"--output-filename-postfix=_outpost_",
                @"--grand-total-prefix=""General Medical System""",
                @"examples\overriden output.xlsx",
            });

            xf.Process();
        }
    }
}
