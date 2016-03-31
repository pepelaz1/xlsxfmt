using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxfmt
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                    printUsage();
                else
                {
                    Console.WriteLine("Processing...");

                    var w = new Stopwatch();
                    w.Start();

                    var xp = new XlsxFormatter(args);
                    xp.Process();

                    w.Stop();
                    Console.WriteLine("Ok");
                    Console.WriteLine("Time elapsed: {0}:{1}:{2}:{3}", 
                        w.Elapsed.Hours.ToString("D2"),
                        w.Elapsed.Minutes.ToString("D2"),
                        w.Elapsed.Seconds.ToString("D2"),
                        w.Elapsed.Milliseconds.ToString("D3"));
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void printUsage()
        {
            Console.WriteLine(@"usage: xlsxfmt source format [options] [output]");
            Console.Write(Environment.NewLine);
            Console.WriteLine(@"source (Excel file)");
            Console.WriteLine(@"format (yaml formatting description file)");
            Console.WriteLine(@"[output] output file name (overrides file name determined from values in format and options)");
            Console.Write(Environment.NewLine);
            Console.WriteLine(@"Options:");
            Console.WriteLine(@"--output-filename-prefix=prefix     prefix to be added to the beginning of the output file name");
            Console.WriteLine(@"--output-filename-postfix=postfix   postfix to be added to the end of the output file name (before the extension)");
            Console.WriteLine(@"--grand-total-prefix=prefix         string to be prepended to the ""Grand..."" on the last line of any subtotaling");
        }
    }
}
