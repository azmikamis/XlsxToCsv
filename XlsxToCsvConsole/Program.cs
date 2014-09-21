using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XlsxToCsv;

namespace XlsxToCsvConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            XlsxToCsvExporter xce = new XlsxToCsvExporter();

            DateTime start = DateTime.Now;

            xce.Export(@"d:\Book1.xlsx", @"d:\test");
            xce.Export(@"d:\Book1.xlsx", "Sheet1", @"d:\test");

            DateTime end = DateTime.Now;
            TimeSpan span = start - end;

            Console.WriteLine(span.ToString());
        }
    }
}
