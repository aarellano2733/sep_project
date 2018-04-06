using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileOutPath = @"D:\SEPs2.0\SEPs\Splice-101789.csv";
            string fileInPath = @"D:\SEPs2.0\SEPs\Splice-101789.xlsm";
            //instaniate class
            ConvertExcel convert = new ConvertExcel();
            //read from excel
            LocationInfo locInfo = convert.ReadInfoFromExcel(fileInPath);
            //write out to CSV
            convert.WriteToCSV(locInfo, fileOutPath);
        }
    }
}
