using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//after installing nuget google.gdata.spreadsheets
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;

namespace GoogleIntegration
{
    class Program
    {
        static void Main(string[] args)
        {
            var google = new GoogleSpreadsheet("", "",0);
            var list = new List<int> { 9,8,7,6,5,4,3,2,1};
            var ints = google.GetCellListInt();
            foreach (int i in ints)
            {
                Console.WriteLine(i);
            }
            //google.ClearSheet();
            google.setCells(google.ToStringList(list), new SheetDimensions(list.Count(),1));
            Console.Read();
        }
    }
}
