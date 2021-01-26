using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    class Percentages
    {
        public string ColumnValue { get; set; }
        public string Percentage { get; set; }
        public string Query { get; set; }

        public static List<List<Percentages>> percentagesList = new List<List<Percentages>>();
    }
}
