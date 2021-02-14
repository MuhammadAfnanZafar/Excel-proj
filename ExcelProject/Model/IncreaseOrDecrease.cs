using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    class IncreaseOrDecrease
    {
        public string columnName { get; set; }
        public string columnValue { get; set; }
        public int rowIndex { get; set; }
        public int colIndex { get; set; }
        public string query { get; set; }
        public string status { get; set; }
        public static List<IncreaseOrDecrease> increaseList { get; set; }
        public static List<IncreaseOrDecrease> decreaseList { get; set; }
    }
}
