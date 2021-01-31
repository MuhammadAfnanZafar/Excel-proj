using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    class Range
    {
        public int Min { get; set; }
        public int Max { get; set; }
        public decimal Percentage { get; set; }

        public static List<Range> ranges { get; set; }

        public Range()
        {
            ranges = new List<Range>();
        }
    }
}
