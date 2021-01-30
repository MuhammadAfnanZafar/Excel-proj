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
        public decimal Sum { get; set; }

        public static List<List<Percentages>> percentagesList = new List<List<Percentages>>();

        public List<Percentages> combinePercentagesList(List<List<Percentages>> percentagesList)
        {
            List<Percentages> lst = new List<Percentages>();
            foreach (var item in percentagesList)
            {
                foreach (var item2 in item)
                {
                    if (item2.ColumnValue != null)
                    {
                        lst.Add(item2);
                    }
                }
            }
            return lst;
        }
    }
}
