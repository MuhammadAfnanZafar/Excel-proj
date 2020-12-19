using Spire.Xls.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    class SpireImplementations
    {
        public WorksheetsCollection GetAllWorksheets(string fileName)
        {

            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(fileName);
            WorksheetsCollection worksheets = workbook.Worksheets;

            return worksheets;
        }
    }
}
