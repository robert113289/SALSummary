using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace Summary.Models
{
    class XcellSchedule
    {
        public Excel.Workbook Workbook { get; set; }
        public Excel.Application Application { get; set; }

        public XcellSchedule()
        {
            Excel.Application xlApp = new Excel.Application();
            Workbook = xlApp.Workbooks.Open(@"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Quick Chip Roll-Out Schedule.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

        }

        public void Close()
        {
            Workbook.Close();
            xlApp.Quit();
        }
        

    }
}
