using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Summary.Data;

namespace Summary
{
    class Program
    {
        static void Main(string[] args)
        {
            StoreDbContext db = new StoreDbContext();

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Quick Chip Roll-Out Schedule.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;



            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;



            string todaysDate = DateTime.Now.ToShortDateString();
            Console.WriteLine("Todays Date is " + todaysDate);

            //iterate over the QuickChip mx915 upgrade date column and print to the console as it appears in the file
            Console.WriteLine("Stores Scheduled to be upgraded today are:");
            for (int row = 2; row <= rowCount; row++)
            {
                //checks if store is on todays schedule
                if (xlRange.Cells[row, 8].Value.ToString() == todaysDate + " 12:00:00 AM")
                {
                    string storeNumber = xlRange.Cells[row, 1].Value.ToString();
                    string numberOfRegisters = xlRange.Cells[row, 6].Value.ToString();
                    Console.Write(storeNumber + "\t");
                    Store store = new Store(storeNumber, numberOfRegisters);
                    db.TodaysStores.Add(store);

                }

            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            // create text file containing todays scheduled store numbers.
            string fileName = @"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\StoreList.txt";
            using (StreamWriter writer = new StreamWriter(fileName, false))
            {
                foreach (Store store in db.TodaysStores)
                {
                    writer.WriteLine(store.StoreNumber.ToString());
                }

            }

            Console.WriteLine("Executing vbs cmd");
            ExecuteStep2();
            Console.WriteLine("Executing next bat command");
            ExecuteStep3();


            //read StoreSummary file in X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\Status\**YearMoDay**\StoreSummary.csv
            string folderDate = DateTime.Now.ToString("yyyyMMdd");
            string filePath = @"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\Status\" + folderDate + @"\StoreSummary.csv";

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlStoreSummaryApp = new Excel.Application();
            Excel.Workbook xlStoreSummaryWorkbook = xlStoreSummaryApp.Workbooks.Open(filePath);
            Excel._Worksheet xlStoreSummaryWorksheet = xlStoreSummaryApp.Sheets[1];
            Excel.Range xlStoreSummaryRange = xlStoreSummaryWorksheet.UsedRange;

            

            int storeSummaryRowCount = xlStoreSummaryRange.Rows.Count;
            int storeSummaryColCount = xlStoreSummaryRange.Columns.Count;



            Console.WriteLine("Reading through Summary excellsheet:");
            for (int row = 1; row <= storeSummaryRowCount - 1; row++)
            {
                //if row is for a server. skip
                if(xlStoreSummaryRange.Cells[row, 1].Value.ToString() != null)
                {
                    string rowheader = xlStoreSummaryRange.Cells[row, 1].Value.ToString();

                    if (rowheader == "MFS")
                    {
                        //Console.WriteLine("The line was for a server. Continuing onto next line....");
                    }

                    //else row is for a pos. build pos and add it to store

                    else
                    {
                        string storenumber = xlStoreSummaryRange.Cells[row, 4].Value.ToString();
                        Store currentStore = db.TodaysStores.Single(c => c.StoreNumber.ToString() == storenumber);

                        string os = xlStoreSummaryRange.Cells[row, 11].Value.ToString();
                        string xpi = xlStoreSummaryRange.Cells[row, 12].Value.ToString();
                        string contactless = xlStoreSummaryRange.Cells[row, 13].Value.ToString();
                        POS pos = new POS(os, xpi, contactless);
                        currentStore.Registers.Add(pos);

                        //Console.WriteLine("the current store on line " + row + "=" + currentStore.ToString());
                        //Console.WriteLine("Press enter to contiune onto next line");
                        //Console.ReadLine();
                    }
                }
                else
                {
                    break;
                }
                
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlStoreSummaryRange);
            Marshal.ReleaseComObject(xlStoreSummaryWorksheet);
            xlStoreSummaryWorkbook.Close();
            Marshal.ReleaseComObject(xlStoreSummaryWorkbook);
            xlStoreSummaryApp.Quit();
            Marshal.ReleaseComObject(xlStoreSummaryApp);




            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Now printing store summary");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Store Summary:");

            foreach( var store in db.TodaysStores)
            {
                Console.WriteLine("");
                Console.WriteLine(store.ToString());
            }







            Console.WriteLine("Finished");
            Console.ReadLine();


        }

        

        public static void ExecuteStep2()
        {
            Process scriptProc = new Process();
            scriptProc.StartInfo.FileName = @"cscript";
            scriptProc.StartInfo.WorkingDirectory = @"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report\"; //<---very important 
            scriptProc.StartInfo.Arguments = "//B //Nologo Step2_CopyFiles.vbs";
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up
            scriptProc.Start();
            scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exit
            scriptProc.Close();

            string message1 = "Step2_CopyFiles.vbs file executed !!";

            Console.WriteLine(message1);
            
        }

        public static void ExecuteStep3()
        {
            string Dir = string.Format(@"X:\Group\Information Technology\Host Support\EMV QC and Contactless\Status Report");

            Process proc = new Process();
            proc.StartInfo.WorkingDirectory = Dir;
            proc.StartInfo.FileName = "Step3_CombineFiles";
            proc.StartInfo.CreateNoWindow = false;
            proc.Start();
            proc.WaitForExit();
            string message = "Step3_combineFiles.Bat file executed !!";
            Console.WriteLine(message);
        }
    }
}
