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
                if (xlRange.Cells[row, 8].Value == null)
                {
                    continue;
                }
                else if (xlRange.Cells[row, 8].Value.ToString() == todaysDate + " 12:00:00 AM")
                {
                    //need to add in null exception handler
                    string storeNumber = xlRange.Cells[row, 1].Value.ToString();
                    string numberOfRegisters = xlRange.Cells[row, 6].Value.ToString();
                    Console.Write(storeNumber + "\t");
                    Store store = new Store(storeNumber, numberOfRegisters)
                    {
                        XcellRowID = row
                    };
                    db.TodaysStores.Add(store);

                }

            }

            if (db.TodaysStores.Count() != 0)
            {
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
                    string rowheader = "";
                    //if row is for a server. skip
                    if (xlStoreSummaryRange.Cells[row, 1].Value.ToString() != null)
                    {
                        try
                        {
                            rowheader = xlStoreSummaryRange.Cells[row, 1].Value.ToString();
                        }
                        catch (Exception)
                        {
                            rowheader = "Null";
                        }


                        if (rowheader == "MFS")
                        {
                            //Console.WriteLine("The line was for a server. Continuing onto next line....");
                            continue;
                        }

                        //else row is for a pos. build pos and add it to store
                        else
                        {
                            if (xlStoreSummaryRange.Cells[row, 4].Value != null)
                            {
                                string storenumber = xlStoreSummaryRange.Cells[row, 4].Value.ToString();
                                Store currentStore = db.TodaysStores.Single(c => c.StoreNumber.ToString() == storenumber);
                                string xpi;
                                string os;
                                string contactless;
                                string name;

                                if (xlStoreSummaryRange.Cells[row, 11].Value != null)
                                {
                                    os = xlStoreSummaryRange.Cells[row, 11].Value.ToString();
                                }
                                else
                                {
                                    os = "Unknown";
                                }

                                if (xlStoreSummaryRange.Cells[row, 12].Value != null)
                                {
                                    xpi = xlStoreSummaryRange.Cells[row, 12].Value.ToString();
                                }
                                else
                                {
                                    xpi = "Unknown";
                                }
                                if (xlStoreSummaryRange.Cells[row, 13].Value != null)
                                {
                                    contactless = xlStoreSummaryRange.Cells[row, 13].Value.ToString();
                                }
                                else
                                {
                                    contactless = "Unknown";
                                }

                                if (xlStoreSummaryRange.Cells[row, 2].Value != null)
                                {
                                    name = xlStoreSummaryRange.Cells[row, 2].Value.ToString();
                                }
                                else
                                {
                                    name = "Unknown";
                                }

                                POS pos = new POS(os, xpi, contactless, name);

                                currentStore.Registers.Add(pos);

                                //Console.WriteLine("the current store on line " + row + "=" + currentStore.ToString());
                                //Console.WriteLine("Press enter to contiune onto next line");
                                //Console.ReadLine();
                            }

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

                foreach (var store in db.TodaysStores)
                {
                    Console.WriteLine("");
                    Console.WriteLine(store.ToString());

                    if (!store.IsUpgraded())
                    {
                        foreach (var pos in store.Registers)
                        {
                            string cellData = "";

                            if (xlWorksheet.Cells[store.XcellRowID, "K"].Value != null)
                            {
                                cellData = xlWorksheet.Cells[store.XcellRowID, "K"].Value.ToString();
                            }
                            if (!pos.IsUpgraded())
                            {
                                Console.WriteLine(pos.ToString());
                                if (cellData.Any())
                                {
                                    string newCellData = cellData + "\n" + pos.ToString();
                                    xlWorksheet.Cells[store.XcellRowID, "K"] = newCellData;
                                }
                                else
                                {
                                    xlWorksheet.Cells[store.XcellRowID, "k"] = pos.ToString();
                                }
                            }

                        }
                    }
                    //write store upgrade status to spreadsheet
                    xlWorksheet.Cells[store.XcellRowID, "J"] = store.UpgradeStatus();
                }
                xlWorksheet.Columns["K"].AutoFit();
                xlWorksheet.Columns["J"].AutoFit();
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
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
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; 
            scriptProc.Start();
            scriptProc.WaitForExit(); 
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
