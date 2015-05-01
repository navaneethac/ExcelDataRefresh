using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelDataRefresh
{
    /// <summary>
    /// Tool uses Microsoft Excel 15.0 Object Library for Excel manipulation
    /// Note: When this tool is scheduled using windows scheduler you may observe below error
        // Error: Microsoft Excel cannot access the file <<filename>>. There are several possible reasons:
            //• The file name or path does not exist.
            //• The file is being used by another program.
            //• The workbook you are trying to save has the same name as a currently open workbook.
    /// Verify if user has required permissions or not. If user has permissions but still this error is observed follow below step.
    /// create 'desktop' folder in below path based on your Office Excel installed version (i.e. 32bit or 64bit)  
    /// 32bit  -  c:\windows\system32\config\systemprofile\desktop 
    /// 64bit -  c:\windows\syswow64\config\systemprofile\desktop
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            string inputpath = "";
            if (args.Length == 0)
            {
                Console.WriteLine("Enter the absolute path of the Excel file:");
                Console.WriteLine("Sample path: http://sample.sharepoint.com/aaa/filename.xlsx");
                inputpath = Console.ReadLine();
            }
            else if (args.Length == 1)
            {
                inputpath = args[0];
            }
            else
            {
                Console.WriteLine("Incorrect number of parameters passed.");
            }

            string path = Util.EncodePath(inputpath);
            Application excel = null;

            try
            {

                excel = new Application();

                excel.Visible = false;

                excel.DisplayAlerts = false;

                Console.WriteLine("Opening the Excel file:");

                Workbook excelbook = excel.Workbooks.Open(path);

                Console.WriteLine("Excel file opened.");

                if (excelbook.Connections.Count == 0)
                {
                    Console.WriteLine("There are no data connections to refresh the data in this Excel file.");
                    excelbook.Close();
                    return;
                }

                if (Util.IsLocal(path))
                {
                    Console.WriteLine("Refreshing the data connections:");
                    excelbook.RefreshAll();
                    Console.WriteLine("Data Connections Refreshed.");
                    Console.WriteLine("Saving the Excel file:");
                    excelbook.Save();
                    Console.WriteLine("Excel file saved:");
                }
                else
                {
                    if (excel.Workbooks.CanCheckOut(path))
                    {
                        Console.WriteLine("Checking out the excel file:");
                        excel.Workbooks.CheckOut(path);
                        Console.WriteLine("Checked out.");

                        Console.WriteLine("Refreshing the data connections:");
                        excelbook.RefreshAll();
                        Console.WriteLine("Data Connections Refreshed.");

                        Console.WriteLine("Checking in the excel file:");
                        excelbook.CheckInWithVersion(Comments: "Auto Refresh.");
                        Console.WriteLine("Excel file checked in.");

                    }
                    else
                    {
                        excelbook.Close();
                        Console.WriteLine("Checkout failed. Possibly due to:");
                        Console.WriteLine("     CheckIn/Checkout not enabled on SharePoint site.");
                        Console.WriteLine("     User has no permission to edit the file.");
                    }
                }               
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occured: " + ex.Message);
                // Enable this code to write to event log for debugging/troubleshooting
                //if (!EventLog.SourceExists("ExcelDataRefresh"))
                //    EventLog.CreateEventSource("ExcelDataRefresh", "Application");

                //EventLog.WriteEntry("ExcelDataRefresh", ex.Message, EventLogEntryType.Error);                
            }
            finally
            {
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                // To release the COM reference otherwise COM may hold reference open 
                // since the reference count is not zero
                excel = null;
                // Calling GC multiple times to make sure that all Excel references are cleaned
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
    }
}
