using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelRWApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            string str;
            string action = "NoValue";
            int rCnt = 0;
            int cCnt = 0;

            try
            {
                action = "Reading Excel file";
                xlApp = new Application();
                xlWorkBook = xlApp.Workbooks.Open("ExcelRWBook.xls", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                Console.WriteLine("Row Count =" + range.Rows.Count);
                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    action = "Reading Column 1 with Row number"+rCnt;
                    str = (string)(range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2;
                    //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    //{
                    //    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    //    //MessageBox.Show(str);
                    //}
                    if (str.Equals("Draft"))
                    {
                        action = "Draft";
                        Console.WriteLine("Running function for Draft");
                        xlWorkSheet.Cells[rCnt, 2] = DraftFucntion();
                        xlWorkBook.Save();
                    }
                    else if (str.Equals("Compose Email"))
                    {
                         action = "Compose Email";
                         Console.WriteLine("Running function for Compose Email");
                         xlWorkSheet.Cells[rCnt, 2] = ComposeEmailFucntion();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 1"))
                    {
                         action = "Other action1";
                         Console.WriteLine("Running function for Other action1");
                         xlWorkSheet.Cells[rCnt, 2] = Function1();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 2"))
                    {
                         action = "Other action2";
                         Console.WriteLine("Running function for Other action2");
                         xlWorkSheet.Cells[rCnt, 2] = Function2();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 3"))
                    {
                         action = "Other action3";
                         Console.WriteLine("Running function for Other action 3");
                         xlWorkSheet.Cells[rCnt, 2] = Function3();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 4"))
                    {
                         action = "Other action 4";
                         Console.WriteLine("Running function for Other action4");
                         xlWorkSheet.Cells[rCnt, 2] = Function4();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 5"))
                    {
                         action = "Other action5";
                         Console.WriteLine("Running function for Other action5");
                         xlWorkSheet.Cells[rCnt, 2] = Function5();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 6"))
                    {
                        action = "Other action6";
                        Console.WriteLine("Running function for Other action6");
                        xlWorkSheet.Cells[rCnt, 2] = Function6();
                        xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 7"))
                    {
                         action = "Other action7";
                         Console.WriteLine("Running function for Other action7");
                         xlWorkSheet.Cells[rCnt, 2] = Function7();
                         xlWorkBook.Save();
                    }
                    else if (str.Equals("Other action 8"))
                    {
                        action = "Other action8";
                        Console.WriteLine("Running function for Other action8");
                        xlWorkSheet.Cells[rCnt, 2] = Function8();
                        xlWorkBook.Save();
                    }
                }

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in executing action "+ action);
                Console.WriteLine(ex.ToString());
            }
        }
        private static string DraftFucntion()
        {
            return "Success";
        }
        private static string ComposeEmailFucntion()
        {
            return "Success";
        }
        private static string Function1()
        {
            return "Success";
        }
        private static string Function2()
        {
            return "Success";
        }
        private static string Function3()
        {
            return "Success";
        }
        private static string Function4()
        {
            return "Success";
        }
        private static string Function5()
        {
            return "Success";
        }
        private static string Function6()
        {
            return "Success";
        }
        private static string Function7()
        {
            return "Success";
        }
        private static string Function8()
        {
            return "Success";
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        } 
    }
}
