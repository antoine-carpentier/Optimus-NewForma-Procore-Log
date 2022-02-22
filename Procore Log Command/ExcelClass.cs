using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewForma_Log_Command
{
    public class ExcelClass
    {

        public static void OpenExcel(string xlPath, List<string> RFINumber, List<string> RFIDescription, List<string> RFIDate, List<string> SubmittalNumber, List<string> SubmittalDescription, List<string> SubmittalDate)
        {
            if (File.Exists(xlPath))
            {
                Excel.Application xlApp = null;
                Excel.Workbooks xlWorkBooks = null;
                Excel.Workbook xlWorkBook = null;
                Excel.Worksheet xlWorkSheet = null;
                Excel.Sheets xlWorkSheets = null;
                Excel.Range xlCell = null;

                xlApp = new();
                xlApp.DisplayAlerts = false;
                xlApp.ScreenUpdating = false;
                xlApp.Visible = false;
                xlWorkBooks = xlApp.Workbooks;
                xlWorkBook = xlWorkBooks.Open(xlPath);
                xlWorkSheets = xlWorkBook.Worksheets;
                xlWorkSheet = (Excel.Worksheet)xlWorkSheets[3];
                Excel.Range xlLastCell = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                int a;
                int DataIndex = 5;
                // find the first empty row for RFIS
                for (a = 5; a <= Math.Max(5, xlLastCell.Row); a++)
                {
                    xlCell = xlWorkSheet.Cells[a, 1];
                    if (string.IsNullOrEmpty(Convert.ToString(xlCell.Value)))
                    {
                        Console.WriteLine($"cell {a} is empty.");
                        DataIndex = a;
                        break;
                    }
                    else
                    {
                        //Console.WriteLine($"cell {a} is not empty. Content = {Convert.ToString(xlCell.Value)}.");
                    }
                }

                if (RFINumber.Count > 0)
                {
                    int firstRFI = Convert.ToInt32(RFINumber[0].Split(" ")[1]);
                    int lastRFI = Convert.ToInt32(RFINumber[RFINumber.Count - 1].Split(" ")[1]);


                    //Write the RFIs data
                    if (lastRFI > firstRFI)
                    {
                        for (a = 0; a < RFINumber.Count; a++)
                        {
                            xlWorkSheet.Cells[DataIndex + a, 1].Value = RFINumber[a].Trim();
                            xlWorkSheet.Cells[DataIndex + a, 2].Value = RFIDescription[a].Trim();
                            xlWorkSheet.Cells[DataIndex + a, 3].Value = RFIDate[a];
                            Console.WriteLine(DataIndex + a + " = " + RFINumber[a].Trim() + " / " + RFIDescription[a].Trim() + " / " + RFIDate[a]);
                        }
                    }
                    else
                    {
                        for (a = 0; a < RFINumber.Count; a++)
                        {
                            xlWorkSheet.Cells[DataIndex + a, 1].Value = RFINumber[^(a + 1)].Trim();
                            xlWorkSheet.Cells[DataIndex + a, 2].Value = RFIDescription[^(a + 1)].Trim();
                            xlWorkSheet.Cells[DataIndex + a, 3].Value = RFIDate[^(a + 1)];
                            Console.WriteLine(DataIndex + a + " = " + RFINumber[^(a + 1)].Trim() + " / " + RFIDescription[^(a + 1)].Trim() + " / " + RFIDate[^(a + 1)]);
                        }
                    }

                }

                //submittal sheet
                xlWorkSheet = (Excel.Worksheet)xlWorkSheets[2];
                xlLastCell = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                //find first empty submittal row
                DataIndex = 4;
                for (a = 4; a <= Math.Max(4, xlLastCell.Row); a++)
                {
                    xlCell = xlWorkSheet.Cells[a, 1];
                    if (string.IsNullOrEmpty(Convert.ToString(xlCell.Value)))
                    {
                        Console.WriteLine($"cell {a} is empty.");
                        DataIndex = a;
                        break;
                    }
                    else
                    {
                        //Console.WriteLine($"cell {a} is not empty. Content = {Convert.ToString(xlCell.Value)}.");
                    }
                }

                //write submittal data
                if (SubmittalNumber.Count > 0)
                {
                    for (a = 0; a < SubmittalNumber.Count; a++)
                    {
                        xlWorkSheet.Cells[DataIndex + a, 1].Value = SubmittalNumber[^(a + 1)].Trim();
                        xlWorkSheet.Cells[DataIndex + a, 3].Value = SubmittalDescription[^(a + 1)].Trim();
                        xlWorkSheet.Cells[DataIndex + a, 4].Value = SubmittalDate[^(a + 1)];
                        Console.WriteLine(DataIndex + a + " = " + SubmittalNumber[^(a + 1)].Trim() + " / " + SubmittalDescription[^(a + 1)].Trim() + " / " + SubmittalDate[^(a + 1)]);
                    }
                }

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Save();
                xlWorkBook.Close();
                xlApp.UserControl = true;
                xlApp.Quit();
                ReleaseComObject(xlCell);
                ReleaseComObject(xlWorkSheets);
                ReleaseComObject(xlWorkSheet);
                ReleaseComObject(xlWorkBook);
                ReleaseComObject(xlWorkBooks);
                ReleaseComObject(xlApp);
            }
            else
            {
                Console.WriteLine("Excel file not located.");
            }
        }
        public static void ReleaseComObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
            }
        }

    }
}
