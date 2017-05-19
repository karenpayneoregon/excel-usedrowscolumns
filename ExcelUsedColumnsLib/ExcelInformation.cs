using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System;

namespace ExcelUsedColumnsLib
{
    public class ExcelInformation
    {
        public List<ExcelInfo> GetUsed(string FileName, List<string> Sheets)
        {
            List<ExcelInfo> Results = new List<ExcelInfo>();

            int RowsUsed = -1;
            int ColsUsed = -1;

            if (File.Exists(FileName))
            {
                Excel.Application xlApp = null;
                Excel.Workbooks xlWorkBooks = null;
                Excel.Workbook xlWorkBook = null;
                Excel.Worksheet xlWorkSheet = null;
                Excel.Sheets xlWorkSheets = null;

                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkBooks = xlApp.Workbooks;
                xlWorkBook = xlWorkBooks.Open(FileName);

                xlApp.Visible = false;

                xlWorkSheets = xlWorkBook.Sheets;

                for (int x = 1; x <= xlWorkSheets.Count; x++)
                {

                    xlWorkSheet = (Excel.Worksheet)xlWorkSheets[x];

                    foreach (var SheetName in Sheets)
                    {

                        if (xlWorkSheet.Name == SheetName)
                        {

                            Excel.Range xlCells = xlWorkSheet.Cells;
                            Excel.Range xlTempRange = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                            RowsUsed = xlTempRange.Row;
                            ColsUsed = xlTempRange.Column;

                            Results.Add(new ExcelInfo
                            {
                                FileName = FileName,
                                SheetName = SheetName,
                                UsedRows = RowsUsed,
                                UsedColumns = ColsUsed,
                                LastCell = $"{ColsUsed.ExcelColumnName()}:{RowsUsed}"
                            });

                            Marshal.FinalReleaseComObject(xlTempRange);
                            xlTempRange = null;

                            Marshal.FinalReleaseComObject(xlCells);
                            xlCells = null;

                        }
                    }
                    Marshal.FinalReleaseComObject(xlWorkSheet);
                    xlWorkSheet = null;
                }

                xlWorkBook.Close();
                xlApp.UserControl = true;
                xlApp.Quit();

                ReleaseComObject(xlWorkSheets);
                ReleaseComObject(xlWorkSheet);
                ReleaseComObject(xlWorkBook);
                ReleaseComObject(xlWorkBooks);
                ReleaseComObject(xlApp);

                return Results;

            }
            else
            {
                throw new Exception("'" + FileName + "' not found.");
            }

            return Results;

        }
        private void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                }
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
        }
    }

}
