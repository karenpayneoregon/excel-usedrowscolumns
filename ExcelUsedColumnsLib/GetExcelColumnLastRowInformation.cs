using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelUsedColumnsLib
{
    public class GetExcelColumnLastRowInformation
    {

        public List<string> GetSheets(string FileName)
        {
            List<string> sheetNames = new List<string>();
            bool Success = true;

            if (!File.Exists(FileName))
            {
                Exception ex = new Exception("Failed to locate '" + FileName + "'");
                throw ex;
            }

            Excel.Application xlApp = null;
            Excel.Workbooks xlWorkBooks = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Workbook xlActiveRanges = null;
            Excel.Names xlNames = null;
            Excel.Sheets xlWorkSheets = null;

            try
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkBooks = xlApp.Workbooks;
                xlWorkBook = xlWorkBooks.Open(FileName);

                xlActiveRanges = xlApp.ActiveWorkbook;
                xlNames = xlActiveRanges.Names;

                xlWorkSheets = xlWorkBook.Sheets;

                for (int x = 1; x <= xlWorkSheets.Count; x++)
                {
                    Excel.Worksheet Sheet1 = (Excel.Worksheet)xlWorkSheets[x];
                    sheetNames.Add(Sheet1.Name);
                    Marshal.FinalReleaseComObject(Sheet1);
                    Sheet1 = null;
                }

                xlWorkBook.Close();
                xlApp.UserControl = true;
                xlApp.Quit();

            }
            catch (Exception ex)
            {
                Success = false;
            }
            finally
            {

                if (xlWorkSheets != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkSheets);
                    xlWorkSheets = null;
                }

                if (xlNames != null)
                {
                    Marshal.FinalReleaseComObject(xlNames);
                    xlNames = null;
                }

                if (xlActiveRanges != null)
                {
                    Marshal.FinalReleaseComObject(xlActiveRanges);
                    xlActiveRanges = null;
                }
                if (xlActiveRanges != null)
                {
                    Marshal.FinalReleaseComObject(xlActiveRanges);
                    xlActiveRanges = null;
                }

                if (xlWorkBook != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }

                if (xlWorkBooks != null)
                {
                    Marshal.FinalReleaseComObject(xlWorkBooks);
                    xlWorkBooks = null;
                }

                if (xlApp != null)
                {
                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;
                }
            }

            return sheetNames;

        }
        /// <summary>
        /// Used to return the last used row for each column within the range of ColumnCount
        /// </summary>
        /// <param name="FileName">Existing Excel file</param>
        /// <param name="SheetName">Name of sheet in FileName</param>
        /// <param name="ColumnCount">How many columns to get data for</param>
        /// <returns></returns>
        /// <remarks>
        /// In regards to ColumnCount, passing 3 would populate the Dictionary with columns A thru C etc.
        /// </remarks>
        public Dictionary<string, int> UsedColumns(string FileName, string SheetName, int ColumnCount)
        {
            Dictionary<string, int> ColumnData = new Dictionary<string, int>();

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


                if (xlWorkSheet.Name == SheetName)
                {

                    Excel.Range xlCells = xlWorkSheet.Cells;
                    Excel.Range xlTempRange1 = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                    var xlTempRange2 = xlWorkSheet.Rows;


                    for (int Col = 1; Col <= ColumnCount; Col++)
                    {

                        var xlTempRange3 = xlWorkSheet.Range[Col.ExcelColumnName() + xlTempRange2.Count];
                        var xlTempRange4 = xlTempRange3.End[Excel.XlDirection.xlUp];

                        ColumnData.Add(Col.ExcelColumnName(), xlTempRange4.Row);
                        Marshal.FinalReleaseComObject(xlTempRange4);
                        xlTempRange4 = null;

                        Marshal.FinalReleaseComObject(xlTempRange3);
                        xlTempRange3 = null;
                    }

                    Marshal.FinalReleaseComObject(xlTempRange2);
                    xlTempRange2 = null;

                    Marshal.FinalReleaseComObject(xlTempRange1);
                    xlTempRange1 = null;

                    Marshal.FinalReleaseComObject(xlCells);
                    xlCells = null;

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

            return ColumnData;
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
            catch (Exception)
            {
                obj = null;
            }
        }
    }

}
