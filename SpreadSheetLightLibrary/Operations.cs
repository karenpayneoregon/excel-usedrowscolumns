using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace SpreadSheetLightLibrary
{
    public class Operations
    {
        public Dictionary<string, string> UsedRowsColumns;
        public Operations()
        {
            UsedRowsColumns = new Dictionary<string, string>();
        }
        public bool GetInformation(string pFileName)
        {
            using (SLDocument sl = new SLDocument(pFileName))
            {
                var sheetNames = sl.GetSheetNames(false);
                foreach (string sheetName in sheetNames)
                {
                    if (sl.SelectWorksheet(sheetName))
                    {
                        SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                        UsedRowsColumns.Add(sheetName, $"{stats.EndColumnIndex.ExcelColumnName()}:{stats.EndRowIndex}");
                    }
                }
                return UsedRowsColumns.Count == sheetNames.Count;
            }
        }
    }
}
