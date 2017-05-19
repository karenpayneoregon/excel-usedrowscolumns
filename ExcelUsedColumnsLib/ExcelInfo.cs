using System;

namespace ExcelUsedColumnsLib
{
    public class ExcelInfo
    {
        /// <summary>
        /// Physical filename (should include full path)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string FileName { get; set; }
        /// <summary>
        /// Sheet name in filename to get information
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string SheetName { get; set; }
        /// <summary>
        /// Last row used for sheetname
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Int32 UsedRows { get; set; }
        /// <summary>
        /// Last used column for sheetname
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Int32 UsedColumns { get; set; }
        public string LastCell { get; set; }

        /// <summary>
        /// For debugging and demoing
        /// Filename is last on purpose
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override string ToString()
        {
            return "'" + SheetName + "' Rows: " + UsedRows.ToString("d3") + " Cols: " + UsedColumns.ToString("d3") + " File: " + FileName;
        }
    }

}
