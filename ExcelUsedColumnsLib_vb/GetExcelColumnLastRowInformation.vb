Option Strict On
Option Infer On
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Public Class GetExcelColumnLastRowInformation

    Public Function GetSheets(ByVal FileName As String) As List(Of String)
        Dim sheetNames As New List(Of String)
        Dim Success As Boolean = True

        If Not IO.File.Exists(FileName) Then
            Dim ex As New Exception("Failed to locate '" & FileName & "'")
            Throw ex
        End If

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlActiveRanges As Excel.Workbook = Nothing
        Dim xlNames As Excel.Names = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing

        Try
            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)

            xlActiveRanges = xlApp.ActiveWorkbook
            xlNames = xlActiveRanges.Names

            xlWorkSheets = xlWorkBook.Sheets

            For x As Integer = 1 To xlWorkSheets.Count
                Dim Sheet1 As Excel.Worksheet = CType(xlWorkSheets(x), Excel.Worksheet)
                sheetNames.Add(Sheet1.Name)
                Marshal.FinalReleaseComObject(Sheet1)
                Sheet1 = Nothing
            Next

            xlWorkBook.Close()
            xlApp.UserControl = True
            xlApp.Quit()

        Catch ex As Exception
            Success = False
        Finally

            If Not xlWorkSheets Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkSheets)
                xlWorkSheets = Nothing
            End If

            If Not xlNames Is Nothing Then
                Marshal.FinalReleaseComObject(xlNames)
                xlNames = Nothing
            End If

            If Not xlActiveRanges Is Nothing Then
                Marshal.FinalReleaseComObject(xlActiveRanges)
                xlActiveRanges = Nothing
            End If
            If Not xlActiveRanges Is Nothing Then
                Marshal.FinalReleaseComObject(xlActiveRanges)
                xlActiveRanges = Nothing
            End If

            If Not xlWorkBook Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If Not xlWorkBooks Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If Not xlApp Is Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If
        End Try

        Return sheetNames

    End Function
    ''' <summary>
    ''' Used to return the last used row for each column within the range of ColumnCount
    ''' </summary>
    ''' <param name="FileName">Existing Excel file</param>
    ''' <param name="SheetName">Name of sheet in FileName</param>
    ''' <param name="ColumnCount">How many columns to get data for</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' In regards to ColumnCount, passing 3 would populate the Dictionary with columns A thru C etc.
    ''' </remarks>
    Public Function UsedColumns(ByVal FileName As String, ByVal SheetName As String, ByVal ColumnCount As Integer) As Dictionary(Of String, Integer)
        Dim ColumnData As New Dictionary(Of String, Integer)

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing

        xlApp = New Excel.Application
        xlApp.DisplayAlerts = False
        xlWorkBooks = xlApp.Workbooks
        xlWorkBook = xlWorkBooks.Open(FileName)

        xlApp.Visible = False

        xlWorkSheets = xlWorkBook.Sheets

        For x As Integer = 1 To xlWorkSheets.Count

            xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)


            If xlWorkSheet.Name = SheetName Then

                Dim xlCells As Excel.Range = xlWorkSheet.Cells()
                Dim xlTempRange1 As Excel.Range = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell)
                Dim xlTempRange2 = xlWorkSheet.Rows


                For Col As Integer = 1 To ColumnCount

                    Dim xlTempRange3 = xlWorkSheet.Range(Col.ExcelColumnName & xlTempRange2.Count)
                    Dim xlTempRange4 = xlTempRange3.End(Excel.XlDirection.xlUp)

                    ColumnData.Add(Col.ExcelColumnName, xlTempRange4.Row)
                    Marshal.FinalReleaseComObject(xlTempRange4)
                    xlTempRange4 = Nothing

                    Marshal.FinalReleaseComObject(xlTempRange3)
                    xlTempRange3 = Nothing
                Next

                Marshal.FinalReleaseComObject(xlTempRange2)
                xlTempRange2 = Nothing

                Marshal.FinalReleaseComObject(xlTempRange1)
                xlTempRange1 = Nothing

                Marshal.FinalReleaseComObject(xlCells)
                xlCells = Nothing

            End If

            Marshal.FinalReleaseComObject(xlWorkSheet)
            xlWorkSheet = Nothing

        Next

        xlWorkBook.Close()
        xlApp.UserControl = True
        xlApp.Quit()

        ReleaseComObject(xlWorkSheets)
        ReleaseComObject(xlWorkSheet)
        ReleaseComObject(xlWorkBook)
        ReleaseComObject(xlWorkBooks)
        ReleaseComObject(xlApp)

        Return ColumnData
    End Function
    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                Marshal.ReleaseComObject(obj)
            End If
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub
End Class
