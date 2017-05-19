Option Strict On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Public Class ExcelInformation
    Public Function GetUsed(ByVal FileName As String, ByVal Sheets As List(Of String)) As List(Of ExcelInfo)
        Dim Results As New List(Of ExcelInfo)

        Dim RowsUsed As Integer = -1
        Dim ColsUsed As Integer = -1

        If IO.File.Exists(FileName) Then
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

                For Each SheetName In Sheets

                    If xlWorkSheet.Name = SheetName Then

                        Dim xlCells As Excel.Range = xlWorkSheet.Cells
                        Dim xlTempRange As Excel.Range = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell)

                        RowsUsed = xlTempRange.Row
                        ColsUsed = xlTempRange.Column

                        Results.Add(New ExcelInfo With {.FileName = FileName, .SheetName = SheetName, .UsedRows = RowsUsed, .UsedColumns = ColsUsed, .LastCell = $"{ColsUsed.ExcelColumnName}:{RowsUsed}"})

                        Marshal.FinalReleaseComObject(xlTempRange)
                        xlTempRange = Nothing

                        Marshal.FinalReleaseComObject(xlCells)
                        xlCells = Nothing

                    End If
                Next
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

            Return Results

        Else
            Throw New Exception("'" & FileName & "' not found.")
        End If

        Return Results

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

