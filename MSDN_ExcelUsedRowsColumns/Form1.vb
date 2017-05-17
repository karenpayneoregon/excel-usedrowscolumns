Imports ExcelUsedColumnsLib

Public Class Form1
    Private FileName As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "W2.xlsx")
    Private SheetNames As New List(Of String)
    Private ExcelInformationData As List(Of ExcelInfo)
    Public Sub New()
        InitializeComponent()
        Dim ops As New GetExcelColumnLastRowInformation
        SheetNames = ops.GetSheets(FileName)
        Dim info = New UsedInformation
        ExcelInformationData = info.UsedInformation(FileName, SheetNames)
    End Sub
    Private Sub cmdAddress1_Click(sender As Object, e As EventArgs) Handles cmdAddress1.Click
        Dim ops As New GetExcelColumnLastRowInformation
        Dim info = New UsedInformation
        ExcelInformationData = info.UsedInformation(FileName, ops.GetSheets(FileName))

        Dim SheetName As String = ExcelInformationData.FirstOrDefault.SheetName

        Dim cellAddress = (
        From item In ExcelInformationData
        Where item.SheetName = ExcelInformationData.FirstOrDefault.SheetName
        Select item.LastCell).FirstOrDefault

        MessageBox.Show($"{SheetName} - {cellAddress}")

    End Sub
    Private Sub Good_Click(sender As Object, e As EventArgs) Handles cmdGood.Click
        Dim ops As New UsedInformation
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = ops.UsedInformation(FileName, SheetNames)
        Fixer()
    End Sub
    Private Sub Fixer()
        DataGridView1.Columns("FileName").Visible = False
        DataGridView1.Columns("UsedRows").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns("UsedColumns").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub
    Private Sub cmdAddress_Click(sender As Object, e As EventArgs) Handles cmdAddress.Click
        Dim cellAddress =
        (
            From item In ExcelInformationData
            Where item.SheetName = ListBox1.Text
            Select item.LastCell).FirstOrDefault

        If cellAddress IsNot Nothing Then
            MessageBox.Show($"{ListBox1.Text} {cellAddress}")
        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.DisplayMember = "SheetName"
        ListBox1.DataSource = ExcelInformationData
    End Sub


End Class
