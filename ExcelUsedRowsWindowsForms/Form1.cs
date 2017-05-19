using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelUsedColumnsLib;
using System.IO;

namespace ExcelUsedRowsWindowsForms
{
    public partial class Form1 : Form
    {
        string FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "W2.xlsx");
        List<string> SheetNames = new List<string>();
        List<ExcelInfo> ExcelInformationData;
        ExcelInformation infoExcel = new ExcelInformation();

        public Form1()
        {
            InitializeComponent();
            GetExcelColumnLastRowInformation ops = new GetExcelColumnLastRowInformation();
            SheetNames = ops.GetSheets(FileName);
            infoExcel = new ExcelInformation();
            ExcelInformationData = infoExcel.GetUsed(FileName, SheetNames);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ListBox1.DisplayMember = "SheetName";
            ListBox1.DataSource = ExcelInformationData;
            DataGridView1.DataSource = infoExcel.GetUsed(FileName, SheetNames);
            Fixer();
        }
        private void Fixer()
        {
            DataGridView1.Columns["FileName"].Visible = false;
            DataGridView1.Columns["UsedRows"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView1.Columns["UsedColumns"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void cmdAddress_Click(object sender, EventArgs e)
        {
            var cellAddress = (from item in ExcelInformationData where item.SheetName == ListBox1.Text  select item.LastCell).FirstOrDefault();
            if (cellAddress != null)
            {
                MessageBox.Show($"{ListBox1.Text} {cellAddress}");
            }
        }

        private void cmdAddress1_Click(object sender, EventArgs e)
        {
            string SheetName = ExcelInformationData.FirstOrDefault().SheetName;
            var cellAddress = (from item in ExcelInformationData where item.SheetName == ExcelInformationData.FirstOrDefault().SheetName select item.LastCell)
                .FirstOrDefault();

            MessageBox.Show($"{SheetName} - {cellAddress}");

        }
    }
}
