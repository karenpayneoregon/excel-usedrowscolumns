using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadSheetLightLibrary;
using System.IO;

namespace SpreadSheetLightDemo
{
    public partial class Form1 : Form
    {
        string FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "W2.xlsx");
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            var ops = new Operations();
            if (ops.GetInformation(FileName))
            {
                foreach (var item in ops.UsedRowsColumns)
                {
                    dataGridView1.Rows.Add(new object[] {item.Key,item.Value });
                }
            }
        }
    }
}
