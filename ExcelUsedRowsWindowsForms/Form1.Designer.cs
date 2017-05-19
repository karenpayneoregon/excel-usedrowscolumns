namespace ExcelUsedRowsWindowsForms
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.DataGridView1 = new System.Windows.Forms.DataGridView();
            this.Panel1 = new System.Windows.Forms.Panel();
            this.cmdAddress1 = new System.Windows.Forms.Button();
            this.ListBox1 = new System.Windows.Forms.ListBox();
            this.cmdAddress = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView1)).BeginInit();
            this.Panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataGridView1
            // 
            this.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DataGridView1.Location = new System.Drawing.Point(0, 0);
            this.DataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.DataGridView1.Name = "DataGridView1";
            this.DataGridView1.RowTemplate.Height = 24;
            this.DataGridView1.Size = new System.Drawing.Size(595, 168);
            this.DataGridView1.TabIndex = 2;
            // 
            // Panel1
            // 
            this.Panel1.Controls.Add(this.cmdAddress1);
            this.Panel1.Controls.Add(this.ListBox1);
            this.Panel1.Controls.Add(this.cmdAddress);
            this.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Panel1.Location = new System.Drawing.Point(0, 168);
            this.Panel1.Margin = new System.Windows.Forms.Padding(2);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(595, 128);
            this.Panel1.TabIndex = 3;
            // 
            // cmdAddress1
            // 
            this.cmdAddress1.Location = new System.Drawing.Point(387, 44);
            this.cmdAddress1.Margin = new System.Windows.Forms.Padding(2);
            this.cmdAddress1.Name = "cmdAddress1";
            this.cmdAddress1.Size = new System.Drawing.Size(74, 26);
            this.cmdAddress1.TabIndex = 5;
            this.cmdAddress1.Text = "Address1";
            this.cmdAddress1.UseVisualStyleBackColor = true;
            this.cmdAddress1.Click += new System.EventHandler(this.cmdAddress1_Click);
            // 
            // ListBox1
            // 
            this.ListBox1.FormattingEnabled = true;
            this.ListBox1.Location = new System.Drawing.Point(190, 14);
            this.ListBox1.Name = "ListBox1";
            this.ListBox1.Size = new System.Drawing.Size(192, 95);
            this.ListBox1.TabIndex = 4;
            // 
            // cmdAddress
            // 
            this.cmdAddress.Location = new System.Drawing.Point(387, 14);
            this.cmdAddress.Margin = new System.Windows.Forms.Padding(2);
            this.cmdAddress.Name = "cmdAddress";
            this.cmdAddress.Size = new System.Drawing.Size(74, 26);
            this.cmdAddress.TabIndex = 3;
            this.cmdAddress.Text = "Address";
            this.cmdAddress.UseVisualStyleBackColor = true;
            this.cmdAddress.Click += new System.EventHandler(this.cmdAddress_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(595, 296);
            this.Controls.Add(this.DataGridView1);
            this.Controls.Add(this.Panel1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DataGridView1)).EndInit();
            this.Panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.DataGridView DataGridView1;
        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.Button cmdAddress1;
        internal System.Windows.Forms.ListBox ListBox1;
        internal System.Windows.Forms.Button cmdAddress;
    }
}

