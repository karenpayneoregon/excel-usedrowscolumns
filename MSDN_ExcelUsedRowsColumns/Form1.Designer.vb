<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.cmdAddress = New System.Windows.Forms.Button()
        Me.cmdGood = New System.Windows.Forms.Button()
        Me.cmdAddress1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(595, 168)
        Me.DataGridView1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdAddress1)
        Me.Panel1.Controls.Add(Me.ListBox1)
        Me.Panel1.Controls.Add(Me.cmdAddress)
        Me.Panel1.Controls.Add(Me.cmdGood)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 168)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(595, 128)
        Me.Panel1.TabIndex = 1
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(190, 14)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(192, 95)
        Me.ListBox1.TabIndex = 4
        '
        'cmdAddress
        '
        Me.cmdAddress.Location = New System.Drawing.Point(387, 14)
        Me.cmdAddress.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdAddress.Name = "cmdAddress"
        Me.cmdAddress.Size = New System.Drawing.Size(74, 26)
        Me.cmdAddress.TabIndex = 3
        Me.cmdAddress.Text = "Address"
        Me.cmdAddress.UseVisualStyleBackColor = True
        '
        'cmdGood
        '
        Me.cmdGood.Location = New System.Drawing.Point(11, 13)
        Me.cmdGood.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdGood.Name = "cmdGood"
        Me.cmdGood.Size = New System.Drawing.Size(62, 26)
        Me.cmdGood.TabIndex = 1
        Me.cmdGood.Text = "Good"
        Me.cmdGood.UseVisualStyleBackColor = True
        '
        'cmdAddress1
        '
        Me.cmdAddress1.Location = New System.Drawing.Point(387, 44)
        Me.cmdAddress1.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdAddress1.Name = "cmdAddress1"
        Me.cmdAddress1.Size = New System.Drawing.Size(74, 26)
        Me.cmdAddress1.TabIndex = 5
        Me.cmdAddress1.Text = "Address1"
        Me.cmdAddress1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(595, 296)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdGood As System.Windows.Forms.Button
    Friend WithEvents cmdAddress As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents cmdAddress1 As Button
End Class
