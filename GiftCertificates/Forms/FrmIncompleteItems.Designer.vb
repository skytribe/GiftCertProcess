<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmIncompleteItems
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
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.RdoBoth = New System.Windows.Forms.RadioButton()
        Me.RdoProcessed = New System.Windows.Forms.RadioButton()
        Me.RdoEntered = New System.Windows.Forms.RadioButton()
        Me.SfDataGrid1 = New Syncfusion.WinForms.DataGrid.SfDataGrid()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3.SuspendLayout()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.RdoBoth)
        Me.Panel3.Controls.Add(Me.RdoProcessed)
        Me.Panel3.Controls.Add(Me.RdoEntered)
        Me.Panel3.Location = New System.Drawing.Point(131, 11)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(476, 47)
        Me.Panel3.TabIndex = 17
        '
        'RdoBoth
        '
        Me.RdoBoth.AutoSize = True
        Me.RdoBoth.Location = New System.Drawing.Point(270, 14)
        Me.RdoBoth.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoBoth.Name = "RdoBoth"
        Me.RdoBoth.Size = New System.Drawing.Size(53, 20)
        Me.RdoBoth.TabIndex = 2
        Me.RdoBoth.Text = "Both"
        Me.RdoBoth.UseVisualStyleBackColor = True
        '
        'RdoProcessed
        '
        Me.RdoProcessed.AutoSize = True
        Me.RdoProcessed.Location = New System.Drawing.Point(140, 14)
        Me.RdoProcessed.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoProcessed.Name = "RdoProcessed"
        Me.RdoProcessed.Size = New System.Drawing.Size(94, 20)
        Me.RdoProcessed.TabIndex = 1
        Me.RdoProcessed.Text = "Processing"
        Me.RdoProcessed.UseVisualStyleBackColor = True
        '
        'RdoEntered
        '
        Me.RdoEntered.AutoSize = True
        Me.RdoEntered.Checked = True
        Me.RdoEntered.Location = New System.Drawing.Point(17, 14)
        Me.RdoEntered.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoEntered.Name = "RdoEntered"
        Me.RdoEntered.Size = New System.Drawing.Size(73, 20)
        Me.RdoEntered.TabIndex = 0
        Me.RdoEntered.TabStop = True
        Me.RdoEntered.Text = "Entered"
        Me.RdoEntered.UseVisualStyleBackColor = True
        '
        'SfDataGrid1
        '
        Me.SfDataGrid1.AccessibleName = "Table"
        Me.SfDataGrid1.Location = New System.Drawing.Point(12, 78)
        Me.SfDataGrid1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SfDataGrid1.Name = "SfDataGrid1"
        Me.SfDataGrid1.PreviewRowHeight = 35
        Me.SfDataGrid1.Size = New System.Drawing.Size(776, 273)
        Me.SfDataGrid1.TabIndex = 16
        Me.SfDataGrid1.Text = "SfDataGrid1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(171, 380)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Process"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(271, 380)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 19
        Me.Button2.Text = "Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Status:"
        '
        'FrmIncompleteItems
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.SfDataGrid1)
        Me.Name = "FrmIncompleteItems"
        Me.Text = "Incomplete Orders"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel3 As Panel
    Friend WithEvents RdoProcessed As RadioButton
    Friend WithEvents RdoEntered As RadioButton
    Friend WithEvents SfDataGrid1 As Syncfusion.WinForms.DataGrid.SfDataGrid
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents RdoBoth As RadioButton
End Class
