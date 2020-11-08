<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReprint
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.SfDateEntry = New Syncfusion.WinForms.Input.SfDateTimeEdit()
        Me.SfDataGrid1 = New Syncfusion.WinForms.DataGrid.SfDataGrid()
        Me.BtnReprint = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.RdoBoth = New System.Windows.Forms.RadioButton()
        Me.RdoCompleted = New System.Windows.Forms.RadioButton()
        Me.RdoProcessed = New System.Windows.Forms.RadioButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 17)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Date :"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(266, 10)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(75, 30)
        Me.Button7.TabIndex = 15
        Me.Button7.Text = "Refresh"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'SfDateEntry
        '
        Me.SfDateEntry.Location = New System.Drawing.Point(79, 10)
        Me.SfDateEntry.Name = "SfDateEntry"
        Me.SfDateEntry.Size = New System.Drawing.Size(171, 34)
        Me.SfDateEntry.TabIndex = 14
        Me.SfDateEntry.Value = New Date(2020, 10, 19, 0, 0, 0, 0)
        '
        'SfDataGrid1
        '
        Me.SfDataGrid1.AccessibleName = "Table"
        Me.SfDataGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SfDataGrid1.Location = New System.Drawing.Point(12, 123)
        Me.SfDataGrid1.Name = "SfDataGrid1"
        Me.SfDataGrid1.PreviewRowHeight = 35
        Me.SfDataGrid1.Size = New System.Drawing.Size(776, 300)
        Me.SfDataGrid1.TabIndex = 13
        Me.SfDataGrid1.Text = "SfDataGrid1"
        '
        'BtnReprint
        '
        Me.BtnReprint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnReprint.Location = New System.Drawing.Point(308, 451)
        Me.BtnReprint.Name = "BtnReprint"
        Me.BtnReprint.Size = New System.Drawing.Size(75, 35)
        Me.BtnReprint.TabIndex = 17
        Me.BtnReprint.Text = "RePrint"
        Me.BtnReprint.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(389, 451)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 35)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'RdoBoth
        '
        Me.RdoBoth.AutoSize = True
        Me.RdoBoth.Checked = True
        Me.RdoBoth.Location = New System.Drawing.Point(664, 16)
        Me.RdoBoth.Name = "RdoBoth"
        Me.RdoBoth.Size = New System.Drawing.Size(58, 21)
        Me.RdoBoth.TabIndex = 2
        Me.RdoBoth.TabStop = True
        Me.RdoBoth.Text = "Both"
        Me.RdoBoth.UseVisualStyleBackColor = True
        '
        'RdoCompleted
        '
        Me.RdoCompleted.AutoSize = True
        Me.RdoCompleted.Location = New System.Drawing.Point(468, 16)
        Me.RdoCompleted.Name = "RdoCompleted"
        Me.RdoCompleted.Size = New System.Drawing.Size(96, 21)
        Me.RdoCompleted.TabIndex = 1
        Me.RdoCompleted.Text = "Completed"
        Me.RdoCompleted.UseVisualStyleBackColor = True
        '
        'RdoProcessed
        '
        Me.RdoProcessed.AutoSize = True
        Me.RdoProcessed.Location = New System.Drawing.Point(272, 16)
        Me.RdoProcessed.Name = "RdoProcessed"
        Me.RdoProcessed.Size = New System.Drawing.Size(96, 21)
        Me.RdoProcessed.TabIndex = 0
        Me.RdoProcessed.Text = "Processed"
        Me.RdoProcessed.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.RdoProcessed)
        Me.Panel1.Controls.Add(Me.RdoBoth)
        Me.Panel1.Controls.Add(Me.RdoCompleted)
        Me.Panel1.Location = New System.Drawing.Point(15, 58)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(773, 59)
        Me.Panel1.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Show Only State :"
        '
        'FrmReprint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 498)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.BtnReprint)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.SfDateEntry)
        Me.Controls.Add(Me.SfDataGrid1)
        Me.Name = "FrmReprint"
        Me.Text = "Reprint Certificate"
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label3 As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents SfDateEntry As Syncfusion.WinForms.Input.SfDateTimeEdit
    Friend WithEvents SfDataGrid1 As Syncfusion.WinForms.DataGrid.SfDataGrid
    Friend WithEvents BtnReprint As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents RdoBoth As RadioButton
    Friend WithEvents RdoCompleted As RadioButton
    Friend WithEvents RdoProcessed As RadioButton
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label1 As Label
End Class
