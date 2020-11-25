<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReport
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
        Me.SfDateTimeEdit1 = New Syncfusion.WinForms.Input.SfDateTimeEdit()
        Me.SfDateTimeEdit2 = New Syncfusion.WinForms.Input.SfDateTimeEdit()
        Me.BtnGo = New System.Windows.Forms.Button()
        Me.BtnExport = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SfDataGrid1 = New Syncfusion.WinForms.DataGrid.SfDataGrid()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SfDateTimeEdit1
        '
        Me.SfDateTimeEdit1.Location = New System.Drawing.Point(83, 23)
        Me.SfDateTimeEdit1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SfDateTimeEdit1.Name = "SfDateTimeEdit1"
        Me.SfDateTimeEdit1.Size = New System.Drawing.Size(232, 34)
        Me.SfDateTimeEdit1.TabIndex = 0
        '
        'SfDateTimeEdit2
        '
        Me.SfDateTimeEdit2.Location = New System.Drawing.Point(83, 63)
        Me.SfDateTimeEdit2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SfDateTimeEdit2.Name = "SfDateTimeEdit2"
        Me.SfDateTimeEdit2.Size = New System.Drawing.Size(232, 34)
        Me.SfDateTimeEdit2.TabIndex = 1
        '
        'BtnGo
        '
        Me.BtnGo.Location = New System.Drawing.Point(331, 23)
        Me.BtnGo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnGo.Name = "BtnGo"
        Me.BtnGo.Size = New System.Drawing.Size(75, 74)
        Me.BtnGo.TabIndex = 2
        Me.BtnGo.Text = "Go"
        Me.BtnGo.UseVisualStyleBackColor = True
        '
        'BtnExport
        '
        Me.BtnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnExport.Location = New System.Drawing.Point(704, 116)
        Me.BtnExport.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnExport.Name = "BtnExport"
        Me.BtnExport.Size = New System.Drawing.Size(75, 23)
        Me.BtnExport.TabIndex = 4
        Me.BtnExport.Text = "Export"
        Me.BtnExport.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "From :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 63)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "To :"
        '
        'SfDataGrid1
        '
        Me.SfDataGrid1.AccessibleName = "Table"
        Me.SfDataGrid1.AllowDraggingColumns = True
        Me.SfDataGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SfDataGrid1.Location = New System.Drawing.Point(24, 116)
        Me.SfDataGrid1.Name = "SfDataGrid1"
        Me.SfDataGrid1.PreviewRowHeight = 35
        Me.SfDataGrid1.Size = New System.Drawing.Size(659, 312)
        Me.SfDataGrid1.TabIndex = 21
        Me.SfDataGrid1.Text = "SfDataGrid1"
        '
        'FrmReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SfDataGrid1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnExport)
        Me.Controls.Add(Me.BtnGo)
        Me.Controls.Add(Me.SfDateTimeEdit2)
        Me.Controls.Add(Me.SfDateTimeEdit1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "FrmReport"
        Me.Text = "GC Order Report"
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SfDateTimeEdit1 As Syncfusion.WinForms.Input.SfDateTimeEdit
    Friend WithEvents SfDateTimeEdit2 As Syncfusion.WinForms.Input.SfDateTimeEdit
    Friend WithEvents BtnGo As Button
    Friend WithEvents BtnExport As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents SfDataGrid1 As Syncfusion.WinForms.DataGrid.SfDataGrid
End Class
