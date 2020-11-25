<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReprintShippingLabel
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RdoProcessed = New System.Windows.Forms.RadioButton()
        Me.RdoBoth = New System.Windows.Forms.RadioButton()
        Me.RdoCompleted = New System.Windows.Forms.RadioButton()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SfDataGrid1 = New Syncfusion.WinForms.DataGrid.SfDataGrid()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.SfDateTimeEdit1 = New Syncfusion.WinForms.Input.SfDateTimeEdit()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel3.SuspendLayout()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel6.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.RdoProcessed)
        Me.Panel3.Controls.Add(Me.RdoBoth)
        Me.Panel3.Controls.Add(Me.RdoCompleted)
        Me.Panel3.Location = New System.Drawing.Point(12, 119)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(656, 59)
        Me.Panel3.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Show Only State :"
        '
        'RdoProcessed
        '
        Me.RdoProcessed.AutoSize = True
        Me.RdoProcessed.Location = New System.Drawing.Point(186, 18)
        Me.RdoProcessed.Name = "RdoProcessed"
        Me.RdoProcessed.Size = New System.Drawing.Size(94, 20)
        Me.RdoProcessed.TabIndex = 0
        Me.RdoProcessed.Text = "Processing"
        Me.RdoProcessed.UseVisualStyleBackColor = True
        '
        'RdoBoth
        '
        Me.RdoBoth.AutoSize = True
        Me.RdoBoth.Checked = True
        Me.RdoBoth.Location = New System.Drawing.Point(563, 20)
        Me.RdoBoth.Name = "RdoBoth"
        Me.RdoBoth.Size = New System.Drawing.Size(53, 20)
        Me.RdoBoth.TabIndex = 2
        Me.RdoBoth.TabStop = True
        Me.RdoBoth.Text = "Both"
        Me.RdoBoth.UseVisualStyleBackColor = True
        '
        'RdoCompleted
        '
        Me.RdoCompleted.AutoSize = True
        Me.RdoCompleted.Location = New System.Drawing.Point(368, 20)
        Me.RdoCompleted.Name = "RdoCompleted"
        Me.RdoCompleted.Size = New System.Drawing.Size(92, 20)
        Me.RdoCompleted.TabIndex = 1
        Me.RdoCompleted.Text = "Completed"
        Me.RdoCompleted.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(296, 600)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 35)
        Me.Button2.TabIndex = 25
        Me.Button2.Text = "Close"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, -13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 16)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Date :"
        '
        'SfDataGrid1
        '
        Me.SfDataGrid1.AccessibleName = "Table"
        Me.SfDataGrid1.AllowDraggingColumns = True
        Me.SfDataGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SfDataGrid1.Location = New System.Drawing.Point(15, 196)
        Me.SfDataGrid1.Name = "SfDataGrid1"
        Me.SfDataGrid1.PreviewRowHeight = 35
        Me.SfDataGrid1.Size = New System.Drawing.Size(659, 398)
        Me.SfDataGrid1.TabIndex = 20
        Me.SfDataGrid1.Text = "SfDataGrid1"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(692, 196)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 55)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "Address"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(692, 257)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(97, 59)
        Me.Button3.TabIndex = 28
        Me.Button3.Text = "Return Address"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Location = New System.Drawing.Point(692, 322)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(97, 59)
        Me.Button4.TabIndex = 29
        Me.Button4.Text = "Return Address Discreet"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.RadioButton2)
        Me.Panel6.Controls.Add(Me.Panel5)
        Me.Panel6.Controls.Add(Me.RadioButton1)
        Me.Panel6.Controls.Add(Me.Panel4)
        Me.Panel6.Location = New System.Drawing.Point(12, 11)
        Me.Panel6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(565, 103)
        Me.Panel6.TabIndex = 30
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(17, 57)
        Me.RadioButton2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(109, 20)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Name Search"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.Label28)
        Me.Panel5.Controls.Add(Me.Button8)
        Me.Panel5.Controls.Add(Me.TextBox1)
        Me.Panel5.Location = New System.Drawing.Point(173, 43)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(380, 57)
        Me.Panel5.TabIndex = 20
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(3, 15)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(94, 16)
        Me.Label28.TabIndex = 17
        Me.Label28.Text = "Search String :"
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(291, 9)
        Me.Button8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(75, 34)
        Me.Button8.TabIndex = 16
        Me.Button8.Text = "Search"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(111, 15)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(171, 22)
        Me.TextBox1.TabIndex = 18
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(17, 4)
        Me.RadioButton1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(101, 20)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Date Search"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Button5)
        Me.Panel4.Controls.Add(Me.SfDateTimeEdit1)
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Location = New System.Drawing.Point(224, 0)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(331, 42)
        Me.Panel4.TabIndex = 19
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(243, 2)
        Me.Button5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 34)
        Me.Button5.TabIndex = 11
        Me.Button5.Text = "Refresh"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'SfDateTimeEdit1
        '
        Me.SfDateTimeEdit1.Location = New System.Drawing.Point(65, 2)
        Me.SfDateTimeEdit1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SfDateTimeEdit1.Name = "SfDateTimeEdit1"
        Me.SfDateTimeEdit1.Size = New System.Drawing.Size(171, 34)
        Me.SfDateTimeEdit1.TabIndex = 10
        Me.SfDateTimeEdit1.Value = New Date(2020, 10, 19, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 2)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 16)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Date :"
        '
        'FrmReprintShippingLabel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(801, 647)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SfDataGrid1)
        Me.Name = "FrmReprintShippingLabel"
        Me.Text = "FrmReprintShippingLabel"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.SfDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents RdoProcessed As RadioButton
    Friend WithEvents RdoBoth As RadioButton
    Friend WithEvents RdoCompleted As RadioButton
    Friend WithEvents Button2 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents SfDataGrid1 As Syncfusion.WinForms.DataGrid.SfDataGrid
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Panel6 As Panel
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Label28 As Label
    Friend WithEvents Button8 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Button5 As Button
    Friend WithEvents SfDateTimeEdit1 As Syncfusion.WinForms.Input.SfDateTimeEdit
    Friend WithEvents Label2 As Label
End Class
