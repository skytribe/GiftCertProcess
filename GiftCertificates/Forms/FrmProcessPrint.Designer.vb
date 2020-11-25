<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmProcessPrint
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RdoInPerson = New System.Windows.Forms.RadioButton()
        Me.RdoPrintDiscrete = New System.Windows.Forms.RadioButton()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.RdoEmail = New System.Windows.Forms.RadioButton()
        Me.RdoPrint = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ChkUpdateStatus = New System.Windows.Forms.CheckBox()
        Me.ChkReturnLabel = New System.Windows.Forms.CheckBox()
        Me.lblRecipientEmail = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblRecipientEmail)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.RdoInPerson)
        Me.GroupBox1.Controls.Add(Me.RdoPrintDiscrete)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.RdoEmail)
        Me.GroupBox1.Controls.Add(Me.RdoPrint)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 11)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(679, 124)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Output Option"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(227, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Purchaser Email :"
        '
        'RdoInPerson
        '
        Me.RdoInPerson.AutoSize = True
        Me.RdoInPerson.Location = New System.Drawing.Point(20, 26)
        Me.RdoInPerson.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoInPerson.Name = "RdoInPerson"
        Me.RdoInPerson.Size = New System.Drawing.Size(82, 20)
        Me.RdoInPerson.TabIndex = 4
        Me.RdoInPerson.TabStop = True
        Me.RdoInPerson.Text = "In Person"
        Me.RdoInPerson.UseVisualStyleBackColor = True
        '
        'RdoPrintDiscrete
        '
        Me.RdoPrintDiscrete.AutoSize = True
        Me.RdoPrintDiscrete.Location = New System.Drawing.Point(211, 26)
        Me.RdoPrintDiscrete.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoPrintDiscrete.Name = "RdoPrintDiscrete"
        Me.RdoPrintDiscrete.Size = New System.Drawing.Size(105, 20)
        Me.RdoPrintDiscrete.TabIndex = 3
        Me.RdoPrintDiscrete.TabStop = True
        Me.RdoPrintDiscrete.Text = "Print Discrete"
        Me.RdoPrintDiscrete.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(348, 66)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(321, 22)
        Me.TextBox1.TabIndex = 2
        '
        'RdoEmail
        '
        Me.RdoEmail.AutoSize = True
        Me.RdoEmail.Location = New System.Drawing.Point(348, 26)
        Me.RdoEmail.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoEmail.Name = "RdoEmail"
        Me.RdoEmail.Size = New System.Drawing.Size(60, 20)
        Me.RdoEmail.TabIndex = 1
        Me.RdoEmail.TabStop = True
        Me.RdoEmail.Text = "Email"
        Me.RdoEmail.UseVisualStyleBackColor = True
        '
        'RdoPrint
        '
        Me.RdoPrint.AutoSize = True
        Me.RdoPrint.Location = New System.Drawing.Point(129, 26)
        Me.RdoPrint.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RdoPrint.Name = "RdoPrint"
        Me.RdoPrint.Size = New System.Drawing.Size(52, 20)
        Me.RdoPrint.TabIndex = 0
        Me.RdoPrint.TabStop = True
        Me.RdoPrint.Text = "Print"
        Me.RdoPrint.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(317, 167)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 39)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(399, 167)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 39)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(721, 37)
        Me.Button3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(141, 62)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Print Shipping Label"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ChkUpdateStatus
        '
        Me.ChkUpdateStatus.AutoSize = True
        Me.ChkUpdateStatus.Location = New System.Drawing.Point(12, 140)
        Me.ChkUpdateStatus.Name = "ChkUpdateStatus"
        Me.ChkUpdateStatus.Size = New System.Drawing.Size(112, 20)
        Me.ChkUpdateStatus.TabIndex = 4
        Me.ChkUpdateStatus.Text = "Update Status"
        Me.ChkUpdateStatus.UseVisualStyleBackColor = True
        '
        'ChkReturnLabel
        '
        Me.ChkReturnLabel.AutoSize = True
        Me.ChkReturnLabel.Location = New System.Drawing.Point(721, 115)
        Me.ChkReturnLabel.Name = "ChkReturnLabel"
        Me.ChkReturnLabel.Size = New System.Drawing.Size(103, 20)
        Me.ChkReturnLabel.TabIndex = 5
        Me.ChkReturnLabel.Text = "Return Label"
        Me.ChkReturnLabel.UseVisualStyleBackColor = True
        '
        'lblRecipientEmail
        '
        Me.lblRecipientEmail.AutoSize = True
        Me.lblRecipientEmail.Location = New System.Drawing.Point(359, 104)
        Me.lblRecipientEmail.Name = "lblRecipientEmail"
        Me.lblRecipientEmail.Size = New System.Drawing.Size(0, 16)
        Me.lblRecipientEmail.TabIndex = 8
        '
        'FrmProcessPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(885, 225)
        Me.Controls.Add(Me.ChkReturnLabel)
        Me.Controls.Add(Me.ChkUpdateStatus)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "FrmProcessPrint"
        Me.Text = "Delivery"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents RdoEmail As RadioButton
    Friend WithEvents RdoPrint As RadioButton
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents RdoPrintDiscrete As RadioButton
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents RdoInPerson As RadioButton
    Friend WithEvents Label3 As Label
    Friend WithEvents ChkUpdateStatus As CheckBox
    Friend WithEvents ChkReturnLabel As CheckBox
    Friend WithEvents lblRecipientEmail As Label
End Class
