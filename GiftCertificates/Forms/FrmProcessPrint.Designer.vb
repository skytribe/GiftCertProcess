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
        Me.lblRecipientEmail = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RdoInPerson = New System.Windows.Forms.RadioButton()
        Me.RdoPrintDiscrete = New System.Windows.Forms.RadioButton()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.RdoEmail = New System.Windows.Forms.RadioButton()
        Me.RdoPrint = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.lblRecipientEmail)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.RdoInPerson)
        Me.GroupBox1.Controls.Add(Me.RdoPrintDiscrete)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.RdoEmail)
        Me.GroupBox1.Controls.Add(Me.RdoPrint)
        Me.GroupBox1.Location = New System.Drawing.Point(22, 26)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(679, 178)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Output Option"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(202, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Purchaser Email"
        '
        'lblRecipientEmail
        '
        Me.lblRecipientEmail.AutoSize = True
        Me.lblRecipientEmail.Location = New System.Drawing.Point(345, 133)
        Me.lblRecipientEmail.Name = "lblRecipientEmail"
        Me.lblRecipientEmail.Size = New System.Drawing.Size(12, 17)
        Me.lblRecipientEmail.TabIndex = 6
        Me.lblRecipientEmail.Text = " "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(202, 133)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Recipient Email"
        '
        'RdoInPerson
        '
        Me.RdoInPerson.AutoSize = True
        Me.RdoInPerson.Location = New System.Drawing.Point(20, 53)
        Me.RdoInPerson.Name = "RdoInPerson"
        Me.RdoInPerson.Size = New System.Drawing.Size(89, 21)
        Me.RdoInPerson.TabIndex = 4
        Me.RdoInPerson.TabStop = True
        Me.RdoInPerson.Text = "In Person"
        Me.RdoInPerson.UseVisualStyleBackColor = True
        '
        'RdoPrintDiscrete
        '
        Me.RdoPrintDiscrete.AutoSize = True
        Me.RdoPrintDiscrete.Location = New System.Drawing.Point(211, 53)
        Me.RdoPrintDiscrete.Name = "RdoPrintDiscrete"
        Me.RdoPrintDiscrete.Size = New System.Drawing.Size(114, 21)
        Me.RdoPrintDiscrete.TabIndex = 3
        Me.RdoPrintDiscrete.TabStop = True
        Me.RdoPrintDiscrete.Text = "Print Discrete"
        Me.RdoPrintDiscrete.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(348, 93)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(322, 22)
        Me.TextBox1.TabIndex = 2
        '
        'RdoEmail
        '
        Me.RdoEmail.AutoSize = True
        Me.RdoEmail.Location = New System.Drawing.Point(348, 53)
        Me.RdoEmail.Name = "RdoEmail"
        Me.RdoEmail.Size = New System.Drawing.Size(63, 21)
        Me.RdoEmail.TabIndex = 1
        Me.RdoEmail.TabStop = True
        Me.RdoEmail.Text = "Email"
        Me.RdoEmail.UseVisualStyleBackColor = True
        '
        'RdoPrint
        '
        Me.RdoPrint.AutoSize = True
        Me.RdoPrint.Location = New System.Drawing.Point(129, 53)
        Me.RdoPrint.Name = "RdoPrint"
        Me.RdoPrint.Size = New System.Drawing.Size(58, 21)
        Me.RdoPrint.TabIndex = 0
        Me.RdoPrint.TabStop = True
        Me.RdoPrint.Text = "Print"
        Me.RdoPrint.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(289, 228)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 40)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(370, 228)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 40)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(707, 33)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(142, 61)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Print Label (Purchaser)"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'FrmProcessPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(889, 282)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "FrmProcessPrint"
        Me.Text = "Delivery"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

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
    Friend WithEvents lblRecipientEmail As Label
    Friend WithEvents Label1 As Label
End Class
