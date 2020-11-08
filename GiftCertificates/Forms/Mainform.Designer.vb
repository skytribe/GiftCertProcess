<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Mainform
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.Color.Magenta
        Me.Button1.Location = New System.Drawing.Point(63, 35)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(257, 54)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Import From Shopify"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.Color.Magenta
        Me.Button2.Location = New System.Drawing.Point(343, 35)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(288, 54)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Manual Entry"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.Color.Magenta
        Me.Button3.Location = New System.Drawing.Point(63, 176)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(257, 52)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Process Certificates"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(63, 245)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(257, 52)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Close"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.ForeColor = System.Drawing.Color.Red
        Me.Button5.Location = New System.Drawing.Point(63, 105)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(257, 54)
        Me.Button5.TabIndex = 4
        Me.Button5.Text = "Modify Entry (Before Processing)"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.ForeColor = System.Drawing.Color.Magenta
        Me.Button6.Location = New System.Drawing.Point(343, 176)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(257, 52)
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "Reprint Certificate"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.ForeColor = System.Drawing.Color.Magenta
        Me.Button7.Location = New System.Drawing.Point(343, 245)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(257, 52)
        Me.Button7.TabIndex = 6
        Me.Button7.Text = "Development Reset/Test"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Mainform
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(642, 312)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Mainform"
        Me.Text = "Gift Certificates"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
End Class
