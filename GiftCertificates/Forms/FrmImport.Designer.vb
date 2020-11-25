<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmImport
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtFilename = New System.Windows.Forms.TextBox()
        Me.BtnBrowse = New System.Windows.Forms.Button()
        Me.BtnImport = New System.Windows.Forms.Button()
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Filename :"
        '
        'TxtFilename
        '
        Me.TxtFilename.Location = New System.Drawing.Point(103, 25)
        Me.TxtFilename.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtFilename.Name = "TxtFilename"
        Me.TxtFilename.Size = New System.Drawing.Size(314, 20)
        Me.TxtFilename.TabIndex = 1
        '
        'BtnBrowse
        '
        Me.BtnBrowse.Location = New System.Drawing.Point(421, 24)
        Me.BtnBrowse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnBrowse.Name = "BtnBrowse"
        Me.BtnBrowse.Size = New System.Drawing.Size(56, 19)
        Me.BtnBrowse.TabIndex = 2
        Me.BtnBrowse.Text = "..."
        Me.BtnBrowse.UseVisualStyleBackColor = True
        '
        'BtnImport
        '
        Me.BtnImport.Location = New System.Drawing.Point(151, 69)
        Me.BtnImport.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnImport.Name = "BtnImport"
        Me.BtnImport.Size = New System.Drawing.Size(56, 34)
        Me.BtnImport.TabIndex = 3
        Me.BtnImport.Text = "Import"
        Me.BtnImport.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(220, 69)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(56, 34)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.Text = "Cancel"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 121)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.BtnImport)
        Me.Controls.Add(Me.BtnBrowse)
        Me.Controls.Add(Me.TxtFilename)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "FrmImport"
        Me.Text = "Import From Shopify"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents TxtFilename As TextBox
    Friend WithEvents BtnBrowse As Button
    Friend WithEvents BtnImport As Button
    Friend WithEvents BtnClose As Button
End Class
