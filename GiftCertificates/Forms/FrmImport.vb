Imports System.Reflection

Public Class FrmImport
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnImport.Click
        Try

            If System.IO.File.Exists(TxtFilename.Text) Then
                'ImportShopifyCSVFile(TxtFilename.Text)
                ImportWooCSVFile(TxtFilename.Text)
                Me.Close()
            Else
                MessageBox.Show("The selected file does not exist.", "Import", MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BtnBrowse.Click
        Try
            Dim dlgOpenFile As New OpenFileDialog
            dlgOpenFile.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*"
            dlgOpenFile.FilterIndex = 0
            dlgOpenFile.Title = "Locate Import CSV file"
            Dim retValue = dlgOpenFile.ShowDialog()

            If retValue <> Windows.Forms.DialogResult.Cancel Then
                TxtFilename.Text = dlgOpenFile.FileName
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub FrmImport_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetupForm()
    End Sub

    Private Sub SetupForm()
        Me.Text = "Import CSV File"
    End Sub
End Class