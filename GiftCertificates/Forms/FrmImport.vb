Public Class FrmImport
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            If System.IO.File.Exists(TxtFilename.Text) Then
                ImportShopifyCSVFile(TxtFilename.Text)
                Me.Close()
            Else
                MessageBox.Show("The selected file does not exist.", "Import", MessageBoxButtons.OK)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim dlgOpenFile As New OpenFileDialog
            dlgOpenFile.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*"
            dlgOpenFile.FilterIndex = 0
            dlgOpenFile.Title = "Locate Shopify CSV file"
            Dim retValue = dlgOpenFile.ShowDialog()

            If retValue <> Windows.Forms.DialogResult.Cancel Then
                TxtFilename.Text = dlgOpenFile.FileName
            End If
        Catch ex As Exception

        End Try
    End Sub

End Class