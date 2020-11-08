Public Class FrmDevelopmentReset
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim lids = GetListOfGCCustomers()
        '//Remove selected Items
        For Each i In lids

            '    Dim id = i.Row(0)
            '    DeleteCustomer(id)
            MessageBox.Show(i)
        Next

        MsgBox("To Implement - To delete generate JR Customers > 3xxxxx")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DeleteImportedGiftCertificates()
        MsgBox("Completed")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("To Implement")
    End Sub
End Class