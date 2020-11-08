Public Class FrmModify
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("Search for all gift certificates with a last name match")
        Dim searchLastName As String = TxtSearchLastName.Text.Trim
        '[dbo].[GCFindCertificatesWithLastName]
        Dim list = SearchGiftCertificates(searchLastName)
        SfDataGrid1.DataSource = list
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("Select the gift certificate item from the list")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Modify the gift certificate to correct any name changes.  This can be done to correct spelling mistakes and reprint")
        MsgBox("It does not change the customer records - so this needs to be used carefully")
    End Sub

    Private Sub FrmModify_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupForm
    End Sub

    Private Sub SetupForm()

    End Sub
End Class