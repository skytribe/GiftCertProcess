Public Class FrmJumprunItemsSelect
    Property CurrentSelectedItem As Integer = -1
    Property CurrentPrice As Double

    Private Sub FrmJumprunItemsSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'PopulateList with Items from Jumprun
        SfDataGrid1.DataSource = RetrieveJumpRunItems()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.CurrentPrice = 0
        Me.CurrentSelectedItem = -1
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '//Set the properties for return
        If SfDataGrid1.SelectedItem IsNot Nothing Then
            CurrentSelectedItem = CType(SfDataGrid1.SelectedItem, ClsJumpRunItem).ID
            CurrentPrice = CType(SfDataGrid1.SelectedItem, ClsJumpRunItem).Price
        End If
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Filter List
        If String.IsNullOrEmpty(TextBox1.Text.Trim) = False Then
            SfDataGrid1.DataSource = RetrieveJumpRunItems(TextBox1.Text.Trim)
        Else
            SfDataGrid1.DataSource = RetrieveJumpRunItems()
        End If
    End Sub
End Class