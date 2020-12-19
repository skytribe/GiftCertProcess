Public Class FrmNewPromo
    Public Property PromoName As String = ""
    Private Sub FrmNewPromo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.PromoName = TextBox1.Text.Trim
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.PromoName = ""
        Me.Close()
    End Sub
End Class