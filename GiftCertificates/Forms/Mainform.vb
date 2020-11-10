Public Class Mainform
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim ImportOnlineTransactionsForm As New FrmImport
            ImportOnlineTransactionsForm.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim ManualEntryForm As New FrmEntry
            ManualEntryForm.ShowDialog()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim ProcessGiftCertifcateQueueForm As New FrmProcess
            ProcessGiftCertifcateQueueForm.ShowDialog()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim x As New FrmModify
        x.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim ReportCertificateForm As New FrmReprint
            ReportCertificateForm.ShowDialog()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim x As New FrmDevelopmentReset
        x.ShowDialog()
    End Sub

    Private Sub Mainform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setupForm

    End Sub

    Private Sub setupForm()
        Me.Text = "Gift Certificates Main Form"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim sb As New System.Text.StringBuilder
        sb.AppendLine("17821 20th ave se")
        sb.AppendLine("Mill Creek, WA 98012")

        DoPrint("Spotty Bowles", sb.ToString)

    End Sub
End Class
