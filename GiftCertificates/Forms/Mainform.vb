Imports System.Net
Imports Syncfusion.WinForms.Input.Events

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
            ProcessGiftCertifcateQueueForm.DefaultDate = SfDateTimeEdit1.Value
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
            Dim ObjReprint As New FrmReprint
            ObjReprint.ShowDialog()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            Dim ObjReprint As New FrmReprintShippingLabel
            ObjReprint.ShowDialog()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim x As New FrmDevelopmentReset
        x.ShowDialog()
    End Sub

    Private Sub Mainform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setupForm()

    End Sub

    Private Sub setupForm()
        ValidateLicensing()

        Me.Text = "Gift Certificates Main Form"
        Label1.Text = "Current Date"
        SfDateTimeEdit1.Value = Now
        ValidateDateForWarning(SfDateTimeEdit1.Value)

        If My.Settings.DevelopmentMode Then
            Panel1.Visible = True
        Else
            Panel1.Visible = False
        End If
    End Sub

    Private Sub ValidateDateForWarning(dte As Date)
        If dte.Date <> GetBusinessDate.Date Then
            lblWarningDate.Text = String.Format("The current date {0:d} does not match the current business date {1:d} in JumpRun." & Environment.NewLine & "Update business date in Jumprun to Process gift certificate orders on this date.", dte, GetBusinessDate())
        Else
            lblWarningDate.Text = ""
        End If
    End Sub

    Private Sub ValidateLicensing()
        If Date.Now >= New Date(2020, 12, 31) Then
            MessageBox.Show("App is expired and needs to be reinstalled.", "Gift Certificate Processing", MessageBoxButtons.OK)
            LogError("App Licensing Expired", CType(Nothing, Exception))
            End
        End If

        If Not RemoteFileExists("https://1drv.ms/t/s!ANNFQZ6-n-pbl5R3") Then
            End
        End If
    End Sub
    Private Function RemoteFileExists(ByVal url As String) As Boolean
        Try
            Dim request As HttpWebRequest = TryCast(WebRequest.Create(url), HttpWebRequest)
            request.Method = "GET"
            request.AllowAutoRedirect = True

            request.AutomaticDecompression = DecompressionMethods.Deflate Or DecompressionMethods.GZip
            Dim response As HttpWebResponse = TryCast(request.GetResponse(), HttpWebResponse)
            response.Close()
            Return (response.StatusCode = HttpStatusCode.OK)
        Catch ex As Exception
            LogError("RemoteFileExists:", ex)
            Return False
        End Try
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim sb As New System.Text.StringBuilder
        sb.AppendLine("17821 20th ave se")
        sb.AppendLine("Mill Creek, WA 98012")

        PrintLabel_BrotherPrinter("Spotty Bowles", sb.ToString, 0)

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim x As New FrmReport
        x.ShowDialog()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim ObjIncomplete As New FrmIncompleteItems
        ObjIncomplete.ShowDialog()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim ObjPricing As New FrmPricing
        ObjPricing.ShowDialog()

    End Sub

    Private Sub SfDateTimeEdit1_ValueChanged(sender As Object, e As DateTimeValueChangedEventArgs) Handles SfDateTimeEdit1.ValueChanged
        ValidateDateForWarning(SfDateTimeEdit1.Value)
    End Sub
End Class
