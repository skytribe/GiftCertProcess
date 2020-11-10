Public Class FrmProcessPrint
    Public Property Certificate As ClsGiftCertificate

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sb As New System.Text.StringBuilder
        Dim addressCityStateZip As String = ""
        Try
            If String.IsNullOrEmpty(Certificate.Recipient_Address1) = False Then
                sb.AppendLine(Certificate.Recipient_Address1.Trim)
                If String.IsNullOrEmpty(Certificate.Recipient_Address2) = False Then
                    sb.AppendLine(Certificate.Recipient_Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", Certificate.Recipient_City.Trim.ToProperCase, Certificate.Recipient_State.Trim.ToProperCase, Certificate.Recipient_Zip.Trim)
                sb.AppendLine(addressCityStateZip)
            Else
                sb.AppendLine(Certificate.Purchaser_Address1.Trim)
                If String.IsNullOrEmpty(Certificate.Purchaser_Address2) = False Then
                    sb.AppendLine(Certificate.Purchaser_Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", Certificate.Purchaser_City.Trim.ToProperCase, Certificate.Purchaser_State.Trim.ToProperCase, Certificate.Purchaser_Zip.Trim)
                sb.AppendLine(addressCityStateZip)

            End If
            DoPrint(Certificate.Purchaser_Name, sb.ToString)
            MsgBox("TO BE IMPLEMENTED Print a shipping label - Using Purchaser")
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If Certificate.delivery = DeliveryOptions.InOffice OrElse
                    Certificate.delivery = DeliveryOptions.USMail OrElse
                    Certificate.delivery = DeliveryOptions.USMailDiscrete Then
                PrintCertificate(Certificate, dest:=Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication)
            ElseIf Certificate.delivery = DeliveryOptions.Email Then
                Button1.Enabled = False
                'PrintCertificate(Certificate, dest:=Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication)
                SendEmail(Certificate, destEmail:=
                          TextBox1.Text)
                Button1.Enabled = True
            Else
            End If



            Me.Close()
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub FrmProcessPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Certificate IsNot Nothing Then
            If Certificate.delivery = DeliveryOptions.Email Then
                RdoEmail.Checked = True
                TextBox1.Enabled = True
                TextBox1.Text = Certificate.Purchaser_Email

                '//This will show if email is selected and recipient email is specified allowing for changing
                '// the overide the default destination email as purchaser
                If String.IsNullOrEmpty(Certificate.Recipient_Email.Trim) = False Then
                    lblRecipientEmail.Text = Certificate.Recipient_Email.Trim
                End If
            ElseIf Certificate.delivery = DeliveryOptions.USMail Then
                RdoPrint.Checked = True
                TextBox1.Enabled = True
                TextBox1.Text = ""
            ElseIf Certificate.delivery = DeliveryOptions.USMailDiscrete Then
                RdoPrintDiscrete.Checked = True
                TextBox1.Enabled = True
                TextBox1.Text = ""
            ElseIf Certificate.delivery = DeliveryOptions.InOffice Then
                RdoInPerson.Checked = True
                TextBox1.Enabled = True
                TextBox1.Text = ""
            End If

        Else
            MsgBox("No Certificate specified")
        End If
    End Sub

    Private Sub RdoEmail_CheckedChanged(sender As Object, e As EventArgs) Handles RdoEmail.CheckedChanged, RdoInPerson.CheckedChanged, RdoPrint.CheckedChanged, RdoPrintDiscrete.CheckedChanged
        DetermineControlStatus()
    End Sub

    Private Sub DetermineControlStatus()
        If RdoEmail.Checked Then
            If String.IsNullOrEmpty(TextBox1.Text.Trim) Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
            End If
            Button3.Enabled = False
            TextBox1.Enabled = True
        Else
            Button3.Enabled = True
            Button1.Enabled = True
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        DetermineControlStatus()
    End Sub


End Class