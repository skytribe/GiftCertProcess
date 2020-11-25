Imports System.Reflection

Public Class FrmProcessPrint
    Public Property Certificate As ClsGiftCertificate2
    Public Property IsReprint As Boolean = False

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sb As New System.Text.StringBuilder
        Dim addressCityStateZip As String = ""
        Try
            If String.IsNullOrEmpty(Certificate.Shipping_Address.Address1) = False Then
                sb.AppendLine(Certificate.Shipping_Address.Address1.Trim)
                If String.IsNullOrEmpty(Certificate.Shipping_Address.Address2) = False Then
                    sb.AppendLine(Certificate.Shipping_Address.Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", Certificate.Shipping_Address.City.Trim.ToProperCase, Certificate.Shipping_Address.State.Trim.ToProperCase, Certificate.Shipping_Address.Zip.Trim)
                sb.AppendLine(addressCityStateZip)
            Else
                sb.AppendLine(Certificate.Billing_Address.Address1.Trim)
                If String.IsNullOrEmpty(Certificate.Billing_Address.Address2) = False Then
                    sb.AppendLine(Certificate.Billing_Address.Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", Certificate.Billing_Address.City.Trim.ToProperCase, Certificate.Billing_Address.State.Trim.ToProperCase, Certificate.Billing_Address.Zip.Trim)
                sb.AppendLine(addressCityStateZip)
            End If

            PrintLabel_BrotherPrinter(Certificate.Purchaser_Name, sb.ToString, PrintLabelTypes.Address)

            '//Print Return Label as well if required
            '//
            If RdoPrint.Checked Or RdoInPerson.Checked Then
                If ChkReturnLabel.Checked Then
                    PrintLabel_BrotherPrinter("Sender", My.Settings.ReturnAddress, PrintLabelTypes.ReturnAddress)
                End If
            ElseIf RdoPrintDiscrete.Checked Then
                If ChkReturnLabel.Checked Then
                    PrintLabel_BrotherPrinter("Sender", My.Settings.ReturnAddressDiscreet, PrintLabelTypes.ReturnAddressDiscreet)
                End If
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            '//Verify with the radio buttons
            If RdoEmail.Checked Then
                Certificate.delivery = DeliveryOptions.Email
            ElseIf RdoInPerson.Checked Then
                Certificate.delivery = DeliveryOptions.InOffice
            ElseIf RdoPrint.Checked Then
                Certificate.delivery = DeliveryOptions.USMail
            ElseIf RdoPrintDiscrete.Checked Then
                Certificate.delivery = DeliveryOptions.USDiscreet
            End If
            If ChkUpdateStatus.Checked Then
                Me.IsReprint = False
            Else
                Me.IsReprint = True
            End If

            If Certificate.delivery = DeliveryOptions.InOffice OrElse
                    Certificate.delivery = DeliveryOptions.USMail OrElse
                    Certificate.delivery = DeliveryOptions.USDiscreet Then
                PrintOrderCertificates(Certificate, destination:=Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication, IsReprint:=IsReprint)
            ElseIf Certificate.delivery = DeliveryOptions.Email Then
                Button1.Enabled = False
                SendEmail(Certificate, destEmail:=
                          TextBox1.Text, IsReprint:=IsReprint)
                Button1.Enabled = True
            End If

            Me.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MsgBox(ex)
        End Try

    End Sub

    Private Sub FrmProcessPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupForm()
    End Sub

    Private Sub SetupForm()
        Try
            If Certificate IsNot Nothing Then
                ChkReturnLabel.Checked = True

                If Certificate.delivery = DeliveryOptions.Email Then
                    RdoEmail.Checked = True
                    TextBox1.Enabled = True
                    TextBox1.Text = Certificate.Billing_Address.Email

                    '//This will show if email is selected and recipient email is specified allowing for changing
                    '// the overide the default destination email as purchaser
                    If String.IsNullOrEmpty(Certificate.Shipping_Address.Email.Trim) = False Then
                        lblRecipientEmail.Text = Certificate.Shipping_Address.Email.Trim
                    End If
                ElseIf Certificate.delivery = DeliveryOptions.USMail Then
                    RdoPrint.Checked = True
                    TextBox1.Enabled = True
                    TextBox1.Text = ""
                ElseIf Certificate.delivery = DeliveryOptions.USDiscreet Then
                    RdoPrintDiscrete.Checked = True
                    TextBox1.Enabled = True
                    TextBox1.Text = ""
                ElseIf Certificate.delivery = DeliveryOptions.InOffice Then
                    RdoInPerson.Checked = True
                    TextBox1.Enabled = True
                    TextBox1.Text = ""
                End If

                If IsReprint Then
                    If Certificate.GC_Status = CertificateStatus.Processing Then
                        ChkUpdateStatus.Checked = True
                    ElseIf Certificate.GC_Status = CertificateStatus.Completed Then
                        ChkUpdateStatus.Checked = False
                    End If
                Else
                    If Certificate.GC_Status = CertificateStatus.Entered Then
                        ChkUpdateStatus.Checked = True
                    End If
                End If


            Else
                MessageBox.Show("No Gift Certificate Order specified", "Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
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
            ChkReturnLabel.Enabled = False

        Else

            Button1.Enabled = True
            TextBox1.Enabled = False

            If RdoInPerson.Checked Then
                Button3.Enabled = False
                ChkReturnLabel.Enabled = False
            Else
                Button3.Enabled = True
                ChkReturnLabel.Enabled = True
            End If
        End If

    End Sub

    Private Sub DeliveryOptionControlChanged(sender As Object, e As EventArgs) Handles RdoEmail.CheckedChanged, RdoInPerson.CheckedChanged, RdoPrint.CheckedChanged, RdoPrintDiscrete.CheckedChanged
        Try
            DetermineControlStatus()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            DetermineControlStatus()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub


End Class