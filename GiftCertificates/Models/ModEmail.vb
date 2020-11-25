Imports System.Net
Imports System.Net.Mail
Imports System.Reflection
Imports System.Text.RegularExpressions

Module ModEmail

#Region "Email Functionality"

    Sub SendEmail(certificate As ClsGiftCertificate2, Optional destEmail As String = "", Optional IsReprint As Boolean = False)
        Dim PublisherDestination = Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication

        Try
            PrintOrderCertificates(certificate, PublisherDestination, ToBePrinted:=False, IsReprint:=IsReprint)

            '//I need to have a bit of a pause to allow for the file to be generated before I can create the email and attach.
            Dim t As New Timer
            t.Tag = DateTime.Now
            TimerEventOccured = False
            t.Enabled = True
            t.Interval = 5000
            AddHandler t.Tick, AddressOf MyTickHandler
            t.Start()
            Do Until TimerEventOccured = True
                My.Application.DoEvents()
            Loop
            t.Stop()


            Dim IntItemsCount = certificate.Item1.Quantity + certificate.Item2.Quantity + certificate.Item3.Quantity + certificate.Item4.Quantity + certificate.Item5.Quantity

            If IsValidEmailFormat(destEmail.Trim) = False AndAlso IsValidEmailFormat(certificate.Billing_Address.Email) = False Then
                Throw New Exception("No valid email address has been provided.")
            End If

            Dim MailMessage As New MailMessage()
            If My.Settings.DevelopmentMode Then
                MsgBox(String.Format("In devmode so will instead send from spottys email to another of spottys emails {0} configured in config file.", My.Settings.DevelopmentModeEmailRecipient))

                MailMessage.From = New MailAddress(My.Settings.SMTPClientUser)
                '//receiver email adress
                MailMessage.To.Add(My.Settings.DevelopmentModeEmailRecipient)
            Else
                ' MsgBox("We will convert to use actual email address specified")
                MailMessage.To.Add(destEmail.Trim)
            End If


            MailMessage.Subject = "Skydive Snohomish Gift Certificate - Thank You!"

            '//attach the file
            'For each gift certificate that was generated on the order
            'Certificate 123 Spotty Bowles.pdb
            For i = 1 To IntItemsCount
                Dim filename = String.Format("Certificate {0}-{1} {2} {3}.pdf", certificate.ID, i, certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
                Dim filepath = System.IO.Path.Combine(My.Settings.PDFOutputFolder, filename)
                If System.IO.File.Exists(filepath) = False Then
                    Throw New Exception("Certificate PDF Not generated for email")
                End If
                MailMessage.Attachments.Add(New Mail.Attachment(filepath))
            Next


            Dim EmailContentPath = GetFileFolderDocumentPath("EmailTemplate.txt")
            Dim Content As String = System.IO.File.ReadAllText(EmailContentPath)
            Dim ContentReplace = Content.Replace("<FirstName>", certificate.Purchaser_FirstName.Trim)

            MailMessage.Body = ContentReplace
            MailMessage.IsBodyHtml = True
            '//SMTP client
            Dim SmtpClient = New SmtpClient(My.Settings.SMTPClient)
            If String.IsNullOrEmpty(My.Settings.SMTPPort) = False Then
                Dim Port = CInt(My.Settings.SMTPPort)
                SmtpClient.Port = Port
            End If

            '//credentials to login in to hotmail account
            SmtpClient.Credentials = New NetworkCredential(My.Settings.SMTPClientUser, My.Settings.SMTPClientPassword)
            '//enabled SSL
            SmtpClient.EnableSsl = True
            '//Send an email
            SmtpClient.Send(MailMessage)

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MsgBox(ex.Message)
        End Try


    End Sub

#Region "Timer Delay Event"
    Dim TimerEventOccured As Boolean = False

    Private Sub MyTickHandler(sender As Object, e As EventArgs)
        TimerEventOccured = True
    End Sub

#End Region

    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Return Regex.IsMatch(s, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
    End Function

#End Region

End Module
