Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Publisher

Imports System.Net.Mail
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Drawing.Printing

Module Module3
    Function SetDefaulPrinter(ByVal strPrinterName As String, Optional strDefaultTray As String = "", Optional SingleSidedOnly As Boolean = False) As Boolean
        Dim strCurrPrinter As String
        Dim WsNetwork As Object
        Dim prntDoc As New PrintDocument

        strCurrPrinter = prntDoc.PrinterSettings.PrinterName
        WsNetwork = Microsoft.VisualBasic.CreateObject("WScript.Network")

        Try
            WsNetwork.SetDefaultPrinter("Canon MF4500w Series")

            prntDoc.PrinterSettings.PrinterName = strPrinterName
            If strDefaultTray = "" Then
            Else
                ' step through the available paper sources and set it to the right one
                For psrc = 0 To prntDoc.PrinterSettings.PaperSources.Count - 1
                    If prntDoc.PrinterSettings.PaperSources(psrc).SourceName = strDefaultTray Then
                        prntDoc.PrinterSettings.DefaultPageSettings.PaperSource = prntDoc.PrinterSettings.PaperSources.Item(psrc)
                        prntDoc.PrinterSettings.Duplex = Duplex.Simplex
                    End If
                Next

            End If
            If SingleSidedOnly Then
                prntDoc.PrinterSettings.Duplex = Duplex.Simplex
            End If

            'set default if selected printer name is a valid installed printer
            If prntDoc.PrinterSettings.IsValid Then
                Return True
            Else
                WsNetwork.SetDefaultPrinter(strCurrPrinter)
                Return False
            End If
        Catch ex As Exception
            WsNetwork.SetDefaultPrinter(strCurrPrinter)
            Return False
        Finally
            WsNetwork = Nothing
            prntDoc = Nothing
        End Try
    End Function
    Function GetDefaultPrinter() As String
        Dim CurrentDefault As String = ""

        Dim prntDoc As New PrintDocument

        'check if there is installed printer
        If PrinterSettings.InstalledPrinters.Count = 0 Then
            MsgBox("No printer installed")
            Return CurrentDefault

        End If


        Return prntDoc.PrinterSettings.PrinterName

    End Function
    Public Sub UpdateCertificateJumpRunCustomers(certificate As ClsGiftCertificate2, PurchaseID As Integer)
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_UpdateJRCustomers"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("ID", certificate.ID)
            sqlComm.Parameters.AddWithValue("PurchaserID", PurchaseID)
            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            sqlCon = Nothing
        End Try
    End Sub

    Public Sub UpdatePricing(pricing As ClsPricing, IsItem As Boolean)
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon1 = New SqlConnection(_strConn)
        Try
            If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon1

            sqlComm.CommandText = "dbo.GCO_UpdatePricing"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("ID", pricing.ID)
            sqlComm.Parameters.AddWithValue("SKU", pricing.SKU)
            sqlComm.Parameters.AddWithValue("Price", pricing.Price)
            sqlComm.Parameters.AddWithValue("JumpRunItemId", pricing.JR_ItemID)
            sqlComm.Parameters.AddWithValue("DiscountableItem", pricing.Discountable)
            sqlComm.Parameters.AddWithValue("IsItem", IsItem)

            sqlComm.ExecuteNonQuery()

            If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon1 = Nothing
        End Try

    End Sub
    Public Sub UpdateGCOrderStatus(certificate As ClsGiftCertificate2, status As CertificateStatus, pdate As Date)
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_UpdateStatus"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("ID", certificate.ID)
            sqlComm.Parameters.AddWithValue("Status", status)
            sqlComm.Parameters.AddWithValue("ProcessDate", pdate)
            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon = Nothing
        End Try

    End Sub
    Public Sub InsertPromoPricing(pricing As ClsPromoPricing)
        GetConnectionString()
        Dim sqlCon1 = New SqlConnection(_strConn)
        Try
            If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon1

            sqlComm.CommandText = "dbo.GCO_InsertGiftCertificatePricingPromotions"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("ID", pricing.ID)
            sqlComm.Parameters.AddWithValue("PromoDescription", pricing.PromoDescription)
            sqlComm.Parameters.AddWithValue("DiscountCode", "")
            sqlComm.Parameters.AddWithValue("itemCode1", pricing.ItemCode1)
            sqlComm.Parameters.AddWithValue("itemPricing1", pricing.ItemPrice1)
            sqlComm.Parameters.AddWithValue("itemCode2", pricing.ItemCode2)
            sqlComm.Parameters.AddWithValue("itemPricing2", pricing.ItemPrice2)
            sqlComm.Parameters.AddWithValue("itemCode3", pricing.ItemCode3)
            sqlComm.Parameters.AddWithValue("itemPricing3", pricing.ItemPrice3)
            sqlComm.Parameters.AddWithValue("itemCode4", pricing.ItemCode4)
            sqlComm.Parameters.AddWithValue("itemPricing4", pricing.ItemPrice4)
            sqlComm.Parameters.AddWithValue("itemCode5", pricing.ItemCode5)
            sqlComm.Parameters.AddWithValue("itemPricing5", pricing.ItemPrice5)
            sqlComm.Parameters.AddWithValue("status", pricing.Status)
            sqlComm.Parameters.AddWithValue("displayInList", 1)

            sqlComm.ExecuteNonQuery()

            If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon1 = Nothing
        End Try

    End Sub
    Public Sub UpdatePromoPricing(pricing As ClsPromoPricing)
        GetConnectionString()
        Dim sqlCon1 = New SqlConnection(_strConn)
        Try
            If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon1

            sqlComm.CommandText = "dbo.GCO_UpdateGiftCertificatePricingPromotions"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("ID", pricing.ID)
            sqlComm.Parameters.AddWithValue("PromoDescription", pricing.PromoDescription)
            sqlComm.Parameters.AddWithValue("DiscountCode", "")
            sqlComm.Parameters.AddWithValue("itemCode1", pricing.ItemCode1)
            sqlComm.Parameters.AddWithValue("itemPricing1", pricing.ItemPrice1)
            sqlComm.Parameters.AddWithValue("itemCode2", pricing.ItemCode2)
            sqlComm.Parameters.AddWithValue("itemPricing2", pricing.ItemPrice2)
            sqlComm.Parameters.AddWithValue("itemCode3", pricing.ItemCode3)
            sqlComm.Parameters.AddWithValue("itemPricing3", pricing.ItemPrice3)
            sqlComm.Parameters.AddWithValue("itemCode4", pricing.ItemCode4)
            sqlComm.Parameters.AddWithValue("itemPricing4", pricing.ItemPrice4)
            sqlComm.Parameters.AddWithValue("itemCode5", pricing.ItemCode5)
            sqlComm.Parameters.AddWithValue("itemPricing5", pricing.ItemPrice5)
            sqlComm.Parameters.AddWithValue("status", Math.Abs(pricing.Status))
            sqlComm.Parameters.AddWithValue("displayInList", Math.Abs(pricing.DisplayInList))

            sqlComm.ExecuteNonQuery()

            If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon1 = Nothing
        End Try

    End Sub
    Public Sub UpdateGCOrderAuthorizer(certificate As ClsGiftCertificate2, Authorizer As String, Optional personalizedFrom As String = "")
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_UpdateAuthorizer"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("ID", certificate.ID)
            sqlComm.Parameters.AddWithValue("Authorizer", Left(Authorizer.Trim.Trim, 10))
            sqlComm.Parameters.AddWithValue("PersonalizedFrom", Left(personalizedFrom.Trim, 255))
            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon = Nothing
        End Try

    End Sub

    Friend Function CheckIfRunning(name As String) As Boolean
        Dim isProcessRunning As Boolean = False

        Try
            Dim p() As Process
            Dim x = Process.GetProcesses

            p = Process.GetProcessesByName(name)
            If p.Count > 0 Then
                ' Process is running
                isProcessRunning = True
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

        End Try

        Return isProcessRunning
    End Function

    Sub SendImportDetails(strRecordfilepath As String)

        Try


            Dim MailMessage As New MailMessage()

            MailMessage.From = New MailAddress("SkyTribe@hotmail.com")
            '//receiver email adress
            MailMessage.To.Add(BCC_Email)


            MailMessage.Subject = " Gift Certificate Process Import Records"

            MailMessage.Attachments.Add(New Mail.Attachment(strRecordfilepath))

            MailMessage.Body = "Import Process File"
            MailMessage.IsBodyHtml = True
            '//SMTP client
            Dim SmtpClient = New SmtpClient("smtp.live.com")
            SmtpClient.Port = 587

            '//credentials to login in to hotmail account
            SmtpClient.Credentials = New NetworkCredential("SkyTribe@hotmail.com", "Icarus365")
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

    Sub SendProcessDetails(strRecord As String)

        Try
            Dim MailMessage As New MailMessage()

            MailMessage.From = New MailAddress("SkyTribe@hotmail.com")
            '//receiver email adress
            MailMessage.To.Add(BCC_Email)


            MailMessage.Subject = " Gift Certificate Process Process Record"

            MailMessage.Body = "Import Process File" & Environment.NewLine & Environment.NewLine & strRecord
            MailMessage.IsBodyHtml = True
            '//SMTP client
            Dim SmtpClient = New SmtpClient("smtp.live.com")
            SmtpClient.Port = 587

            '//credentials to login in to hotmail account
            SmtpClient.Credentials = New NetworkCredential("SkyTribe@hotmail.com", "Icarus365")
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

    Function GetAuthorizerCode(s As String) As String
        'Locate parenthesis and extract contents

        If s.Contains("(") And s.Contains(")") Then
            Try
                Dim scodestart = s.IndexOf("(")
                Dim scodeend = s.IndexOf(")")

                Dim length1 = scodeend - scodestart
                Dim scode = s.Substring(scodestart + 1, length1 - 1)
                Return scode.Trim
            Catch ex As Exception
                Return ""
            End Try

        Else
            Return s.Trim
        End If
    End Function

    Sub SetupPricing()
        'THIS IS USED TO VERIFY THAT THE PRCING ITEMS IN JUMPRUN ARE PRESENT 
        '88,145,91,90,618  - STANDARD PRCING
        '637, 632, 633, 634  - DISCOUNT PRICES
        '638,639,640,641, 618 - 50 BUCKS OF ALL

    End Sub
End Module
