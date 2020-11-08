Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Publisher

Imports System.Net.Mail
Imports System.Net
Imports System.Text.RegularExpressions

Module Module3
    Public Sub GetCertificatesToProcess(entrydate As Date, dg As Syncfusion.WinForms.DataGrid.SfDataGrid)
        'Get a list of certificates for a given date.

        Dim gclist1 As New List(Of ClsGiftCertificate)
        gclist1 = RetrieveGiftCertificatesFromQueue(entrydate)

        dg.DataSource = gclist1
        dg.Refresh()


    End Sub

    Public Sub ProcessCertificate(Certificate As ClsGiftCertificate)

        'We have 5 scenarios
        'Both purchase and recipient as new and different
        'Both purchase and recipient as new and the same
        'Purchase is new and recipient is existing
        'Purchase is exising and recipient is new
        'Purchase is new and recipient is existing

        'This will determine if we need to create new customer records
        '

        'Create JumpRun Record if required for purchaser
        'Add payment Record

        'If Certificate.Purchaser_FirstName Then
        'Create Jumprun Record if required for recipient
        'Create a transfer record to recipient


        'UpdateCertificateStatus
        'UpdateCertificateStatus(certificate , CertificateStatus.Processed)

    End Sub
    Public Sub UpdateCertificateJumpRunCustomers(certificate As ClsGiftCertificate, PurchaseID As Integer, RecipientID As Integer)
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "GC_UpdateJRCustomers"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("ID", certificate.ID)
            sqlComm.Parameters.AddWithValue("PurchaserID", PurchaseID)
            sqlComm.Parameters.AddWithValue("RecipientID", RecipientID)

            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("UpdateCertificateStats:", ex)
        Finally

            sqlCon = Nothing
        End Try

    End Sub

    Public Sub UpdateCertificateStatus(certificate As ClsGiftCertificate, status As CertificateStatus)
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GC_UpdateStatus"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("ID", certificate.ID)
            sqlComm.Parameters.AddWithValue("Status", status)
            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("UpdateCertificateStats:", ex)
        Finally

            sqlCon = Nothing
        End Try

    End Sub

    Public Sub PrintCertificate(certificate As ClsGiftCertificate, dest As Microsoft.Office.Interop.Publisher.PbMailMergeDestination)
        Try
            '//Generate Mailmerge File
            Dim MailMergeFile As String = "Mailmerge.txt"


            Dim certificatestoprint As Integer = 0
            certificatestoprint = GenerateMailMergeFile(certificate, MailMergeFile)

            'Write Mailmerge and print using TandemCertificate master publisher document
            If CheckIfRunning("MSPUB") Then
                Dim response = MessageBox.Show("Publisher is open, close before continuing" & Environment.NewLine & "DO you want to close it now ?", "Print Certificate", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                If response = DialogResult.No Then
                    Exit Sub
                Else
                    Try
                        For Each p In Process.GetProcessesByName("MSPUB")
                            p.Kill()
                        Next
                    Catch ex As Exception

                    End Try

                End If
            End If

            Dim appFolder = My.Application.Info.DirectoryPath
            Dim MasterFileFolder = System.IO.Path.Combine(appFolder, "Files")
            Dim CertificateToPrint = System.IO.Path.Combine(MasterFileFolder, "GIFTcERTIFICATE.pub")
            Dim PrintCertificateDocument = System.IO.Path.Combine(appFolder, CertificateToPrint)

            If certificatestoprint > 0 And System.IO.File.Exists(MailMergeFile) Then
                Dim Application As Publisher.Application = New Publisher.Application()
                Application.ActiveWindow.Visible = True
                Dim Path As String = PrintCertificateDocument
                Try
                    Application.Open(Path, False, True)
                    Application.ActiveDocument.MailMerge.Execute(True, Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication)

                    If dest = PbMailMergeDestination.pbSendToPrinter Then
                        Application.Quit()
                    ElseIf dest = PbMailMergeDestination.pbMergeToNewPublication Then
                        'ElseIf dest = PbMailMergeDestination.pbSendEmail Then

                        'Certificate 123 Spotty Bowles.pdb
                        Dim filename = String.Format("Certificate {0} {1} {2}.pdf", certificate.ID, certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
                        Application.ActiveDocument.ExportAsFixedFormat(Format:=PbFixedFormatType.pbFixedFormatTypePDF, Filename:="c:\temp\" & filename)

                        'Application.ActiveDocument.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypePDF, "c:\temp\Cert.pdf")
                        Application.Quit()

                        ', PbFixedFormatIntent.pbIntentStandard, True, -1, -1, -1, -1, -1, -1, -1, True, PbPrintStyle.pbPrintStyleDefault, False, False, False, Nothing)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                If CheckIfRunning("MSPUB") Then

                    Try
                        For Each p In Process.GetProcessesByName("MSPUB")
                            p.Kill()
                        Next
                    Catch ex As Exception

                    End Try


                End If
            End If
            UpdateCertificateStatus(certificate, CertificateStatus.Completed)
        Catch ex As Exception
            WriteException(ex)
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Function GenerateMailMergeFile(certificate As ClsGiftCertificate, mmsource As String) As Integer

        'Details to print on certificate



        Dim str As String = ""
        Dim sb As New StringBuilder
        'Header Line
        sb.AppendLine("Number" & "," & "Recipient" & "," & "10k" & "," & "12k" & "," & "Video" & "," & "Other" & "," & "OtherAmount" & "," & "Total" & "," & "From" & "," & "Authorized")
        Dim certificatestoprint As Integer = 0

        Dim ActualAltitude = ""

        Dim name As String = ""

        Dim GCFrom As String
        If String.IsNullOrEmpty(certificate.Recipient_PersonalizedFrom) Then
            GCFrom = String.Format("{0} {1}", certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
        Else
            GCFrom = certificate.Recipient_PersonalizedFrom.Trim
        End If

        Dim GCTo As String = ""
        If String.IsNullOrEmpty(certificate.Purchaser_PersonalizedTo) Then
            GCTo = String.Format("{0} {1}", certificate.Recipient_FirstName.Trim, certificate.Recipient_LastName.Trim)
        Else
            GCTo = certificate.Purchaser_PersonalizedTo.Trim
        End If


        Dim st1 = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}",
certificate.GC_Number,
GCTo,
certificate.Item_Tandem10k,
certificate.Item_Tandem12k,
certificate.Item_Video,
certificate.Item_Other,
certificate.Item_OtherAmount,
certificate.GC_CalculateTotal,
GCFrom,
certificate.GC_Authorization)

        sb.AppendLine(st1)


        Try
            mmsource = System.IO.Path.Combine(My.Application.Info.DirectoryPath, mmsource)
            Dim StdEncoder As New System.Text.ASCIIEncoding
            Dim orfWriter As System.IO.StreamWriter = New System.IO.StreamWriter(mmsource, False, StdEncoder)
            orfWriter.Write(sb.ToString)
            orfWriter.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Return 1
    End Function


    Public Sub PrintLabel(certificate As ClsGiftCertificate)

    End Sub




    Public Sub WriteException(ex1 As Exception)
        '/./Write Exception to a log file.
        Dim s As String = "----------------------------------------------------------" & Environment.NewLine
        s = s & String.Format("{0:d/M/yyyy HH:mm:ss}", Now) & Environment.NewLine
        s = s & ex1.ToString

        My.Computer.FileSystem.WriteAllText("CheckinManifest.Log", s, True)
        MessageBox.Show(s)
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
            WriteException(ex)
        End Try

        Return isProcessRunning
    End Function



    'Sub JumpRunDeleteCustomer(Id As String)
    '    GetConnectionString()
    '    Dim sqlCon = New SqlConnection(_strConn)
    '    Try

    '        If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()


    '        Dim sqlComm As New SqlCommand()
    '        sqlComm.Connection = sqlCon

    '        sqlComm.CommandText = "dbo.PeopleDelete"

    '        sqlComm.CommandType = CommandType.StoredProcedure


    '        sqlComm.Parameters.AddWithValue("wCustId_1", Id)
    '        sqlComm.Parameters.AddWithValue("wBillToId_2", Id)


    '        sqlComm.ExecuteNonQuery()

    '        If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        LogError("RemoveFromJumpRun:", ex)
    '    Finally

    '        sqlCon = Nothing
    '    End Try
    'End Sub


    'Private Function CalculatePreferredName(ByVal c1 As JumpRunUpdate, Optional Editaction As Boolean = False) As JumpRunUpdate
    '    Dim c As JumpRunUpdate = c1
    '    Try
    '        Dim sqlCon1 = New SqlConnection(_strConn)
    '        Dim icountmatches As Integer = 0
    '        Dim actualfirstname = c1.FirstName
    '        Dim preferredname = c1.PreferredName

    '        If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

    '        Using cmd As New SqlClient.SqlCommand("dbo.MatchingCustomerNamesExactCount", sqlCon1)
    '            cmd.CommandType = CommandType.StoredProcedure

    '            cmd.Parameters.AddWithValue("FirstName", c.FirstName)
    '            cmd.Parameters.AddWithValue("LastName", c.LastName)
    '            cmd.Parameters.Add("RowCount", SqlDbType.Int).Direction = ParameterDirection.Output
    '            cmd.ExecuteNonQuery()
    '            icountmatches = Convert.ToInt32(cmd.Parameters("RowCount").Value)
    '        End Using

    '        If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

    '        '//THis is editing existing record - so no need to recalculate a new name
    '        If Editaction = True And icountmatches <= 1 Then
    '            Return c
    '        End If

    '        Dim s = Asc(c.FirstName.Trim.Last)
    '        If s >= 48 And s <= 57 Then
    '            'First Name ends in a Number - so lets just leave it alone rather than
    '            'trying to recalculate as most firstnames dont end in a number
    '        Else
    '            Dim SimilarFirstNames As New List(Of PreferredNameFirstNamePair)
    '            If icountmatches > 0 Then
    '                Dim sqlComm1 As New SqlCommand()
    '                sqlComm1.Connection = sqlCon1
    '                sqlComm1.CommandText = "dbo.MatchingCustomerNames"
    '                sqlComm1.CommandType = CommandType.StoredProcedure
    '                sqlComm1.Parameters.AddWithValue("FirstName", c.FirstName)
    '                sqlComm1.Parameters.AddWithValue("LastName", c.LastName)

    '                If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

    '                Dim sqlReader As SqlDataReader = sqlComm1.ExecuteReader()
    '                Dim icount As Integer = 0

    '                'Get all the customer name matches
    '                ' because there may be multiple record for John Smith or Tdm-John Smith


    '                'If its an edit then use the customer ID and leave the preferred name the same
    '                'If its new then find the next unique number

    '                If sqlReader.HasRows Then
    '                    While (sqlReader.Read())
    '                        icount = icount + 1
    '                        SimilarFirstNames.Add(New PreferredNameFirstNamePair(CInt(sqlReader.GetInt32(0)), sqlReader.GetString(39), sqlReader.GetString(1)))
    '                    End While
    '                End If

    '                If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

    '                If Editaction = True Then
    '                    '//Edit
    '                    If SimilarFirstNames.Count > 0 Then
    '                        Dim x = (From i1 In SimilarFirstNames Where i1.scust = c1.wCustId Select i1).First
    '                        If x IsNot Nothing Then
    '                            c.PreferredName = x.preferedname
    '                        Else
    '                            LogError("CalculatePreferredName:", "No records matching with customer ID in SimilarFirstNames (1)")
    '                        End If
    '                    Else
    '                        LogError("CalculatePreferredName:", "No records in SimilarFirstNames (2)")
    '                    End If


    '                Else
    '                    '//New
    '                    '//Determine the non matching firstname with number and adjust the preferredName

    '                    For i = 1 To 100
    '                        If SimilarFirstNames.Count > 0 Then
    '                            c.PreferredName = String.Format("{0},{1}{2}", c.LastName, actualfirstname, i.ToString)

    '                            If DoesIndexExist(c.PreferredName) = False Then
    '                                Exit For
    '                            End If

    '                        Else
    '                            If DoesIndexExist(c.PreferredName) = False Then
    '                                Exit For
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            End If

    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        LogError("CalculatePreferredName:", ex)
    '    End Try

    '    Return c
    'End Function

    ''//Verifies if the People table contains this string which is a unique Key
    'Public Function DoesIndexExist(customerKeyString As String) As Boolean
    '    GetConnectionString()
    '    Dim sqlCommand As New SqlCommand

    '    Dim sqlconnect As SqlConnection = Nothing

    '    Try
    '        sqlconnect = New SqlConnection(_strConn)
    '        sqlconnect.Open()
    '        sqlCommand.CommandText = String.Format("Select * from tPeople WHERE sCust = '{0}'", customerKeyString)
    '        sqlCommand.CommandType = CommandType.Text
    '        sqlCommand.Connection = sqlconnect

    '        Dim ds = New DataSet()
    '        Dim adap = New SqlDataAdapter(sqlCommand)
    '        adap.Fill(ds, "People")

    '        If ds.Tables(0).Rows.Count > 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        LogError("DoesIndexExist:", ex)
    '    Finally
    '        sqlconnect.Close()

    '    End Try
    '    Return False
    'End Function


    'Public Function InsertNewRecord(c As JumpRunUpdate) As Boolean
    '    Dim SuccessState As Boolean = False
    '    Dim RecordDetail = "Insert {0} {1}  - {2} : {3}"
    '    GetConnectionString()
    '    Try
    '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


    '        _sqlCon = New SqlConnection(_strConn)
    '        If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

    '        Dim c1 = CalculatePreferredName(c)
    '        c.PreferredName = c1.PreferredName.ToProperCase
    '        c.FirstName = c1.FirstName.ToProperCase

    '        'Just leave the alpha characters and see if the last name is contained in
    '        'the Nickname (Preferredname) - if its not stop and error rather than continue on the update.
    '        If c.PreferredName.SanitizeString.ToLower.Contains(c.LastName.SanitizeString.ToLower) = False Then
    '            Throw New Exception("Last name is not included in the Nickname, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
    '        ElseIf String.IsNullOrEmpty(c.PreferredName.SanitizeString) Then
    '            Throw New Exception("Nickname does not include names, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
    '        End If

    '        Dim sqlComm As New SqlCommand()
    '        sqlComm.Connection = _sqlCon
    '        Dim IsLicensedJumper As Boolean = False

    '        '//USPA Number determines is licensed or student
    '        If String.IsNullOrEmpty(c.USPANumber) = False Then
    '            IsLicensedJumper = True
    '        End If
    '        If IsLicensedJumper = False Then
    '            sqlComm.CommandText = "dbo.rwjr_ExportNewUser"
    '        Else
    '            sqlComm.CommandText = "dbo.rwjr_ExportNewUserLicensedSmartWaiver"
    '        End If
    '        sqlComm.CommandType = CommandType.StoredProcedure

    '        sqlComm.Parameters.AddWithValue("PreferredName", c.PreferredName.ToProperCase)
    '        If IsLicensedJumper = True Then
    '            sqlComm.Parameters.AddWithValue("Student ", determineStudent(c.USPANumber.SanitizeString)) 'c.Student)
    '        Else
    '            sqlComm.Parameters.AddWithValue("Student ", 1)
    '        End If
    '        sqlComm.Parameters.AddWithValue("FirstName", c.FirstName.ToProperCase.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("LastName", c.LastName.ToProperCase.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("MI", c.MI.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("Phone1", c.Phone1)
    '        sqlComm.Parameters.AddWithValue("Street1", c.Street1.ToProperCase.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("City", c.City.ToProperCase.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("State", c.State.ToProperCase.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("Zip", c.Zip.SanitizeString)
    '        sqlComm.Parameters.AddWithValue("Email", c.Email.ToProperCase)
    '        sqlComm.Parameters.AddWithValue("EmergencyContact ", c.EmergencyContact.ToProperCase)
    '        sqlComm.Parameters.AddWithValue("EmercencyPhone ", c.EmercencyPhone)
    '        sqlComm.Parameters.AddWithValue("Birth", c.Birth)
    '        If c.Birth <= New DateTime(1920, 1, 1) Then
    '            LogError("InsertNewRecord:date", "")
    '            LogError("Birth=" & c.Birth, "")
    '        End If
    '        sqlComm.Parameters.AddWithValue("Gender", c.Gender)
    '        sqlComm.Parameters.AddWithValue("CountryId", c.CountryId)
    '        sqlComm.Parameters.AddWithValue("Referrer", c.Referrer)
    '        sqlComm.Parameters.AddWithValue("WorkEmail", c.WorkEmail)
    '        sqlComm.Parameters.AddWithValue("ManifestWt", c.Weight)

    '        If IsLicensedJumper Then
    '            '//Additional Fields
    '            sqlComm.Parameters.AddWithValue("USPANumber", c.USPANumber)
    '            sqlComm.Parameters.AddWithValue("USPALicense", c.LicenseNumber)
    '            sqlComm.Parameters.AddWithValue("TotalJumps", c.TotalJumps)
    '        End If

    '        If My.Settings.UseWaiverDate Then
    '            If c.WaiverDate Is Nothing Then
    '                sqlComm.Parameters.AddWithValue("WaiverDate", Now)
    '            Else
    '                sqlComm.Parameters.AddWithValue("WaiverDate", c.WaiverDate)
    '            End If
    '        Else
    '            sqlComm.Parameters.AddWithValue("WaiverDate", Now)
    '        End If
    '        sqlComm.Parameters.AddWithValue("sOpInsert", "SmartWaive")
    '        sqlComm.Parameters.AddWithValue("SmartWaiverID", c.SmartWaiverID)
    '        sqlComm.ExecuteNonQuery()
    '        SuccessState = True

    '        UpdateImportTable(c.SmartWaiverID, c.wCustId)

    '        If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        LogError("InsertNewRecord:", ex)
    '        LogError("PreferredName=" & c.PreferredName, "")
    '        LogError(String.Format("FirstName = {0}, LastName = {1}", c.FirstName, c.LastName), "")
    '    Finally
    '        _sqlCon = Nothing
    '    End Try
    '    Return SuccessState
    'End Function


    'Public Function UpdateExistingRecord(c As JumpRunUpdate) As Boolean
    '    Dim SuccessState As Boolean = False
    '    Dim RecordDetail = "Update {0} {1}  - {2} : {3}"
    '    Try

    '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


    '        GetConnectionString()

    '        _sqlCon = New SqlConnection(_strConn)

    '        '//get the preferredname from jumprun
    '        Dim ds As DataSet = ReturnJumpRunItem(c.wCustId)

    '        Dim currentItem As String = ds.Tables(0).Rows(0).ItemArray(39)


    '        If currentItem.ToUpper.StartsWith("TDM") OrElse
    '            currentItem.ToUpper.StartsWith("SL") OrElse
    '            currentItem.ToUpper.StartsWith("COACH") OrElse
    '            currentItem.ToUpper.StartsWith("VIDEO") Then
    '            c.PreferredName = currentItem
    '            c.FirstName = c.FirstName.ToProperCase
    '        Else
    '            Dim c1 = CalculatePreferredName(c, True)
    '            c.PreferredName = c1.PreferredName.ToProperCase
    '            c.FirstName = c.FirstName.ToProperCase.SanitizeString
    '        End If

    '        '//Existing nickname - does it have a prefix TDM - ,SL - , COACH - VID - 


    '        'Dim c1 = CalculatePreferredName(c, True)   '//If edit it should exist, if new then shouldnt
    '        'c.PreferredName = c1.PreferredName
    '        'c.FirstName = c1.FirstName

    '        Using (_sqlCon)
    '            Dim sqlComm As New SqlCommand()

    '            _sqlCon.Open()

    '            sqlComm.Connection = _sqlCon
    '            Dim IsLicensedJumper As Boolean = False
    '            If String.IsNullOrEmpty(c.USPANumber) = False Then
    '                IsLicensedJumper = True
    '                sqlComm.CommandText = "spjr_EditExsistingLicensedCustomerSmartWaiver"

    '            Else
    '                sqlComm.CommandText = "spjr_EditExsistingCustomerSmartWaiver"
    '            End If


    '            'Just leave the alpha characters and see if the last name is contained in
    '            'the Nickname (Preferredname) - if its not stop and error rather than continue on the update.
    '            If c.PreferredName.SanitizeString.ToLower.Contains(c.LastName.SanitizeString.ToLower) = False Then
    '                Throw New Exception("Last name is not included in the Nickname, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
    '            ElseIf String.IsNullOrEmpty(c.PreferredName.SanitizeString) Then
    '                Throw New Exception("Nickname does not include names, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
    '            End If

    '            sqlComm.CommandType = CommandType.StoredProcedure
    '            sqlComm.Parameters.AddWithValue("wCustId", c.wCustId)
    '            sqlComm.Parameters.AddWithValue("PreferredName", c.PreferredName.ToProperCase)
    '            sqlComm.Parameters.AddWithValue("FirstName", c.FirstName.ToProperCase.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("LastName", c.LastName.ToProperCase.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("MI", c.MI.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("Phone1", c.Phone1 & "")
    '            sqlComm.Parameters.AddWithValue("Street1", c.Street1.ToProperCase.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("City", c.City.ToProperCase.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("State", c.State.ToProperCase.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("Zip", c.Zip.SanitizeString & "")
    '            sqlComm.Parameters.AddWithValue("Email", c.Email.ToProperCase & "")
    '            sqlComm.Parameters.AddWithValue("WorkEmail", c.WorkEmail & "")
    '            sqlComm.Parameters.AddWithValue("EmergencyContact ", c.EmergencyContact.ToProperCase & "")
    '            sqlComm.Parameters.AddWithValue("EmercencyPhone ", c.EmercencyPhone & "")
    '            sqlComm.Parameters.AddWithValue("Birth", c.Birth)
    '            sqlComm.Parameters.AddWithValue("Gender", c.Gender)
    '            sqlComm.Parameters.AddWithValue("CountryId", c.CountryId)
    '            sqlComm.Parameters.AddWithValue("ManifestWt", c.Weight)
    '            sqlComm.Parameters.AddWithValue("Referrer", c.Referrer)


    '            '//Additional Fields
    '            If IsLicensedJumper Then
    '                sqlComm.Parameters.AddWithValue("Student ", determineStudent(c.USPANumber)) 'c.Student)
    '                sqlComm.Parameters.AddWithValue("USPANumber", c.USPANumber.SanitizeString)
    '                '  sqlComm.Parameters.AddWithValue("USPAExpiresDate", c.USPAExpiresDate)
    '                sqlComm.Parameters.AddWithValue("USPALicense", c.LicenseNumber)
    '                'sqlComm.Parameters.AddWithValue("LastJumpDate", c.LastJumpDate)
    '                sqlComm.Parameters.AddWithValue("TotalJumps", c.TotalJumps)
    '                'sqlComm.Parameters.AddWithValue("LastRepackDate", c.LastRepackDate)
    '            Else
    '                sqlComm.Parameters.AddWithValue("Student ", 1)
    '            End If

    '            If My.Settings.UseWaiverDate Then
    '                If c.WaiverDate Is Nothing Then
    '                    sqlComm.Parameters.AddWithValue("WaiverDate", Now)
    '                Else
    '                    sqlComm.Parameters.AddWithValue("WaiverDate", c.WaiverDate)
    '                End If
    '            Else
    '                sqlComm.Parameters.AddWithValue("WaiverDate", Now)
    '            End If
    '            sqlComm.Parameters.AddWithValue("sOpInsert", "SmartWaive")
    '            sqlComm.Parameters.AddWithValue("SmartWaiverID", c.SmartWaiverID)
    '            sqlComm.ExecuteNonQuery()
    '            SuccessState = True

    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        LogError("UpdateExistingRecord:", ex)
    '    Finally
    '        _sqlCon = Nothing
    '    End Try
    '    Return SuccessState
    'End Function

    Sub SendEmail(certificate As ClsGiftCertificate, Optional destEmail As String = "")
        Dim publisherdest = Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication

        Try
            PrintCertificate(certificate, publisherdest)

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
                't.Stop()
            Loop

            t.Stop()

            'Certificate 123 Spotty Bowles.pdb
            Dim filename = String.Format("Certificate {0} {1} {2}.pdf", certificate.ID, certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
            Dim filepath = System.IO.Path.Combine("c:\temp\", filename)

            If System.IO.File.Exists(filepath) = False Then
                Throw New Exception("Certificate PDF Not generated for email")
            End If

            If IsValidEmailFormat(destEmail.Trim) = False AndAlso IsValidEmailFormat(certificate.Purchaser_Email) = False Then
                Throw New Exception("No valid email address has been provided.")
            End If


            Dim MailMessage As New MailMessage()
            If DEVMODE Then
                MsgBox("In devmode so will instead send from spottys email to another of spottys emails")

                MailMessage.From = New MailAddress("skytribe@hotmail.com")
                '//receiver email adress

                MailMessage.To.Add("skytribe@spottysworld.com")
            Else
                MsgBox("We will convert to use actual email address specified")
            End If


            MailMessage.Subject = "Skydive Snohomish Gift Certificate - Thank You!"

            '//attach the file
            MailMessage.Attachments.Add(New Mail.Attachment(filepath))

            Dim appFolder = My.Application.Info.DirectoryPath
            Dim MasterFileFolder = System.IO.Path.Combine(appFolder, "Files")
            Dim EmailContentPath = System.IO.Path.Combine(MasterFileFolder, "EmailTemplate.txt")


            Dim Content As String = System.IO.File.ReadAllText(EmailContentPath)
            Dim ContentReplace = Content.Replace("<FirstName>", certificate.Purchaser_FirstName.Trim)

            MailMessage.Body = ContentReplace
            MailMessage.IsBodyHtml = True
            '//SMTP client
            Dim SmtpClient = New SmtpClient("smtp.live.com")
            '//port number for Hot mail
            SmtpClient.Port = 587 ' 25
            'SmtpServer.Port = 587

            '//credentials to login in to hotmail account
            SmtpClient.Credentials = New NetworkCredential("skytribe@hotmail.com", "Lightning160")
            '//enabled SSL
            SmtpClient.EnableSsl = True
            '//Send an email
            SmtpClient.Send(MailMessage)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Dim TimerEventOccured As Boolean = False

    Private Sub MyTickHandler(sender As Object, e As EventArgs)
        TimerEventOccured = True
    End Sub

    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Return Regex.IsMatch(s, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
    End Function
End Module
