Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Publisher

Imports System.Net.Mail
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Reflection

Module Module3


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

    Public Sub UpdateGCOrderAuthorizer(certificate As ClsGiftCertificate2, Authorizer As String)
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

    '//Reset After Test
    '//Use gcProcess sOpInsert to identify
    '//
    '//Remove any newly created People or PeopleAncillary
    '//Remove All GiftCertificate Tables
    '//Remove Inv or InvAll 
    '//Remove Payment or PaymentAll
    '//Remove Redeem
    '

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

End Module
