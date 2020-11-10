Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

Public Module Module4


    Public _strConn As String = "Data Source=Localhost\SQLEXPRESS8;Initial Catalog=TrainJumpRun;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"
    Private _sqlCon As SqlConnection
    Public Property ErrorString1 As String = ""
    Public BlnDevMode = False

    <Extension>
    Public Function getString(o As Object) As String
        If IsDBNull(o) Then
            Return ""
        End If

        Return CStr(o) & ""
    End Function

    <Extension>
    Public Function getNullableDate(o As Object) As Date?
        If IsDBNull(o) Then
            Return Nothing
        End If

        Return CDate(o)
    End Function

    Sub GetConnectionString()
        Try
            _strConn = _strConn ' My.Settings.ConnectionString
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("GetConnectionString:", ex)
        End Try
    End Sub




    Public Function RetrieveGiftCertificatesFromQueue(d As Date, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            If devmode = False Then
                Dim sqlstring = String.Format("Select * from dbo.GiftCertificates where DateEntered >= '{0} 12:00:00 AM' AND DateEntered <= '{0} 11:59:59 PM'", d.Date.ToShortDateString, d.Date.ToShortDateString)


                cmd.CommandText = sqlstring
            Else
                cmd.CommandText = "Select * from dbo.GiftCertificates"
            End If

            'cmd.CommandText = "Select * from dbo.GiftCertificates"
            cmd.CommandType = CommandType.Text
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim PossibleRecord As New ClsGiftCertificate


                PossibleRecord.ID = reader("Id").ToString()

                PossibleRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                PossibleRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                PossibleRecord.Purchaser_Address1 = reader("Purchaser_Address1").ToString()
                PossibleRecord.Purchaser_Address2 = reader("Purchaser_Address2").ToString()
                PossibleRecord.Purchaser_City = reader("Purchaser_City").ToString()
                PossibleRecord.Purchaser_State = reader("Purchaser_State").ToString()
                PossibleRecord.Purchaser_Zip = reader("Purchaser_Zip").ToString()
                PossibleRecord.Purchaser_Phones1 = reader("Purchaser_Phone1").ToString()
                PossibleRecord.Purchaser_Phones2 = reader("Purchaser_Phone2").ToString()
                PossibleRecord.Purchaser_Email = reader("Purchaser_Email").ToString()

                PossibleRecord.Recipient_FirstName = reader("Recipient_FirstName").ToString()
                PossibleRecord.Recipient_LastName = reader("Recipient_LastName").ToString()
                PossibleRecord.Recipient_Address1 = reader("Recipient_Address1").ToString()
                PossibleRecord.Recipient_Address2 = reader("Recipient_Address2").ToString()
                PossibleRecord.Recipient_City = reader("Recipient_City").ToString()
                PossibleRecord.Recipient_State = reader("Recipient_State").ToString()
                PossibleRecord.Recipient_Zip = reader("Recipient_Zip").ToString()
                PossibleRecord.Recipient_Phones1 = reader("Recipient_Phone1").ToString()
                PossibleRecord.Recipient_Phones2 = reader("Recipient_Phone2").ToString()
                PossibleRecord.Recipient_Email = reader("Recipient_Email").ToString()
                PossibleRecord.Item_Tandem10k = reader("Item_Tandem10k").ToString()
                PossibleRecord.Item_Tandem12k = reader("Item_Tandem12k").ToString()
                PossibleRecord.Item_Video = reader("Item_Video").ToString()
                PossibleRecord.Item_Other = reader("Item_Other").ToString()
                PossibleRecord.Item_OtherAmount = reader("Item_OtherAmount").ToString()

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    PossibleRecord.JR_PurchaseID = reader("JR_PurchaseID")
                End If
                If IsDBNull(reader("JR_RecipientID")) = False Then
                    PossibleRecord.JR_RecipientID = reader("JR_RecipientID")
                End If
                PossibleRecord.GC_DateEntered = reader("DateEntered").ToString()
                PossibleRecord.HearAbout = reader("HearAbout").ToString()
                PossibleRecord.delivery = reader("DeliveryOption").ToString()
                PossibleRecord.PointOfSale = reader("PointOfSales").ToString()
                PossibleRecord.Notes = reader("Notes").ToString()
                PossibleRecord.GC_Authorization = reader("Authorization").ToString()
                PossibleRecord.GC_Username = reader("UserName").ToString()
                PossibleRecord.ID = reader("ID").ToString()
                PossibleRecord.GC_Status = reader("Status")
                PossibleRecord.GC_Number = reader("Online_Certificate_Number").ToString
                PossibleRecord.PaymentMethod = reader("PaymentMethod").ToString



                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("LoadMatchData:", ex)
            ' LogError("LoadMatchData: sfield=", sfield)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles

    End Function



    Public Function RetrieveGiftCertificatesFromQueue(name As String, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)
            Dim sqlstring = String.Format("Select * from dbo.GiftCertificates where Purchaser_FirstName like  '%{0}%' OR Purchaser_LastName like  '%{1}%'", name.Trim, name.Trim)

            If devmode = False Then
                cmd.CommandText = sqlstring
            Else
                '//Devmode
                cmd.CommandText = "Select * from dbo.GiftCertificates"
            End If

            'cmd.CommandText = "Select * from dbo.GiftCertificates"
            cmd.CommandType = CommandType.Text
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim PossibleRecord As New ClsGiftCertificate


                PossibleRecord.ID = reader("Id").ToString()

                PossibleRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                PossibleRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                PossibleRecord.Purchaser_Address1 = reader("Purchaser_Address1").ToString()
                PossibleRecord.Purchaser_Address2 = reader("Purchaser_Address2").ToString()
                PossibleRecord.Purchaser_City = reader("Purchaser_City").ToString()
                PossibleRecord.Purchaser_State = reader("Purchaser_State").ToString()
                PossibleRecord.Purchaser_Zip = reader("Purchaser_Zip").ToString()
                PossibleRecord.Purchaser_Phones1 = reader("Purchaser_Phone1").ToString()
                PossibleRecord.Purchaser_Phones2 = reader("Purchaser_Phone2").ToString()
                PossibleRecord.Purchaser_Email = reader("Purchaser_Email").ToString()

                PossibleRecord.Recipient_FirstName = reader("Recipient_FirstName").ToString()
                PossibleRecord.Recipient_LastName = reader("Recipient_LastName").ToString()
                PossibleRecord.Recipient_Address1 = reader("Recipient_Address1").ToString()
                PossibleRecord.Recipient_Address2 = reader("Recipient_Address2").ToString()
                PossibleRecord.Recipient_City = reader("Recipient_City").ToString()
                PossibleRecord.Recipient_State = reader("Recipient_State").ToString()
                PossibleRecord.Recipient_Zip = reader("Recipient_Zip").ToString()
                PossibleRecord.Recipient_Phones1 = reader("Recipient_Phone1").ToString()
                PossibleRecord.Recipient_Phones2 = reader("Recipient_Phone2").ToString()
                PossibleRecord.Recipient_Email = reader("Recipient_Email").ToString()
                PossibleRecord.Item_Tandem10k = reader("Item_Tandem10k").ToString()
                PossibleRecord.Item_Tandem12k = reader("Item_Tandem12k").ToString()
                PossibleRecord.Item_Video = reader("Item_Video").ToString()
                PossibleRecord.Item_Other = reader("Item_Other").ToString()
                PossibleRecord.Item_OtherAmount = reader("Item_OtherAmount").ToString()

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    PossibleRecord.JR_PurchaseID = reader("JR_PurchaseID")
                End If
                If IsDBNull(reader("JR_RecipientID")) = False Then
                    PossibleRecord.JR_RecipientID = reader("JR_RecipientID")
                End If
                PossibleRecord.GC_DateEntered = reader("DateEntered").ToString()
                PossibleRecord.HearAbout = reader("HearAbout").ToString()
                PossibleRecord.delivery = reader("DeliveryOption").ToString()
                PossibleRecord.PointOfSale = reader("PointOfSales").ToString()
                PossibleRecord.Notes = reader("Notes").ToString()
                PossibleRecord.GC_Authorization = reader("Authorization").ToString()
                PossibleRecord.GC_Username = reader("UserName").ToString()
                PossibleRecord.ID = reader("ID").ToString()
                PossibleRecord.GC_Status = reader("Status")
                PossibleRecord.GC_Number = reader("Online_Certificate_Number").ToString
                PossibleRecord.PaymentMethod = reader("PaymentMethod").ToString



                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("LoadMatchData:", ex)
            ' LogError("LoadMatchData: sfield=", sfield)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles

    End Function


    Public Function SearchGiftCertificates(lastname As String, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            Dim sqlstring = String.Format("Select * from dbo.GiftCertificates where  Purchaser_LastName like '{0}%'", lastname)

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sqlstring
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim PossibleRecord As New ClsGiftCertificate


                PossibleRecord.ID = reader("Id").ToString()

                PossibleRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                PossibleRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                PossibleRecord.Purchaser_Address1 = reader("Purchaser_Address1").ToString()
                PossibleRecord.Purchaser_Address2 = reader("Purchaser_Address2").ToString()
                PossibleRecord.Purchaser_City = reader("Purchaser_City").ToString()
                PossibleRecord.Purchaser_State = reader("Purchaser_State").ToString()
                PossibleRecord.Purchaser_Zip = reader("Purchaser_Zip").ToString()
                PossibleRecord.Purchaser_Phones1 = reader("Purchaser_Phone1").ToString()
                PossibleRecord.Purchaser_Phones2 = reader("Purchaser_Phone2").ToString()
                PossibleRecord.Purchaser_Email = reader("Purchaser_Email").ToString()

                PossibleRecord.Recipient_FirstName = reader("Recipient_FirstName").ToString()
                PossibleRecord.Recipient_LastName = reader("Recipient_LastName").ToString()
                PossibleRecord.Recipient_Address1 = reader("Recipient_Address1").ToString()
                PossibleRecord.Recipient_Address2 = reader("Recipient_Address2").ToString()
                PossibleRecord.Recipient_City = reader("Recipient_City").ToString()
                PossibleRecord.Recipient_State = reader("Recipient_State").ToString()
                PossibleRecord.Recipient_Zip = reader("Recipient_Zip").ToString()
                PossibleRecord.Recipient_Phones1 = reader("Recipient_Phone1").ToString()
                PossibleRecord.Recipient_Phones2 = reader("Recipient_Phone2").ToString()
                PossibleRecord.Recipient_Email = reader("Recipient_Email").ToString()
                PossibleRecord.Item_Tandem10k = reader("Item_Tandem10k").ToString()
                PossibleRecord.Item_Tandem12k = reader("Item_Tandem12k").ToString()
                PossibleRecord.Item_Video = reader("Item_Video").ToString()
                PossibleRecord.Item_Other = reader("Item_Other").ToString()
                PossibleRecord.Item_OtherAmount = reader("Item_OtherAmount").ToString()

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    PossibleRecord.JR_PurchaseID = reader("JR_PurchaseID")
                End If
                If IsDBNull(reader("JR_RecipientID")) = False Then
                    PossibleRecord.JR_RecipientID = reader("JR_RecipientID")
                End If
                PossibleRecord.GC_DateEntered = reader("DateEntered").ToString()
                PossibleRecord.HearAbout = reader("HearAbout").ToString()
                PossibleRecord.delivery = reader("DeliveryOption").ToString()
                PossibleRecord.PointOfSale = reader("PointOfSales").ToString()
                PossibleRecord.Notes = reader("Notes").ToString()
                PossibleRecord.GC_Authorization = reader("Authorization").ToString()
                PossibleRecord.GC_Username = reader("UserName").ToString()
                PossibleRecord.ID = reader("ID").ToString()
                PossibleRecord.GC_Status = reader("Status")
                PossibleRecord.GC_Number = reader("Online_Certificate_Number").ToString
                PossibleRecord.PaymentMethod = reader("PaymentMethod").ToString



                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("LoadMatchData:", ex)
            ' LogError("LoadMatchData: sfield=", sfield)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles

    End Function


    Public Function LoadMatchData(c As ClsPersonSearch) As List(Of JumpRunPossibles)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of JumpRunPossibles)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "dbo.sprw_GetPossibleExistingCustomersFromJumpRun"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = _sqlCon
            cmd.Parameters.Add("@FirstName", SqlDbType.NVarChar)
            cmd.Parameters("@FirstName").Value = c.FirstName & ""
            cmd.Parameters.Add("@LastName", SqlDbType.NVarChar)
            cmd.Parameters("@LastName").Value = c.LastName & ""
            cmd.Parameters.Add("@Email", SqlDbType.NVarChar)
            cmd.Parameters("@Email").Value = c.Email & ""
            cmd.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar)
            cmd.Parameters("@PhoneNumber").Value = "" ', 'c.P
            cmd.Parameters.Add("@PreferedName", SqlDbType.NVarChar)
            cmd.Parameters("@PreferedName").Value = String.Format("{0},{1}", c.LastName, c.FirstName)
            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read

                Dim PossibleRecord As New JumpRunPossibles
                sfield = "WCustid"

                PossibleRecord.wCustId = reader(0)
                sfield = "sCust"
                PossibleRecord.sCust = getString(reader(1))
                sfield = "sFirstName"
                PossibleRecord.sFirstName = getString(reader(2))
                If c.FirstName.ToLower = PossibleRecord.sFirstName.ToLower And String.IsNullOrEmpty(c.FirstName) = False Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If
                sfield = "sLastName"
                PossibleRecord.sLastName = getString(reader(3))
                If c.LastName.ToLower = PossibleRecord.sLastName.ToLower And String.IsNullOrEmpty(c.LastName) = False Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                sfield = "sEmail"
                PossibleRecord.sEmail = getString(reader(6))
                If String.IsNullOrEmpty(c.Email) = False AndAlso c.Email.ToLower = PossibleRecord.sEmail.ToLower Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                sfield = "sEmail2"
                PossibleRecord.sEmail2 = getString(reader(7))
                If String.IsNullOrEmpty(c.Email) = False AndAlso c.Email.ToLower = PossibleRecord.sEmail2.ToLower Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                sfield = "sPhone"
                PossibleRecord.sPhone1 = getString(reader(8))
                If String.IsNullOrEmpty(c.Phone) = False AndAlso c.Phone = PossibleRecord.sPhone1 Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                sfield = "sZip"
                PossibleRecord.sZip = getString(reader(15))
                If String.IsNullOrEmpty(c.Zip) = False AndAlso PossibleRecord.sZip.StartsWith(c.Zip) Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If




                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("LoadMatchData:", ex)
            ' LogError("LoadMatchData: sfield=", sfield)
        Finally
            _sqlCon.Close()
        End Try

        Dim matchingRecords = From i In LstPossibles
                              Where i.PercentageMatch >= 20
                              Order By i.PercentageMatch Descending
                              Select i

        Return matchingRecords.ToList
    End Function


    '
    Public Function InsertNewGiftCertRecord(c As ClsGiftCertificate) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.GC_InsertGiftCertificate"

            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("DateEntered", Now)
            sqlComm.Parameters.AddWithValue("HearAbout", 0)
            sqlComm.Parameters.AddWithValue("DeliveryOption", c.delivery)
            sqlComm.Parameters.AddWithValue("PointOfSales", c.PointOfSale)
            sqlComm.Parameters.AddWithValue("Authorization", c.GC_Authorization)
            sqlComm.Parameters.AddWithValue("UserName", c.GC_Username)
            sqlComm.Parameters.AddWithValue("Status", c.GC_Status)
            sqlComm.Parameters.AddWithValue("Purchaser_FirstName", c.Purchaser_FirstName)
            sqlComm.Parameters.AddWithValue("Purchaser_LastName", c.Purchaser_LastName)
            sqlComm.Parameters.AddWithValue("Purchaser_Address1", c.Purchaser_Address1)
            sqlComm.Parameters.AddWithValue("Purchaser_Address2", c.Purchaser_Address2)
            sqlComm.Parameters.AddWithValue("Purchaser_City", c.Purchaser_City)
            sqlComm.Parameters.AddWithValue("Purchaser_State", c.Purchaser_State)
            sqlComm.Parameters.AddWithValue("Purchaser_Zip", c.Purchaser_Zip)
            sqlComm.Parameters.AddWithValue("Purchaser_Phone1", c.Purchaser_Phones1)
            sqlComm.Parameters.AddWithValue("Purchaser_Phone2", c.Purchaser_Phones2)
            sqlComm.Parameters.AddWithValue("Purchaser_Email", c.Purchaser_Email)
            sqlComm.Parameters.AddWithValue("Recipient_FirstName", c.Recipient_FirstName)
            sqlComm.Parameters.AddWithValue("Recipient_LastName", c.Recipient_LastName)
            sqlComm.Parameters.AddWithValue("Recipient_Address1", c.Recipient_Address1)
            sqlComm.Parameters.AddWithValue("Recipient_Address2", c.Recipient_Address2)
            sqlComm.Parameters.AddWithValue("Recipient_City", c.Recipient_City)
            sqlComm.Parameters.AddWithValue("Recipient_State", c.Recipient_State)
            sqlComm.Parameters.AddWithValue("Recipient_Zip", c.Recipient_Zip)
            sqlComm.Parameters.AddWithValue("Recipient_Phone1", c.Recipient_Phones1)
            sqlComm.Parameters.AddWithValue("Recipient_Phone2", c.Recipient_Phones2)
            sqlComm.Parameters.AddWithValue("Recipient_Email", c.Recipient_Email)
            sqlComm.Parameters.AddWithValue("Item_Tandem10k", c.Item_Tandem10k)
            sqlComm.Parameters.AddWithValue("Item_Tandem12k", c.Item_Tandem12k)
            sqlComm.Parameters.AddWithValue("Item_Video", c.Item_Video)
            sqlComm.Parameters.AddWithValue("Item_Other", c.Item_Other)
            sqlComm.Parameters.AddWithValue("Item_OtherAmount", c.Item_OtherAmount)
            sqlComm.Parameters.AddWithValue("Item_CalculatedTotal", c.GC_CalculateTotal)

            sqlComm.Parameters.AddWithValue("JR_PurchaseID", c.JR_PurchaseID)
            sqlComm.Parameters.AddWithValue("JR_RecipientID", c.JR_RecipientID)
            '//TO REMOVE THESE CC Detail fields
            sqlComm.Parameters.AddWithValue("Payment_CreditCardNumber", "")
            sqlComm.Parameters.AddWithValue("Payment_CreditCardExpiry", Now)
            sqlComm.Parameters.AddWithValue("GC_Number", c.GC_Number)
            sqlComm.Parameters.AddWithValue("PaymentMethod", c.PaymentMethod)



            sqlComm.ExecuteNonQuery()
            SuccessState = True

            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        LogError("InsertNewRecord:", ex)
            '        LogError("PreferredName=" & c.PreferredName, "")
            '        LogError(String.Format("FirstName = {0}, LastName = {1}", c.FirstName, c.LastName), "")
        Finally
            _sqlCon = Nothing
        End Try
        Return SuccessState
    End Function


    Public Function WhatRadioIsSelected(ByVal grp As Panel) As String
        Dim rbtn As RadioButton
        Dim rbtnName As String = String.Empty
        Try
            Dim ctl As Control
            For Each ctl In grp.Controls
                If TypeOf ctl Is RadioButton Then
                    rbtn = DirectCast(ctl, RadioButton)
                    If rbtn.Checked Then
                        rbtnName = rbtn.Name
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            Dim stackframe As New Diagnostics.StackFrame(1)
            Throw New Exception("An error occurred in routine, '" & stackframe.GetMethod.ReflectedType.Name & "." & System.Reflection.MethodInfo.GetCurrentMethod.Name & "'." & Environment.NewLine & "  Message was: '" & ex.Message & "'")
        End Try
        Return rbtnName
    End Function

    <Extension>
    Public Function ToProperCase(x As String) As String
        If String.IsNullOrEmpty(x) = False Then
            Return StrConv(x, VbStrConv.ProperCase)
        Else
            Return ""
        End If
    End Function


    <Extension>
    Public Function SanitizeString(o As String, Optional ByVal allowspecificchars As Boolean = False) As String
        If String.IsNullOrEmpty(0) Then

            Return ""
        Else
            Dim sb As New System.Text.StringBuilder

            For Each c In o
                If Char.IsLetterOrDigit(c) Or c = " " Then
                    sb.Append(c)
                ElseIf allowspecificchars AndAlso c = "-" Then
                    sb.Append(c)
                ElseIf allowspecificchars AndAlso c = "'" Then
                    sb.Append(c)
                End If
            Next
            Return sb.ToString
        End If

        Return CStr(o) & ""



    End Function




    Private Function CalculatePreferredName(ByVal c1 As ClsJumpRunCustomer, Optional Editaction As Boolean = False) As ClsJumpRunCustomer
        Dim c As ClsJumpRunCustomer = c1
        Try
            Dim sqlCon1 = New SqlConnection(_strConn)
            Dim icountmatches As Integer = 0
            Dim actualfirstname = c1.FirstName.Trim
            Dim preferredname = c1.PreferredName.Trim

            If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

            Using cmd As New SqlClient.SqlCommand("dbo.MatchingCustomerNamesExactCount", sqlCon1)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.AddWithValue("FirstName", c.FirstName.Trim)
                cmd.Parameters.AddWithValue("LastName", c.LastName.Trim)
                cmd.Parameters.Add("RowCount", SqlDbType.Int).Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()
                icountmatches = Convert.ToInt32(cmd.Parameters("RowCount").Value)
            End Using

            If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

            '//THis is editing existing record - so no need to recalculate a new name
            If Editaction = True And icountmatches <= 1 Then
                Return c
            End If

            Dim s = Asc(c.FirstName.Trim.Last)
            If s >= 48 And s <= 57 Then
                'First Name ends in a Number - so lets just leave it alone rather than
                'trying to recalculate as most firstnames dont end in a number
            Else
                Dim SimilarFirstNames As New List(Of PreferredNameFirstNamePair)
                If icountmatches > 0 Then
                    Dim sqlComm1 As New SqlCommand()
                    sqlComm1.Connection = sqlCon1
                    sqlComm1.CommandText = "dbo.MatchingCustomerNames1"
                    sqlComm1.CommandType = CommandType.StoredProcedure
                    sqlComm1.Parameters.AddWithValue("FirstName", c.FirstName.Trim)
                    sqlComm1.Parameters.AddWithValue("LastName", c.LastName.Trim)

                    If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

                    Dim sqlReader As SqlDataReader = sqlComm1.ExecuteReader()
                    Dim icount As Integer = 0

                    'Get all the customer name matches
                    ' because there may be multiple record for John Smith or Tdm-John Smith


                    'If its an edit then use the customer ID and leave the preferred name the same
                    'If its new then find the next unique number

                    If sqlReader.HasRows Then
                        While (sqlReader.Read())
                            icount = icount + 1

                            Dim s1 = sqlReader.GetInt32(0)
                            Dim s2 = sqlReader.GetString(1)
                            Dim s3 = sqlReader.GetString(2)
                            'c, prefere,first

                            SimilarFirstNames.Add(New PreferredNameFirstNamePair(CInt(sqlReader.GetInt32(0)), sqlReader.GetString(1), sqlReader.GetString(2)))
                        End While
                    End If

                    If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

                    '//New
                    '//Determine the non matching firstname with number and adjust the preferredName

                    For i = 1 To 100
                        If SimilarFirstNames.Count > 0 Then
                            c.PreferredName = String.Format("{0},{1}{2}", c.LastName.Trim, actualfirstname.Trim, i.ToString)
                            If DoesIndexExist(c.PreferredName.Trim) = False Then
                                Exit For
                            End If
                        Else
                            If DoesIndexExist(c.PreferredName.Trim) = False Then
                                Exit For
                            End If
                        End If
                    Next
                Else

                End If

            End If

        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  LogError("CalculatePreferredName:", ex)
        End Try

        Return c
    End Function


    Public Function InsertNewCustomerRecord(c As ClsJumpRunCustomer, ByRef newID As Integer) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()
            If String.IsNullOrEmpty(c.PreferredName.Trim) Then
                c.PreferredName = String.Format("{0},{1} ", c.LastName.Trim.ToProperCase, c.FirstName.Trim.ToProperCase)
            End If

            Dim c1 = CalculatePreferredName(c)
            c.PreferredName = c1.PreferredName.ToProperCase
            c.FirstName = c1.FirstName.ToProperCase.Trim

            'Just leave the alpha characters and see if the last name is contained in
            'the Nickname (Preferredname) - if its not stop and error rather than continue on the update.
            If c.PreferredName.Trim.SanitizeString.ToLower.Contains(c.LastName.SanitizeString.ToLower.Trim) = False Then
                Throw New Exception("Last name is not included in the Nickname, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
            ElseIf String.IsNullOrEmpty(c.PreferredName.SanitizeString.Trim) Then
                Throw New Exception("Nickname does not include names, something weird happened.   Capture the scenario about what customer records we were using.  Nickname=" & c.PreferredName & " LastName=" & c.LastName)
            End If

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            '    
            sqlComm.CommandText = "dbo.GC_InsertNewUserCustomer"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("PreferredName", c.PreferredName.ToProperCase.Trim)
            sqlComm.Parameters.AddWithValue("Student ", 1)
            sqlComm.Parameters.AddWithValue("FirstName", c.FirstName.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("LastName", c.LastName.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("MI", "")
            sqlComm.Parameters.AddWithValue("Phone1", c.Phone1.Trim)
            sqlComm.Parameters.AddWithValue("Street1", c.Street1.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("Street2", c.Street2.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("City", c.City.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("State", c.State.ToProperCase.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("Zip", c.Zip.SanitizeString.Trim)
            sqlComm.Parameters.AddWithValue("Email", c.Email.ToProperCase.Trim)
            sqlComm.Parameters.AddWithValue("sOpInsert ", "GC")
            sqlComm.Parameters.Add("new_identity", SqlDbType.Int).Direction = ParameterDirection.Output


            sqlComm.ExecuteNonQuery()
            Dim id = Convert.ToInt32(sqlComm.Parameters("new_identity").Value)
            newID = id
            SuccessState = True

            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        LogError("InsertNewRecord:", ex)
            '        LogError("PreferredName=" & c.PreferredName, "")
            '        LogError(String.Format("FirstName = {0}, LastName = {1}", c.FirstName, c.LastName), "")
        Finally
            _sqlCon = Nothing
        End Try
        Return SuccessState
    End Function


    Function GetListOfGCCustomers() As IEnumerable(Of Integer)
        Dim LstCustomers As New List(Of Integer)

        GetConnectionString()
        Try
            '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


            Dim sqlCon = New SqlConnection(_strConn)
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm1 As New SqlCommand()
            sqlComm1.Connection = sqlCon
            sqlComm1.CommandText = "dbo.GC_GetCustomers"
            sqlComm1.CommandType = CommandType.StoredProcedure

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlReader As SqlDataReader = sqlComm1.ExecuteReader()
            Dim icount As Integer = 0

            'Get all the customer name matches
            ' because there may be multiple record for John Smith or Tdm-John Smith


            'If its an edit then use the customer ID and leave the preferred name the same
            'If its new then find the next unique number


            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    icount = icount + 1

                    Dim CId = sqlReader.GetInt32(0)

                    'c, prefere,first
                    LstCustomers.Add(CId)
                End While
            End If
            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        LogError("InsertNewRecord:", ex)
            '        LogError("PreferredName=" & c.PreferredName, "")
            '        LogError(String.Format("FirstName = {0}, LastName = {1}", c.FirstName, c.LastName), "")
        Finally
            _sqlCon = Nothing
        End Try
        Return LstCustomers.AsEnumerable
    End Function



    Sub DeleteCustomer(Id As String)
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()


            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.PeopleDelete"

            sqlComm.CommandType = CommandType.StoredProcedure


            sqlComm.Parameters.AddWithValue("wCustId_1", Id)
            sqlComm.Parameters.AddWithValue("wBillToId_2", Id)


            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("RemoveFromJumpRun:", ex)
        Finally

            sqlCon = Nothing
        End Try
    End Sub



    'Private Function CalculatePreferredName(ByVal c1 As JumpRunUpdate, Optional Editaction As Boolean = False) As JumpRunUpdate
    'Dim c As JumpRunUpdate = c1
    'Try
    '    Dim sqlCon1 = New SqlConnection(_strConn)
    '    Dim icountmatches As Integer = 0
    '    Dim actualfirstname = c1.FirstName
    '    Dim preferredname = c1.PreferredName

    '    If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

    '    Using cmd As New SqlClient.SqlCommand("dbo.MatchingCustomerNamesExactCount", sqlCon1)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Parameters.AddWithValue("FirstName", c.FirstName)
    '        cmd.Parameters.AddWithValue("LastName", c.LastName)
    '        cmd.Parameters.Add("RowCount", SqlDbType.Int).Direction = ParameterDirection.Output
    '        cmd.ExecuteNonQuery()
    '        icountmatches = Convert.ToInt32(cmd.Parameters("RowCount").Value)
    '    End Using

    '    If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

    '    '//THis is editing existing record - so no need to recalculate a new name
    '    If Editaction = True And icountmatches <= 1 Then
    '        Return c
    '    End If

    '    Dim s = Asc(c.FirstName.Trim.Last)
    '    If s >= 48 And s <= 57 Then
    '        'First Name ends in a Number - so lets just leave it alone rather than
    '        'trying to recalculate as most firstnames dont end in a number
    '    Else
    '        Dim SimilarFirstNames As New List(Of PreferredNameFirstNamePair)
    '        If icountmatches > 0 Then
    '            Dim sqlComm1 As New SqlCommand()
    '            sqlComm1.Connection = sqlCon1
    '            sqlComm1.CommandText = "dbo.MatchingCustomerNames"
    '            sqlComm1.CommandType = CommandType.StoredProcedure
    '            sqlComm1.Parameters.AddWithValue("FirstName", c.FirstName)
    '            sqlComm1.Parameters.AddWithValue("LastName", c.LastName)

    '            If sqlCon1.State = ConnectionState.Closed Then sqlCon1.Open()

    '            Dim sqlReader As SqlDataReader = sqlComm1.ExecuteReader()
    '            Dim icount As Integer = 0

    '            'Get all the customer name matches
    '            ' because there may be multiple record for John Smith or Tdm-John Smith


    '            'If its an edit then use the customer ID and leave the preferred name the same
    '            'If its new then find the next unique number

    '            If sqlReader.HasRows Then
    '                While (sqlReader.Read())
    '                    icount = icount + 1
    '                    SimilarFirstNames.Add(New PreferredNameFirstNamePair(CInt(sqlReader.GetInt32(0)), sqlReader.GetString(39), sqlReader.GetString(1)))
    '                End While
    '            End If

    '            If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

    '            If Editaction = True Then
    '                '//Edit
    '                If SimilarFirstNames.Count > 0 Then
    '                    Dim x = (From i1 In SimilarFirstNames Where i1.scust = c1.wCustId Select i1).First
    '                    If x IsNot Nothing Then
    '                        c.PreferredName = x.preferedname
    '                    Else
    '                        LogError("CalculatePreferredName:", "No records matching with customer ID in SimilarFirstNames (1)")
    '                    End If
    '                Else
    '                    LogError("CalculatePreferredName:", "No records in SimilarFirstNames (2)")
    '                End If


    '            Else
    '                '//New
    '                '//Determine the non matching firstname with number and adjust the preferredName

    '                For i = 1 To 100
    '                    If SimilarFirstNames.Count > 0 Then
    '                        c.PreferredName = String.Format("{0},{1}{2}", c.LastName, actualfirstname, i.ToString)

    '                        If DoesIndexExist(c.PreferredName) = False Then
    '                            Exit For
    '                        End If

    '                    Else
    '                        If DoesIndexExist(c.PreferredName) = False Then
    '                            Exit For
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If

    '    End If

    'Catch ex As Exception
    '    MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    LogError("CalculatePreferredName:", ex)
    'End Try

    'Return c
    'End Function

    '//Verifies if the People table contains this string which is a unique Key
    Public Function DoesIndexExist(customerKeyString As String) As Boolean
        GetConnectionString()
        Dim sqlCommand As New SqlCommand

        Dim sqlconnect As SqlConnection = Nothing

        Try
            sqlconnect = New SqlConnection(_strConn)
            sqlconnect.Open()
            sqlCommand.CommandText = String.Format("Select * from tPeople WHERE sCust = '{0}'", customerKeyString)
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.Connection = sqlconnect

            Dim ds = New DataSet()
            Dim adap = New SqlDataAdapter(sqlCommand)
            adap.Fill(ds, "People")

            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("DoesIndexExist:", ex)
        Finally
            sqlconnect.Close()

        End Try
        Return False
    End Function

    Public Sub DeleteImportedGiftCertificates()
        'Set status to completed/processed
        GetConnectionString()
        Dim sqlCon = New SqlConnection(_strConn)
        Try

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "GC_DeleteImportGiftCertificates"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("UpdateCertificateStats:", ex)
        Finally

            sqlCon = Nothing
        End Try

    End Sub


    Public Function DoesGCExist(GCNumber As String) As Boolean
        GetConnectionString()
        Dim sqlCommand As New SqlCommand

        Dim sqlconnect As SqlConnection = Nothing

        Try
            sqlconnect = New SqlConnection(_strConn)
            sqlconnect.Open()
            Dim s = "SELECT * from dbo.GiftCertificates where Online_Certificate_Number = '" + GCNumber + "'"
            sqlCommand.CommandText = s
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.Connection = sqlconnect

            Dim ds = New DataSet()
            Dim adap = New SqlDataAdapter(sqlCommand)
            adap.Fill(ds, "Certs")

            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("DoesIndexExist:", ex)
        Finally
            sqlconnect.Close()

        End Try
        Return False
    End Function
End Module
