Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events

Public Class FrmProcess
    Public Property CurrentGiftCertificate As ClsGiftCertificate
    Public Property CurrentFilter As FilterStates


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim sp As String = ""
        Dim sr As String = ""

        If CurrentGiftCertificate.JR_PurchaseID = 0 Then
            sp = "[Create New]"
            Dim c1 As New ClsJumpRunCustomer
            c1.FirstName = CurrentGiftCertificate.Purchaser_FirstName
            c1.LastName = CurrentGiftCertificate.Purchaser_LastName
            c1.Street1 = CurrentGiftCertificate.Purchaser_Address1
            c1.Street2 = CurrentGiftCertificate.Purchaser_Address2
            c1.City = CurrentGiftCertificate.Purchaser_City
            c1.State = CurrentGiftCertificate.Purchaser_State
            c1.Zip = CurrentGiftCertificate.Purchaser_Zip
            c1.Phone1 = CurrentGiftCertificate.Purchaser_Phones1
            c1.Email = CurrentGiftCertificate.Purchaser_Email

            Dim NewRef As Integer = 0
            Dim NewId = InsertNewCustomerRecord(c1, NewRef)
            CurrentGiftCertificate.JR_PurchaseID = NewRef
        Else
            sp = CurrentGiftCertificate.JR_PurchaseID.ToString
        End If


        'If No recipient is specified then recipient = same as distination
        If CheckBox1.Checked Then
            sr = "[Same as Purchaser]"
        Else
            If CurrentGiftCertificate.JR_RecipientID = 0 Then
                sr = "[Create New]"
                Dim c2 As New ClsJumpRunCustomer
                c2.FirstName = CurrentGiftCertificate.Recipient_FirstName
                c2.LastName = CurrentGiftCertificate.Recipient_LastName
                c2.Street1 = CurrentGiftCertificate.Recipient_Address1
                c2.Street2 = CurrentGiftCertificate.Recipient_Address2
                c2.City = CurrentGiftCertificate.Recipient_City
                c2.State = CurrentGiftCertificate.Recipient_State
                c2.Zip = CurrentGiftCertificate.Recipient_Zip
                c2.Phone1 = CurrentGiftCertificate.Recipient_Phones1
                c2.Email = CurrentGiftCertificate.Recipient_Email

                Dim NewRef2 As Integer = 0
                Dim NewId = InsertNewCustomerRecord(c2, NewRef2)
                CurrentGiftCertificate.JR_RecipientID = NewRef2
            Else
                sr = CurrentGiftCertificate.JR_RecipientID.ToString
            End If

        End If






        Dim ActionString = String.Format("Purchase Record in JRUN {0},   Recipient Record in JRUN {1}", sp, sr)

        MsgBox("Based upon the recipient/purchaser options chosen create new customers/Payment records in Jumprun" & Environment.NewLine & ActionString)

        UpdateCertificateJumpRunCustomers(CurrentGiftCertificate, CurrentGiftCertificate.JR_PurchaseID, CurrentGiftCertificate.JR_RecipientID)
        UpdateCertificateStatus(CurrentGiftCertificate, CertificateStatus.Processed)

        Dim x As New FrmProcessPrint
        x.Certificate = CurrentGiftCertificate

        x.ShowDialog()

        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("Mark as create new Jumprun Customer")
        CurrentGiftCertificate.JR_PurchaseID = 0
        LblJumprunCustomerStatusPurchaser.Text = "Create a new customer in Jumprun"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Select an existing Jumprun Customer")
        Dim jrsel = CType(SfDGPurchaser.SelectedItem, JumpRunPossibles)
        If jrsel IsNot Nothing Then
            CurrentGiftCertificate.JR_PurchaseID = jrsel.wCustId
        Else
            LblJumprunCustomerStatusPurchaser.Text = ""
        End If

        LblJumprunCustomerStatusPurchaser.Text = "Use selected customer in Jumprun " & CurrentGiftCertificate.JR_PurchaseID.ToString
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try


            'GetCertificatesToProcess(dteEntry, SfDataGrid1)

            Try

                '//Which was checked
                Dim s1 = WhatRadioIsSelected(Me.Panel3)
                Select Case s1.ToLower

                    Case "rdoentered"
                        CurrentFilter = FilterStates.Entered
                    Case "rdoprocessed"
                        CurrentFilter = FilterStates.Processed
                    Case "rdocompleted"
                        CurrentFilter = FilterStates.Completed
                    Case "rdoall"
                        CurrentFilter = FilterStates.All

                End Select
                Dim dteEntry As Date = SfDateEntry.Value
                Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
                SfDataGrid1.DataSource = fl
            Catch ex As Exception
                MsgBox(ex.Message)

            End Try



        Catch ex As Exception

        End Try

    End Sub

    Private Sub SfListView1_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub FrmProcess_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetupForm()
        Dim dteEntry As Date = SfDateEntry.Value
        'GetCertificatesToProcess(dteEntry, SfDataGrid1)
        'Dim gclist As New List(Of ClsGiftCertificate)
        'gclist = RetrieveGiftCertificatesFromQueue(Now.Date)



        'Dim gclist1 As New List(Of ClsGiftCertificate)
        'gclist1 = RetrieveGiftCertificatesFromQueue(New Date(2020, 1, 1))
        ''gclist1 = RetrieveGiftCertificatesFromQueue(SfDateEntry.Value)
        'SfDataGrid1.DataSource = gclist1

        ''Dim gclist2 As New List(Of ClsGiftCertificate)
        ''gclist2 = RetrieveGiftCertificatesFromQueue(New Date(2021, 1, 1))
    End Sub



    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate)
                PopulateCurrentGiftCertificateDetails(CurrentGiftCertificate)
                ' Button1.Enabled = True
            Else
                BlankDisplayFields()
                Button1.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("SfDataGrid1_SelectionChanged:", ex)
        End Try
    End Sub

    Private Sub BlankDisplayFields()
        'Used to populate a form from the records        
        SetTextField(lblPurchaser_FirstName, "")
        SetTextField(lblPurchaser_LastName, "")
        SetTextField(lblPurchaser_Address1, "")
        SetTextField(lblPurchaser_Address2, "")
        SetTextField(lblPurchaser_Address1, "")
        SetTextField(lblPurchaser_Address2, "")
        SetTextField(lblPurchaser_City, "")
        SetTextField(lblPurchaser_State, "")
        SetTextField(lblPurchaser_Zip, "")
        SetTextField(lblPurchaser_Phone1, "")
        SetTextField(lblPurchaser_Phone2, "")
        SetTextField(lblPurchaser_Email, "")


        'Recipient Fields
        SetTextField(Me.lblRecipient_FirstName, "")
        SetTextField(Me.lblRecipient_LastName, "")
        SetTextField(Me.lblRecipient_Address1, "")
        SetTextField(Me.lblRecipient_Address2, "")
        SetTextField(Me.lblRecipient_City, "")
        SetTextField(Me.lblRecipient_State, "")
        SetTextField(Me.lblRecipient_Zip, "")
        SetTextField(Me.lblRecipient_Phone1, "")
        SetTextField(Me.lblRecipient_Phone2, "")
        SetTextField(Me.lblRecipient_Email, "")

        SetTextField(Me.lblAuthorization, "")
        SetTextField(Me.LblDateEntered, "")

        'Item Fields
        Label51.Text = ""
        Label52.Text = ""
        Label53.Text = ""
        CheckBox2.Checked = False
        Label54.Text = ""
        Label47.Text = ""
        Label27.Text = ""

        lblDeliveryOption.Text = ""
        lblPointOfSale.Text = ""
        lblPaymentMethod.Text = ""

        Panel2.Visible = False

        ''IS recipient even specified if not then 
        'If String.IsNullOrEmpty(lblRecipient_FirstName.Text) AndAlso
        '        String.IsNullOrEmpty(lblRecipient_LastName.Text) AndAlso
        '        String.IsNullOrEmpty(lblRecipient_City.Text) AndAlso
        '        String.IsNullOrEmpty(lblRecipient_Zip.Text) AndAlso
        '        String.IsNullOrEmpty(lblRecipient_Phone1.Text) AndAlso
        '        String.IsNullOrEmpty(lblRecipient_Email.Text) Then
        '    CheckBox1.Enabled = False
        '    CheckBox1.Checked = True
        '    Panel2.Visible = False
        'ElseIf lblRecipient_FirstName.Text.Trim = lblPurchaser_FirstName.Text.Trim AndAlso
        '        lblRecipient_LastName.Text.Trim = lblPurchaser_LastName.Text.Trim Then
        '    CheckBox1.Enabled = False
        '    CheckBox1.Checked = True
        'Else
        '    CheckBox1.Enabled = False
        '    CheckBox1.Checked = False
        '    Panel2.Visible = True
        'End If
        ''Get Potential Matches for name
        'Dim objperson1 = New ClsPersonSearch() With {.FirstName = lblPurchaser_FirstName.Text.Trim, .LastName = lblPurchaser_LastName.Text.Trim, .Email = "", .Zip = lblPurchaser_Zip.Text.Trim}
        'Dim x1 = LoadMatchData(objperson1)

        '' Dim objperson2 = New ClsPersonSearch() With {.FirstName = "Brandon", .LastName = "Parrish", .Email = "Parrishb@Live.Com"}


        'SfDGPurchaser.DataSource = x1
        'If x1.Count = 0 Then
        '    If CurrentGiftCertificate IsNot Nothing Then
        '        CurrentGiftCertificate.JR_PurchaseID = 0
        '        Button1.Enabled = False
        '        LblJumprunCustomerStatusPurchaser.Text = "Create a new customer in Jumprun"
        '        Button2.Enabled = True
        '    Else
        '        '//They can add or select
        '        Button1.Enabled = True
        '        Button2.Enabled = True
        '    End If
        'Else
        '    Button1.Enabled = True
        '    Button2.Enabled = True
        'End If

        'Dim objperson2 = New ClsPersonSearch() With {.FirstName = lblRecipient_FirstName.Text.Trim, .LastName = lblRecipient_LastName.Text.Trim, .Email = lblRecipient_Email.Text.Trim}

        'Dim x2 = LoadMatchData(objperson2)
        'SfDGRecipient.DataSource = x2
        'If x2.Count = 0 Then
        '    If CurrentGiftCertificate IsNot Nothing Then
        '        CurrentGiftCertificate.JR_RecipientID = 0
        '        Button4.Enabled = False
        '        LblJumprunCustomerStatusRecipient.Text = "Create a new customer in Jumprun"
        '        Button3.Enabled = True
        '    Else
        '        '//They can add or select
        '        Button4.Enabled = True
        '        Button3.Enabled = True
        '    End If
        'Else
        '    Button4.Enabled = True
        '    Button3.Enabled = True
        'End If


        'If CurrentGiftCertificate.GC_Status = CertificateStatus.Entered Then
        '    'TODO:
        '    'Enable Process Button
        '    'Enable The Create Buttons

        '    If x1.Count > 0 Then
        '        'if an item is selected in the datagrid
        '        '   Enable a select button
        '        'end if
        '    End If
        '    If x2.Count > 0 Then
        '        'if an item is selected in the datagrid
        '        '   Enable a select button
        '        'end if
        '    End If

        'End If
    End Sub

    Sub SetupForm()
        Me.SfDateEntry.Value = Now.Date
        Me.CheckBox2.Enabled = False

        'Populate fields from LIst which I want to show - Purchaser, Recipient, Shipper 
        SfDataGrid1.TableControl.VerticalScrollBarVisible = True


        SfDataGrid1.AutoGenerateColumns = False
        SfDataGrid1.AllowResizingColumns = True
        SfDataGrid1.Columns.Clear()
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "ID", .HeaderText = "Id"})
        'SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_FirstName", .HeaderText = "Purchase First"})
        'SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_LastName", .HeaderText = "Purchase Last"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_Name", .HeaderText = "Purchaser Name"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Recipient_Name", .HeaderText = "Recipient Name"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_Status", .HeaderText = "Status"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "PointOfSale", .HeaderText = "Point Of Sale"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_Number", .HeaderText = "Online Certificate Number"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item_Tandem10k", .HeaderText = "Tandem 10k"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item_Tandem12k", .HeaderText = "Tandem 12k"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item_Video", .HeaderText = "Video"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item_OtherAmount", .HeaderText = "Other"})
        'SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item_OtherAmount", .HeaderText = "Other"})

        'Setup the Jumprun Datagrid columns (for bothg grids)
        SfDGPurchaser.AutoGenerateColumns = False
        SfDGPurchaser.AllowResizingColumns = True
        SfDGPurchaser.Columns.Clear()
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "wCustId", .HeaderText = "Id"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "PercentageMatch", .HeaderText = "Score"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sCust", .HeaderText = "Name"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sStreet1", .HeaderText = "Address"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sCity", .HeaderText = "City"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sState", .HeaderText = "State"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sZip", .HeaderText = "Zip"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sEmail", .HeaderText = "Email"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone1", .HeaderText = "Phone 1"})
        SfDGPurchaser.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone2", .HeaderText = "Phone 2"})

        SfDGRecipient.AutoGenerateColumns = False
        SfDGRecipient.AllowResizingColumns = True
        SfDGRecipient.Columns.Clear()
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "wCustId", .HeaderText = "Id"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "PercentageMatch", .HeaderText = "Score"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sCust", .HeaderText = "Name"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sStreet1", .HeaderText = "Address"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sCity", .HeaderText = "City"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sState", .HeaderText = "State"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sZip", .HeaderText = "Zip"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sEmail", .HeaderText = "Email"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone1", .HeaderText = "Phone 1"})
        SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone2", .HeaderText = "Phone 2"})


        'Status of buttons will be disabled by default
        BlankDisplayFields()
    End Sub

    Sub PopulateCurrentGiftCertificateDetails(gc As ClsGiftCertificate)
        'Used to populate a form from the records        
        SetTextField(lblPurchaser_FirstName, gc.Purchaser_FirstName)
        SetTextField(lblPurchaser_LastName, gc.Purchaser_LastName)
        SetTextField(lblPurchaser_Address1, gc.Purchaser_Address1)
        SetTextField(lblPurchaser_Address2, gc.Purchaser_Address2)
        SetTextField(lblPurchaser_Address1, gc.Purchaser_Address1)
        SetTextField(lblPurchaser_Address2, gc.Purchaser_Address2)
        SetTextField(lblPurchaser_City, gc.Purchaser_City)
        SetTextField(lblPurchaser_State, gc.Purchaser_State)
        SetTextField(lblPurchaser_Zip, gc.Purchaser_Zip)
        SetTextField(lblPurchaser_Phone1, gc.Purchaser_Phones1)
        SetTextField(lblPurchaser_Phone2, gc.Purchaser_Phones2)
        SetTextField(lblPurchaser_Email, gc.Purchaser_Email)


        'Recipient Fields
        SetTextField(Me.lblRecipient_FirstName, gc.Recipient_FirstName)
        SetTextField(Me.lblRecipient_LastName, gc.Recipient_LastName)
        SetTextField(Me.lblRecipient_Address1, gc.Recipient_Address1)
        SetTextField(Me.lblRecipient_Address2, gc.Recipient_Address2)
        SetTextField(Me.lblRecipient_City, gc.Recipient_City)
        SetTextField(Me.lblRecipient_State, gc.Recipient_State)
        SetTextField(Me.lblRecipient_Zip, gc.Recipient_Zip)
        SetTextField(Me.lblRecipient_Phone1, gc.Recipient_Phones1)
        SetTextField(Me.lblRecipient_Phone2, gc.Recipient_Phones2)
        SetTextField(Me.lblRecipient_Email, gc.Recipient_Email)

        SetTextField(Me.lblAuthorization, gc.GC_Authorization)
        SetTextField(Me.LblDateEntered, gc.GC_DateEntered.ToShortDateString)

        'Item Fields
        Label51.Text = gc.Item_Tandem10k.ToString
        Label52.Text = gc.Item_Tandem12k.ToString
        Label53.Text = gc.Item_Video.ToString
        CheckBox2.Checked = gc.Item_Other
        Label54.Text = gc.Item_OtherAmount.ToString
        Label47.Text = gc.GC_CalculateTotal.ToString '//This willbe the calculated from the line items

        Dim DeliveryOption As String = ""
        Select Case gc.delivery
            Case DeliveryOptions.Email
                DeliveryOption = "Email"
            Case DeliveryOptions.USMail
                DeliveryOption = "US Mail"
            Case DeliveryOptions.USMailDiscrete
                DeliveryOption = "US Mail Discrete"
            Case DeliveryOptions.InOffice
                DeliveryOption = "In Office"
        End Select
        lblDeliveryOption.Text = DeliveryOption

        Dim POS As String = ""
        Select Case gc.PointOfSale
            Case PointOfSale.Online
                POS = "Online"
            Case PointOfSale.PhoneInPerson
                POS = "In Person"
        End Select
        lblPointOfSale.Text = POS


        Dim PM As String = ""
        Select Case gc.PaymentMethod
            Case PaymentMethod.Cash
                PM = "Cash"
            Case PaymentMethod.CreditCard
                PM = "Credit Card"

            Case PaymentMethod.Online
                PM = "Online"
        End Select
        lblPaymentMethod.Text = PM


        'IS recipient even specified if not then 
        If String.IsNullOrEmpty(lblRecipient_FirstName.Text) AndAlso
                String.IsNullOrEmpty(lblRecipient_LastName.Text) AndAlso
                String.IsNullOrEmpty(lblRecipient_City.Text) AndAlso
                String.IsNullOrEmpty(lblRecipient_Zip.Text) AndAlso
                String.IsNullOrEmpty(lblRecipient_Phone1.Text) AndAlso
                String.IsNullOrEmpty(lblRecipient_Email.Text) Then
            CheckBox1.Enabled = False
            CheckBox1.Checked = True
            Panel2.Visible = False
        ElseIf lblRecipient_FirstName.Text.Trim = lblPurchaser_FirstName.Text.Trim AndAlso
                lblRecipient_LastName.Text.Trim = lblPurchaser_LastName.Text.Trim Then
            CheckBox1.Enabled = False
            CheckBox1.Checked = True
            Panel2.Visible = False

        Else
            CheckBox1.Enabled = False
            CheckBox1.Checked = False
            Panel2.Visible = True
        End If
        'Get Potential Matches for name
        Dim objperson1 = New ClsPersonSearch() With {.FirstName = lblPurchaser_FirstName.Text.Trim, .LastName = lblPurchaser_LastName.Text.Trim, .Email = "", .Zip = lblPurchaser_Zip.Text.Trim}
        Dim x1 = LoadMatchData(objperson1)

        ' Dim objperson2 = New ClsPersonSearch() With {.FirstName = "Brandon", .LastName = "Parrish", .Email = "Parrishb@Live.Com"}


        SfDGPurchaser.DataSource = x1
        If x1.Count = 0 Then
            If CurrentGiftCertificate IsNot Nothing Then
                CurrentGiftCertificate.JR_PurchaseID = 0
                Button1.Enabled = False
                LblJumprunCustomerStatusPurchaser.Text = "Create a new customer in Jumprun"
                Button2.Enabled = True
            Else
                '//They can add or select
                Button1.Enabled = True
                Button2.Enabled = True
            End If
        Else
            Button1.Enabled = True
            Button2.Enabled = True
        End If

        If gc.JR_PurchaseID > 0 Then
            If x1.Count > 0 Then
                SetJumprRunMatchInDataGrid(gc.JR_PurchaseID, SfDGPurchaser)
            End If
        End If


        Dim objperson2 = New ClsPersonSearch() With {.FirstName = lblRecipient_FirstName.Text.Trim, .LastName = lblRecipient_LastName.Text.Trim, .Email = lblRecipient_Email.Text.Trim}

        Dim x2 = LoadMatchData(objperson2)
        SfDGRecipient.DataSource = x2
        If x2.Count = 0 Then
            If CurrentGiftCertificate IsNot Nothing Then
                CurrentGiftCertificate.JR_RecipientID = 0
                Button4.Enabled = False
                LblJumprunCustomerStatusRecipient.Text = "Create a new customer in Jumprun"
                Button3.Enabled = True
            Else
                '//They can add or select
                Button4.Enabled = True
                Button3.Enabled = True
            End If
        Else
            Button4.Enabled = True
            Button3.Enabled = True
        End If
        If gc.JR_RecipientID > 0 Then
            If x1.Count > 0 Then
                SetJumprRunMatchInDataGrid(gc.JR_RecipientID, SfDGRecipient)
            End If
        End If

        If CurrentGiftCertificate.GC_Status = CertificateStatus.Entered Then
            'TODO:
            'Enable Process Button
            'Enable The Create Buttons

            If x1.Count > 0 Then
                'if an item is selected in the datagrid
                '   Enable a select button
                'end if
            End If
            If x2.Count > 0 Then
                'if an item is selected in the datagrid
                '   Enable a select button
                'end if
            End If

        Else
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False

        End If
    End Sub

    Private Sub SetJumprRunMatchInDataGrid(ID As Integer, sfgrid As SfDataGrid)
        Dim records = sfgrid.View.Records
        For Each record In records
            Dim obj = TryCast(record.Data, JumpRunPossibles)
            If obj.wCustId = ID Then
                sfgrid.SelectedItems.Add(obj)
            End If
        Next record
        'Set selected item to match id if it exists
    End Sub

    Private Sub SetTextField(control As Label, item As String)
        If String.IsNullOrEmpty(item) Then
            control.Text = ""
        Else
            control.Text = item
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CurrentGiftCertificate.JR_RecipientID = 1


        Dim jrsel = CType(SfDGRecipient.SelectedItem, JumpRunPossibles)
        If jrsel IsNot Nothing Then
            CurrentGiftCertificate.JR_RecipientID = jrsel.wCustId
            LblJumprunCustomerStatusRecipient.Text = "Use selected customer in Jumprun" & CurrentGiftCertificate.JR_RecipientID.ToString

        Else
            LblJumprunCustomerStatusRecipient.Text = ""
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CurrentGiftCertificate.JR_RecipientID = 0
        LblJumprunCustomerStatusRecipient.Text = "Create a new customer in Jumprun"
    End Sub

    Public Function GetCertificatesToProcess(entrydate As Date, Filter As FilterStates) As List(Of ClsGiftCertificate)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate)
            gclist1 = RetrieveGiftCertificatesFromQueue(entrydate)

            Dim FilteredList As List(Of ClsGiftCertificate)

            Select Case Filter
                Case FilterStates.Entered
                    CurrentFilter = FilterStates.Entered
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Select i1).ToList

                Case FilterStates.Processed
                    CurrentFilter = FilterStates.Processed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Processed Select i1).ToList



                Case FilterStates.Completed
                    CurrentFilter = FilterStates.Completed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Select i1).ToList

                Case FilterStates.ProcessAndCompleteOnly
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Or
                                                          i1.GC_Status = CertificateStatus.Processed
                                    Select i1).ToList
                Case FilterStates.All
                    FilteredList = (From i1 In gclist1 Select i1).ToList
            End Select

            Return FilteredList

        Catch ex As Exception

        End Try

    End Function

    Private Sub RdoEntered_CheckedChanged(sender As Object, e As EventArgs) Handles RdoEntered.CheckedChanged, RdoCompleted.CheckedChanged, RdoProcessed.CheckedChanged, RdoAll.CheckedChanged
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processed
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All

            End Select
            Dim dteEntry As Date = SfDateEntry.Value
            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            SfDataGrid1.DataSource = fl
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub SfDataGrid1_Click(sender As Object, e As EventArgs) Handles SfDataGrid1.Click

    End Sub
End Class