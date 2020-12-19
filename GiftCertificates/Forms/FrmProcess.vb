Imports System.ComponentModel
Imports System.Reflection
Imports Newtonsoft.Json
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Enums
Imports Syncfusion.WinForms.DataGrid.Events


Public Class FrmProcess


    Public Property CurrentGiftCertificate As ClsGiftCertificate2
    Public Property CurrentFilter As FilterStates
    Public Property DefaultDate As Nullable(Of DateTime) = Nothing
    Public Property JRCustomerSpecified As Boolean = False

    Public ObjPricing As ClsPromoPricing
    Public CurrentItemPricing As ClsPromoPricing


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim sp As String = ""
        Dim dteInsertDate As DateTime = Now
        Dim ExtraAltitudeDiscount As Boolean = False

        Try
            If GetBusinessDate() <> CurrentGiftCertificate.GC_DateEntered Then
                MessageBox.Show(String.Format("The entered date {0:d} does not match the current JumpRun business date {1:d}.   The transaction will be processed using the current jumprun business date.", CurrentGiftCertificate.GC_DateEntered, GetBusinessDate), "Process Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If


            If CurrentGiftCertificate.JR_PurchaserID = 0 Then
                sp = "<New Customer>"
            Else
                sp = CurrentGiftCertificate.JR_PurchaserID.ToString
            End If

            Dim StrAuthorizer = CboAuthorizer.Text

            If String.IsNullOrEmpty(StrAuthorizer) Then
                MessageBox.Show("Need to provider an Authorizer", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Else
                CurrentGiftCertificate.GC_Authorization = StrAuthorizer
            End If

            If JRCustomerSpecified = False Then
                MessageBox.Show("You need to provider an determine whether to link to existing JumpRun customer or create a new one", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If



            Dim ConfirmationString As String = ""
            If sp = "<New Customer>" Then
                ConfirmationString = String.Format("Based upon the options currently selected you have chosen to CREATE A NEW JumpRun Customer" & Environment.NewLine & Environment.NewLine & "DO you wish to continue ?", sp)


            Else
                ConfirmationString = String.Format("Based upon the options currently selected you have chosen to use an existing JumpRun Customer,    Customer ID {0}" & Environment.NewLine & Environment.NewLine & "DO YOU WISH TO CONTINUE ?", sp)
            End If

            If MessageBox.Show(ConfirmationString, "Confirm Jumprun Customer Match", MessageBoxButtons.YesNo) = DialogResult.No Then
                Exit Sub
            End If
            Dim PricingIsValid As Boolean = False

            CurrentItemPricing = CalculateCurrentOrderPricing(CurrentGiftCertificate, False)

            If CurrentGiftCertificate.GC_DiscountCode <> String.Empty Then
                '//We need to look at the discount code and set the pricing accordingly
                Dim code = CurrentGiftCertificate.GC_DiscountCode.ToUpper
                Dim intPromoPricing = 0
                Dim DiscountAmount = 0

                Select Case code
                    Case "SANTA20"
                        intPromoPricing = 5
                        DiscountAmount = -50
                    Case "ELF20"
                        intPromoPricing = 6
                        DiscountAmount = -35
                    Case "CC20"
                        intPromoPricing = 7
                        DiscountAmount = -25
                End Select
                ObjPricing = RetrievePricingPromoForId(intPromoPricing)
                ObjPricing.ItemPrice1 = ObjPricing.ItemPrice1
                ObjPricing.ItemDiscount1 = Math.Abs(DiscountAmount)
                ObjPricing.ItemPrice2 = ObjPricing.ItemPrice2
                ObjPricing.ItemDiscount2 = Math.Abs(DiscountAmount)
                ObjPricing.ItemDiscount3 = ObjPricing.ItemPrice3
                ObjPricing.ItemDiscount3 = Math.Abs(DiscountAmount)
                ObjPricing.ItemDiscount4 = ObjPricing.ItemPrice4
                ObjPricing.ItemDiscount4 = Math.Abs(DiscountAmount)

                Dim FFDiscountAmount As Double

                'DISC-FREE EXTRA AL
                Dim xfd As Double
                If My.Settings.FreeAltitudeID <> 0 Then
                    xfd = GetJumpRunItemPrice(My.Settings.FreeAltitudeID)
                Else
                    xfd = 0
                End If

                '//Lets double check extra freefall on top of this
                Dim x2a = ValidatePricingDifferenceBetweenWebstoreAndPromo(CurrentGiftCertificate, CurrentItemPricing, ObjPricing)
                If x2a = False Then
                    Select Case code
                        Case "SANTA20"
                            intPromoPricing = 8
                            DiscountAmount = -50
                            FFDiscountAmount = -50 - Math.Abs(xfd)
                        Case "ELF20"
                            intPromoPricing = 9
                            DiscountAmount = -35
                            FFDiscountAmount = -35 - Math.Abs(xfd)
                        Case "CC20"
                            intPromoPricing = 10
                            DiscountAmount = -25
                            FFDiscountAmount = -25 - Math.Abs(xfd)
                    End Select
                    ObjPricing = RetrievePricingPromoForId(intPromoPricing)
                    ObjPricing.ItemPrice1 = ObjPricing.ItemPrice1
                    ObjPricing.ItemDiscount1 = Math.Abs(DiscountAmount)
                    ObjPricing.ItemPrice2 = ObjPricing.ItemPrice2
                    ObjPricing.ItemDiscount2 = Math.Abs(FFDiscountAmount)
                    ObjPricing.ItemDiscount3 = ObjPricing.ItemPrice3
                    ObjPricing.ItemDiscount3 = Math.Abs(DiscountAmount)
                    ObjPricing.ItemDiscount4 = ObjPricing.ItemPrice4
                    ObjPricing.ItemDiscount4 = Math.Abs(FFDiscountAmount)
                    Dim x2b = ValidatePricingDifferenceBetweenWebstoreAndPromo(CurrentGiftCertificate, CurrentItemPricing, ObjPricing)
                    If x2b = False Then
                        MsgBox("Problem with pricing - stopping")

                        Exit Sub
                    Else
                        ExtraAltitudeDiscount = True
                        PricingIsValid = True

                    End If
                Else
                    PricingIsValid = True

                End If
            Else
                Dim intPromoPricing = CType(CboPromoPricing.SelectedItem, KeyValuePair(Of Integer, String)).Key
                ObjPricing = RetrievePricingPromoForId(intPromoPricing)

                Dim x2 = ValidatePricingDifferenceBetweenWebstoreAndPromo(CurrentGiftCertificate, CurrentItemPricing, ObjPricing)
                If x2 = False Then


                    'DISC-FREE EXTRA AL
                    Dim xfd1 As Double
                    If My.Settings.FreeAltitudeID <> 0 Then
                        xfd1 = GetJumpRunItemPrice(My.Settings.FreeAltitudeID)
                    Else
                        xfd1 = 0
                    End If
                    ObjPricing.ItemPrice1 = ObjPricing.ItemPrice1
                    ObjPricing.ItemDiscount1 = 0
                    ObjPricing.ItemPrice2 = ObjPricing.ItemPrice2
                    ObjPricing.ItemDiscount2 = Math.Abs(xfd1)
                    ObjPricing.ItemDiscount3 = ObjPricing.ItemPrice3
                    ObjPricing.ItemDiscount3 = 0
                    ObjPricing.ItemDiscount4 = ObjPricing.ItemPrice4
                    ObjPricing.ItemDiscount4 = Math.Abs(xfd1)
                    Dim x2a = ValidatePricingDifferenceBetweenWebstoreAndPromo(CurrentGiftCertificate, CurrentItemPricing, ObjPricing)
                    If x2a = False Then
                        PricingIsValid = False
                    Else

                        If MessageBox.Show("It looks Like free extra altitude may have been in play at the time the order was taken.  Adding it will balance the transaction.   Do you wish to continue ?", "Potential Missing Discount", MessageBoxButtons.YesNo) = DialogResult.No Then
                            PricingIsValid = False
                            Exit Sub
                        Else
                            ExtraAltitudeDiscount = True
                            PricingIsValid = True
                        End If


                    End If
                Else
                    PricingIsValid = True
                End If


            End If
            '//If discount code then we need to do some special processing
            'else
            '  prcing should be validated and correct
            If PricingIsValid Then
                MsgBox("Pricing Looks Goood. We can continue.")
            Else
                MsgBox("Problem with pricing - If discount code Is specified standard pricing should be selected, if no discount code is selected then select appropriate promo and try again")
                Exit Sub
            End If



            Dim difference = ValidateIndividualCertificateRecordsFromGCOrderWithPayment(CurrentGiftCertificate, CType(CboPromoPricing.SelectedItem, KeyValuePair(Of Integer, String)).Key)

            If CurrentGiftCertificate.JR_PurchaserID = 0 Then
                sp = "[Create New]"
                Dim c1 As New ClsJumpRunCustomer
                c1.FirstName = CurrentGiftCertificate.Purchaser_FirstName
                c1.LastName = CurrentGiftCertificate.Purchaser_LastName
                c1.Street1 = CurrentGiftCertificate.Billing_Address.Address1
                c1.Street2 = CurrentGiftCertificate.Billing_Address.Address2
                c1.City = CurrentGiftCertificate.Billing_Address.City
                c1.State = CurrentGiftCertificate.Billing_Address.State
                c1.Zip = CurrentGiftCertificate.Billing_Address.Zip
                c1.Phone1 = CurrentGiftCertificate.Billing_Address.Phone1
                c1.Email = CurrentGiftCertificate.Billing_Address.Email
                c1.dtInsert = dteInsertDate
                c1.sOpInsert = sOpInsertUser
                Dim NewRef As Integer = 0
                Dim NewId = InsertNewCustomerRecord(c1, NewRef)
                CurrentGiftCertificate.JR_PurchaserID = NewRef
            Else
                sp = CurrentGiftCertificate.JR_PurchaserID.ToString
            End If

            If CboAuthorizer.SelectedItem IsNot Nothing Then
                Dim authID = CInt(CType(CboAuthorizer.SelectedItem, KeyValuePair(Of Integer, String)).Key)
                StrAuthorizer = RetrieveAuthorizerCode(authID)
            Else
                StrAuthorizer = "SS"
            End If

            UpdateGCOrderAuthorizer(CurrentGiftCertificate, StrAuthorizer, TextBox3.Text.Trim)
            UpdateCertificateJumpRunCustomers(CurrentGiftCertificate, CurrentGiftCertificate.JR_PurchaserID)


            GenerateIndividualCertificateRecordsFromGCOrder(CurrentGiftCertificate, ObjPricing, ExtraAltitudeDiscrepency:=ExtraAltitudeDiscount)


            UpdateGCOrderStatus(CurrentGiftCertificate, CertificateStatus.Processing, dteInsertDate)

            Dim output = JsonConvert.SerializeObject(CurrentGiftCertificate)
            SendProcessDetails(output)


            Dim x As New FrmProcessPrint
            x.Certificate = CurrentGiftCertificate

            x.ShowDialog()

            If RadioButton1.Checked Then
                Button7.PerformClick()
            Else
                Button8.PerformClick()
            End If


            '//TODO: RESET THE CREATE OR SELECT CUSTOMER BACK TO DEFAULT

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function ValidatePricingDifferenceBetweenWebstoreAndPromo(GC As ClsGiftCertificate2, currentItemPricing As ClsPromoPricing, objPricing As ClsPromoPricing) As Object
        Dim Valid As Boolean = True

        'For each item with a qty verify that the ObjPricing item price matches the currentItemPricing price + discount
        If GC.Item1.Quantity > 0 Then
            If currentItemPricing.TotalPrice1 <> objPricing.TotalPrice1 Then
                Valid = False
            End If
        End If
        If GC.Item2.Quantity > 0 Then
            If currentItemPricing.TotalPrice2 <> objPricing.TotalPrice2 Then
                Valid = False
            End If
        End If
        If GC.Item3.Quantity > 0 Then
            If currentItemPricing.TotalPrice3 <> objPricing.TotalPrice3 Then
                Valid = False
            End If
        End If
        If GC.Item4.Quantity > 0 Then
            If currentItemPricing.TotalPrice4 <> objPricing.TotalPrice4 Then
                Valid = False
            End If
        End If
        If GC.Item5.Quantity > 0 Then
            If currentItemPricing.TotalPrice5 <> objPricing.TotalPrice5 Then
                Valid = False
            End If
        End If

        Return Valid
    End Function

    Private Function CalculateCurrentOrderPricing(GCOrder As ClsGiftCertificate2, Optional PriceincludesDiscount As Boolean = False) As ClsPromoPricing
        If CurrentGiftCertificate Is Nothing Then
            Throw New Exception("No gift certificate specified")
        End If

        Dim CUrrentPricing = New ClsPromoPricing

        Try

            CUrrentPricing.ID = 0

            '//Count Tandem Itesm - ie.  Exclude Video Only Item
            ''Exclude Item3 Qty
            Dim IndividualDiscountAmount As Integer = +System.Math.Abs(CalculateItemDiscountAmountFromOrder(GCOrder))

            If GCOrder.Item1.Quantity > 0 Then
                If PriceincludesDiscount Then
                    CUrrentPricing.ItemPrice1 = GCOrder.Item1.Price
                Else
                    CUrrentPricing.ItemPrice1 = GCOrder.Item1.Price + IndividualDiscountAmount
                End If

                CUrrentPricing.ItemDiscount1 = 0 - IndividualDiscountAmount
            End If

            If GCOrder.Item2.Quantity > 0 Then
                If PriceincludesDiscount Then
                    CUrrentPricing.ItemPrice2 = GCOrder.Item2.Price
                Else
                    CUrrentPricing.ItemPrice2 = GCOrder.Item2.Price + IndividualDiscountAmount
                End If
                CUrrentPricing.ItemDiscount2 = 0 - IndividualDiscountAmount
            End If

            If GCOrder.Item3.Quantity > 0 Then
                If PriceincludesDiscount Then
                    CUrrentPricing.ItemPrice3 = GCOrder.Item3.Price
                Else
                    CUrrentPricing.ItemPrice3 = GCOrder.Item3.Price + IndividualDiscountAmount
                End If
                CUrrentPricing.ItemDiscount3 = 0 - IndividualDiscountAmount
            End If

            If GCOrder.Item4.Quantity > 0 Then
                If PriceincludesDiscount Then
                    CUrrentPricing.ItemPrice4 = GCOrder.Item4.Price
                Else
                    CUrrentPricing.ItemPrice4 = GCOrder.Item4.Price + IndividualDiscountAmount
                End If
                CUrrentPricing.ItemDiscount4 = 0 - IndividualDiscountAmount
            End If

            If GCOrder.Item5.Quantity > 0 Then
                If PriceincludesDiscount Then
                    CUrrentPricing.ItemPrice5 = GCOrder.Item5.Price
                Else
                    CUrrentPricing.ItemPrice5 = GCOrder.Item5.Price + IndividualDiscountAmount
                End If
                CUrrentPricing.ItemDiscount5 = 0 - IndividualDiscountAmount
            End If


            Return CUrrentPricing

        Catch ex As Exception

        End Try





    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If CurrentGiftCertificate Is Nothing Then
                MsgBox("No GC Order has been selected")
                'CurrentGiftCertificate.JR_PurchaserID = 0
                'LblJumprunCustomerStatusPurchaser.Text = "Create a new customer in Jumprun"
                'JRCustomerSpecified = True
            Else
                CurrentGiftCertificate.JR_PurchaserID = 0
                LblJumprunCustomerStatusPurchaser.Text = "Create a new customer in Jumprun"
                JRCustomerSpecified = True
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim ObjJumprunSelectedCustomer = CType(SfDGPurchaser.SelectedItem, ClsJumpRunPossibleCustomers)
            If ObjJumprunSelectedCustomer IsNot Nothing Then
                CurrentGiftCertificate.JR_PurchaserID = ObjJumprunSelectedCustomer.wCustId
            Else
                LblJumprunCustomerStatusPurchaser.Text = ""
            End If

            LblJumprunCustomerStatusPurchaser.Text = "Use selected customer in Jumprun " & CurrentGiftCertificate.JR_PurchaserID.ToString
            JRCustomerSpecified = True
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All

            End Select
            Dim dteEntry As Date = SfDateEntry.Value
            SfDataGrid1.SelectedItems.Clear()

            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            If fl IsNot Nothing AndAlso fl.Count > 0 Then
                SfDataGrid1.DataSource = fl
            Else
                BlankDisplayFields()
                SfDGPurchaser.DataSource = Nothing
            End If
            AllowColumnReordering()


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub





    Private Sub FrmProcess_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            SetupForm()
            Dim dteEntry As Date = SfDateEntry.Value
            initializegrid()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub



    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate2)
                PopulateCurrentGiftCertificateDetails(CurrentGiftCertificate)
                ' Button1.Enabled = True
                If CurrentGiftCertificate.GC_Status = 0 Then
                    Button5.Enabled = True
                Else
                    Button5.Enabled = False
                End If

                If CurrentGiftCertificate.GC_DiscountCode <> String.Empty Then
                    CboPromoPricing.SelectedValue = 1
                End If
            Else
                BlankDisplayFields()
                Button5.Enabled = False
                Button1.Enabled = False
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub BlankDisplayFields()
        'Used to populate a form from the records        
        SetTextField(lblPurchaser_FirstName, "")
        SetTextField(lblPurchaser_LastName, "")
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

        SetTextField(Me.lblDiscountAmount, "")
        SetTextField(Me.LblDateEntered, "")

        'Item Fields
        LblItem1Qty.Text = ""
        LblItem2Qty.Text = ""
        LblItem3Qty.Text = ""
        LblItem4Qty.Text = ""
        LblItem5Qty.Text = ""


        lblTotalAmount.Text = ""
        lblDiscountAmount.Text = ""
        lblDiscountCode.Text = ""

        lblDeliveryOption.Text = ""
        lblPointOfSale.Text = ""
        lblPaymentMethod.Text = ""
        LblJumprunCustomerStatusPurchaser.Text = "JumpRun customer Association Not Specified"
        Panel2.Visible = False

        SfDGPurchaser.DataSource = Nothing


    End Sub

    Sub SetupForm()
        Try
            If Me.DefaultDate.HasValue Then
                Me.SfDateEntry.Value = Me.DefaultDate.Value
            Else
                Me.SfDateEntry.Value = Now.Date

            End If

            Dim x = RetrievePricingPromos(True)

            CboPromoPricing.DataSource = x
            CboPromoPricing.DisplayMember = "Value"
            CboPromoPricing.ValueMember = "Key"


            'Populate fields from LIst which I want to show - Purchaser, Recipient, Shipper 
            SfDataGrid1.TableControl.VerticalScrollBarVisible = True

            SfDataGrid1.AutoGenerateColumns = False
            SfDataGrid1.AllowResizingColumns = True
            SfDataGrid1.Columns.Clear()
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "ID", .HeaderText = "Id"})

            'SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_FirstName", .HeaderText = "Purchase First"})
            'SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_LastName", .HeaderText = "Purchase Last"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_Name", .HeaderText = "Purchaser Name"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_Status", .HeaderText = "Status"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "PointOfSale", .HeaderText = "Point Of Sale"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Online_OrderNumber", .HeaderText = "Online Order Number"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item1.Quantity", .HeaderText = "Tandem 10k"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item2.Quantity", .HeaderText = "Tandem 12k"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item3.Quantity", .HeaderText = "Tandem 10k With Vid"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item4.Quantity", .HeaderText = "Tandem 12k With Vid"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item5.Quantity", .HeaderText = "Video"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_TotalAmount", .HeaderText = "OrderAmount"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_TotalDiscount", .HeaderText = "DiscountAmount"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "PersonalizedFrom", .HeaderText = "Personalized From"})
            SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "OriginalOrderDate", .HeaderText = "Order Date"})

            SetupPurchaserDatagrid()

            'SfDGRecipient.AutoGenerateColumns = False
            'SfDGRecipient.AllowResizingColumns = True
            'SfDGRecipient.Columns.Clear()
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "wCustId", .HeaderText = "Id"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "PercentageMatch", .HeaderText = "Score"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sCust", .HeaderText = "Name"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sStreet1", .HeaderText = "Address"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sCity", .HeaderText = "City"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sState", .HeaderText = "State"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sZip", .HeaderText = "Zip"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sEmail", .HeaderText = "Email"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone1", .HeaderText = "Phone 1"})
            'SfDGRecipient.Columns.Add(New GridTextColumn() With {.MappingName = "sPhone2", .HeaderText = "Phone 2"})


            LblItem1.Text = GetDescriptionForItemId(1)
            LblItem2.Text = GetDescriptionForItemId(2)
            LblItem3.Text = GetDescriptionForItemId(3)
            LblItem4.Text = GetDescriptionForItemId(4)
            LblItem5.Text = GetDescriptionForItemId(5)

            CboAuthorizer.DataSource = PopulateAuthorizerList()
            CboAuthorizer.ValueMember = "Key"
            CboAuthorizer.DisplayMember = "Value"

            'Status of buttons will be disabled by default
            BlankDisplayFields()

            AddHandler SfDGPurchaser.QueryRowStyle, AddressOf SfDataGrid1_QueryRowStyle
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub SetupPurchaserDatagrid()
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
    End Sub

    Private Sub SfDataGrid1_QueryRowStyle(sender As Object, e As QueryRowStyleEventArgs) Handles SfDGPurchaser.QueryRowStyle
        If e.RowType = RowType.DefaultRow Then
            Dim item = TryCast(e.RowData, ClsJumpRunPossibleCustomers)
            If item IsNot Nothing Then
                If item.PercentageMatch >= 20 And item.PercentageMatch < 60 Then
                    e.Style.BackColor = Color.LightGoldenrodYellow
                    e.Style.TextColor = Color.Black
                ElseIf item.PercentageMatch >= 60 And item.PercentageMatch < 85 Then
                    e.Style.BackColor = Color.LightGreen
                    e.Style.TextColor = Color.Black
                ElseIf item.PercentageMatch >= 85 Then
                    e.Style.BackColor = Color.Green
                    e.Style.TextColor = Color.White
                Else
                    e.Style.BackColor = Color.White
                    e.Style.TextColor = Color.Black
                End If
            Else
                e.Style.BackColor = Color.White
                e.Style.TextColor = Color.Black
        End If


        End If
    End Sub

    Function PopulateAuthorizerList() As List(Of KeyValuePair(Of Integer, String))
        Dim AuthorizerList As New List(Of KeyValuePair(Of Integer, String))
        Try
            AuthorizerList = RetrieveAuthorizers()
        Catch ex As Exception
            ' AuthorizerList.Clear()
            'AuthorizerList.Add(New KeyValuePair(Of String, String)(6, "Spotty"))
            'AuthorizerList.Add(New KeyValuePair(Of String, String)("JS", "Jim"))
            'AuthorizerList.Add(New KeyValuePair(Of String, String)("TH", "Tyson"))
            'AuthorizerList.Add(New KeyValuePair(Of String, String)("LB", "Leila"))
            'AuthorizerList.Add(New KeyValuePair(Of String, String)("SS", "SS Office"))
        End Try


        Return AuthorizerList

    End Function

    Sub PopulateCurrentGiftCertificateDetails(gc As ClsGiftCertificate2)
        Try
            'Used to populate a form from the records        
            SetTextField(lblPurchaser_FirstName, gc.Purchaser_FirstName)
            SetTextField(lblPurchaser_LastName, gc.Purchaser_LastName)
            SetTextField(lblPurchaser_Address1, gc.Billing_Address.Address1)
            SetTextField(lblPurchaser_Address2, gc.Billing_Address.Address2)
            SetTextField(lblPurchaser_City, gc.Billing_Address.City)
            SetTextField(lblPurchaser_State, gc.Billing_Address.State)
            SetTextField(lblPurchaser_Zip, gc.Billing_Address.Zip)
            SetTextField(lblPurchaser_Phone1, gc.Billing_Address.Phone1)
            SetTextField(lblPurchaser_Phone2, gc.Billing_Address.Address2)
            SetTextField(lblPurchaser_Email, gc.Billing_Address.Email)


            'Recipient Fields
            SetTextField(Me.lblRecipient_Address1, gc.Shipping_Address.Address1)
            SetTextField(Me.lblRecipient_Address2, gc.Shipping_Address.Address2)
            SetTextField(Me.lblRecipient_City, gc.Shipping_Address.City)
            SetTextField(Me.lblRecipient_State, gc.Shipping_Address.State)
            SetTextField(Me.lblRecipient_Zip, gc.Shipping_Address.Zip)
            SetTextField(Me.lblRecipient_Email, gc.Shipping_Address.Email)

            SetTextField(Me.lblDiscountAmount, gc.GC_Authorization)
            SetTextField(Me.LblDateEntered, gc.GC_DateEntered.ToShortDateString)
            SetTextField(Me.TextBox3, gc.PersonalizedFrom)

            'Item Fields
            LblItem1Qty.Text = gc.Item1.Quantity
            LblItem2Qty.Text = gc.Item2.Quantity
            LblItem3Qty.Text = gc.Item3.Quantity
            LblItem4Qty.Text = gc.Item4.Quantity
            LblItem5Qty.Text = gc.Item5.Quantity


            Dim DeliveryOption As String = ""
            Select Case gc.delivery
                Case DeliveryOptions.Email
                    DeliveryOption = "Email"
                Case DeliveryOptions.USMail
                    DeliveryOption = "US Mail"
                Case DeliveryOptions.USDiscreet
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
                Case PaymentMethod.Online_authorize_net_cim_credit_card
                    PM = "authorize_net_cim_credit_card"
            End Select
            lblPaymentMethod.Text = PM
            lblTotalAmount.Text = String.Format("{0}", gc.GC_TotalAmount)
            lblDiscountAmount.Text = String.Format("{0}", gc.GC_TotalDiscount)
            lblDiscountCode.Text = gc.GC_DiscountCode

            'IS recipient even specified if not then 
            If String.IsNullOrEmpty(lblRecipient_FirstName.Text) AndAlso
                    String.IsNullOrEmpty(lblRecipient_LastName.Text) AndAlso
                    String.IsNullOrEmpty(lblRecipient_City.Text) AndAlso
                    String.IsNullOrEmpty(lblRecipient_Zip.Text) AndAlso
                    String.IsNullOrEmpty(lblRecipient_Phone1.Text) AndAlso
                    String.IsNullOrEmpty(lblRecipient_Email.Text) Then

                Panel2.Visible = False
            ElseIf lblRecipient_FirstName.Text.Trim = lblPurchaser_FirstName.Text.Trim AndAlso
                    lblRecipient_LastName.Text.Trim = lblPurchaser_LastName.Text.Trim Then

                Panel2.Visible = False

            Else

                Panel2.Visible = True
            End If

            SetupPurchaserDatagrid()
            'Get Potential Matches for name
            Dim objperson1 = New ClsPersonSearch() With {.FirstName = lblPurchaser_FirstName.Text.Trim,
                                                         .LastName = lblPurchaser_LastName.Text.Trim,
                                                         .Email = lblPurchaser_Email.Text.Trim,
                                                         .Phone = lblPurchaser_Phone1.Text.Trim,
                                                         .Zip = lblPurchaser_Zip.Text.Trim}
            Dim x1 = LoadMatchData(objperson1)


            SfDGPurchaser.DataSource = x1
            If x1.Count = 0 Then
                If CurrentGiftCertificate IsNot Nothing Then
                    CurrentGiftCertificate.JR_PurchaserID = 0
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

            If gc.JR_PurchaserID > 0 Then
                If x1.Count > 0 Then
                    SetJumpRunMatchInDataGrid(gc.JR_PurchaserID, SfDGPurchaser)
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


            Else
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub SetJumpRunMatchInDataGrid(ID As Integer, sfgrid As SfDataGrid)
        Try
            'Set selected item to match id if it exists
            Dim records = sfgrid.View.Records
            For Each record In records
                Dim obj = TryCast(record.Data, ClsJumpRunPossibleCustomers)
                If obj.wCustId = ID Then
                    sfgrid.SelectedItems.Add(obj)
                End If
            Next record
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub SetTextField(control As Label, item As String)
        Try
            If String.IsNullOrEmpty(item) Then
                control.Text = ""
            Else
                control.Text = item
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub
    Private Sub SetTextField(control As TextBox, item As String)
        Try
            If String.IsNullOrEmpty(item) Then
                control.Text = ""
            Else
                control.Text = item
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub



    Public Function GetCertificatesToProcess(entrydate As Date, Filter As FilterStates) As List(Of ClsGiftCertificate2)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate2)
            gclist1 = RetrieveGCOrdersFromQueue(entrydate)

            Dim FilteredList As List(Of ClsGiftCertificate2)

            Select Case Filter
                Case FilterStates.Entered
                    CurrentFilter = FilterStates.Entered
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Select i1).ToList

                Case FilterStates.Processing
                    CurrentFilter = FilterStates.Processing
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Processing Select i1).ToList



                Case FilterStates.Completed
                    CurrentFilter = FilterStates.Completed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Select i1).ToList

                Case FilterStates.ProcessAndCompleteOnly
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
                Case FilterStates.All
                    FilteredList = (From i1 In gclist1 Select i1).ToList
            End Select

            Return FilteredList

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function
    Public Function GetCertificatesToProcess(name As String, Filter As FilterStates) As List(Of ClsGiftCertificate2)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate2)
            gclist1 = RetrieveGCOrdersFromQueue(name)

            Dim FilteredList As List(Of ClsGiftCertificate2)

            Select Case Filter
                Case FilterStates.Entered
                    CurrentFilter = FilterStates.Entered
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Select i1).ToList

                Case FilterStates.Processing
                    CurrentFilter = FilterStates.Processing
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Processing Select i1).ToList



                Case FilterStates.Completed
                    CurrentFilter = FilterStates.Completed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Select i1).ToList

                Case FilterStates.ProcessAndCompleteOnly
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
                Case FilterStates.All
                    FilteredList = (From i1 In gclist1 Select i1).ToList
            End Select

            Return FilteredList

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
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
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All

            End Select
            Dim dteEntry As Date = SfDateEntry.Value
            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            If fl IsNot Nothing AndAlso fl.Count > 0 Then
                SfDataGrid1.DataSource = fl
            Else
                SfDataGrid1.DataSource = Nothing
                BlankDisplayFields()
                SfDGPurchaser.DataSource = Nothing
            End If
            AllowColumnReordering()

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


        End Try
    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Try
            If RadioButton1.Checked Then
                Panel4.Enabled = True
                Panel5.Enabled = False
            Else
                Panel4.Enabled = False
                Panel5.Enabled = True

            End If
            SfDataGrid1.DataSource = Nothing
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All
            End Select
            SfDataGrid1.SelectedItems.Clear()

            Dim fl = GetCertificatesToProcess(TextBox1.Text, CurrentFilter)
            SfDataGrid1.DataSource = fl
            AllowColumnReordering()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub SfDGPurchaser_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDGPurchaser.SelectionChanged
        '//Determine if select button is visible or not
        Dim Possible As ClsJumpRunPossibleCustomers
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                Possible = CType(SfDGPurchaser.SelectedItem, ClsJumpRunPossibleCustomers)

                If Possible IsNot Nothing Then
                    Button1.Enabled = True
                Else
                    Button1.Enabled = False
                End If
            Else

                Button1.Enabled = False
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        Using file = System.IO.File.Create("FrmProcess1.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using
        Using file = System.IO.File.Create("FrmProcess2.xml")
            Me.SfDGPurchaser.Serialize(file)
        End Using
    End Sub

    Public Sub initializegrid()
        AllowColumnReordering()

        If IO.File.Exists("FrmProcess1.xml") Then
            Try
                Using file = System.IO.File.Open("FrmProcess1.xml", System.IO.FileMode.Open)
                    Me.SfDataGrid1.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If
        If IO.File.Exists("FrmProcess2.xml") Then
            Try
                Using file = System.IO.File.Open("FrmProcess2.xml", System.IO.FileMode.Open)
                    Me.SfDGPurchaser.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub AllowColumnReordering()
        SfDataGrid1.AllowDraggingColumns = True
        For Each c In SfDataGrid1.Columns
            c.AllowDragging = True

        Next
        SfDataGrid1.AllowDrop = True

        SfDGPurchaser.AllowDraggingColumns = True
        For Each c In SfDGPurchaser.Columns
            c.AllowDragging = True
        Next
        SfDGPurchaser.AllowDrop = True
    End Sub

    Private Sub TabPageBillingAddress_Click(sender As Object, e As EventArgs) Handles TabPageBillingAddress.Click

    End Sub

    Private Sub CboPromoPricing_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPromoPricing.SelectedIndexChanged
        ObjPricing = RetrievePricingPromoForId(CType(CboPromoPricing.SelectedItem, KeyValuePair(Of Integer, String)).Key)
    End Sub

    Private Sub FrmProcess_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Using file = System.IO.File.Create("FrmProcess1.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using
        Using file = System.IO.File.Create("FrmProcess2.xml")
            Me.SfDGPurchaser.Serialize(file)
        End Using
    End Sub


End Class