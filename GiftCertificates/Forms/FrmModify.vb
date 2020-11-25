Imports System.Reflection
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events

Public Class FrmModify
    Private BlnChanged As Boolean = False
    Public Property CurrentGiftCertificate As ClsGiftCertificate2

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("Search for all gift certificates with a last name match")
        Dim searchLastName As String = TxtSearchLastName.Text.Trim
        '[dbo].[GCFindCertificatesWithLastName]
        Dim list = GCOrders_Search(searchLastName)
        SfDataGrid1.DataSource = list
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("Select the gift certificate item from the list")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        MsgBox("Modify the gift certificate to correct any name changes.  This can be done to correct spelling mistakes and reprint")
        MsgBox("It does not change the customer records - so this needs to be used carefully")
    End Sub

    Private Sub FrmModify_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupForm()
    End Sub

    Private Sub SetupForm()

        LblItem1Description.Text = GetDescriptionForItemId(1)
        LblItem2Description.Text = GetDescriptionForItemId(2)
        LblItem3Description.Text = GetDescriptionForItemId(3)
        LblItem4Description.Text = GetDescriptionForItemId(4)
        LblItem5Description.Text = GetDescriptionForItemId(5)

        CboPayment.DataSource = Lists.RetrievePaymentMethodList
        CboPayment.DisplayMember = "Value"
        CboPayment.ValueMember = "Key"

        CBoHearAbout.DataSource = Lists.RetrieveHearAboutList
        CBoHearAbout.DisplayMember = "Value"
        CBoHearAbout.ValueMember = "Key"

        TxtId.Text = "<New>"
        RdoDeliveryInPerson.Checked = True
        RdoPosPhone.Checked = True
        SfDateTimeEdit1.Value = Now.Date

        ComboBox1.DataSource = PopulateQuantityItems()
        ComboBox1.DisplayMember = "Value"
        ComboBox1.ValueMember = "Key"
        LblItem1Description.Text = GetDescriptionForItemId(1)


        ComboBox2.DataSource = PopulateQuantityItems()
        ComboBox2.DisplayMember = "Value"
        ComboBox2.ValueMember = "Key"
        LblItem2Description.Text = GetDescriptionForItemId(2)

        ComboBox3.DataSource = PopulateQuantityItems()
        ComboBox3.DisplayMember = "Value"
        ComboBox3.ValueMember = "Key"
        LblItem3Description.Text = GetDescriptionForItemId(3)

        ComboBox4.DataSource = PopulateQuantityItems()
        ComboBox4.DisplayMember = "Value"
        ComboBox4.ValueMember = "Key"
        LblItem4Description.Text = GetDescriptionForItemId(4)

        ComboBox5.DataSource = PopulateQuantityItems()
        ComboBox5.DisplayMember = "Value"
        ComboBox5.ValueMember = "Key"
        LblItem5Description.Text = GetDescriptionForItemId(5)

        PopulateDiscountCodes()
        ComboBox6.SelectedIndex = 0

        LblItem1Amount.Text = 0
    End Sub
    Sub PopulateDiscountCodes()
        Try
            Dim LstPricing = RetrieveDiscounts()
            ComboBox6.Items.Clear()
            ComboBox6.Items.Add(" ")
            For Each i In LstPricing
                ComboBox6.Items.Add(i.SKU)
            Next
        Catch ex As Exception

        End Try
    End Sub
    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        '//Populate The Form
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate2)
                PopulateCurrentGiftCertificateDetails(CurrentGiftCertificate)
                ' Button1.Enabled = True

            Else
                BlankDisplayFields()

            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub BlankDisplayFields()
        ''Used to populate a form from the records        
        'SetTextField(lblPurchaser_FirstName, "")
        'SetTextField(lblPurchaser_LastName, "")
        'SetTextField(lblPurchaser_Address1, "")
        'SetTextField(lblPurchaser_Address2, "")
        'SetTextField(lblPurchaser_City, "")
        'SetTextField(lblPurchaser_State, "")
        'SetTextField(lblPurchaser_Zip, "")
        'SetTextField(lblPurchaser_Phone1, "")
        'SetTextField(lblPurchaser_Phone2, "")
        'SetTextField(lblPurchaser_Email, "")


        ''Recipient Fields
        'SetTextField(Me.lblRecipient_FirstName, "")
        'SetTextField(Me.lblRecipient_LastName, "")
        'SetTextField(Me.lblRecipient_Address1, "")
        'SetTextField(Me.lblRecipient_Address2, "")
        'SetTextField(Me.lblRecipient_City, "")
        'SetTextField(Me.lblRecipient_State, "")
        'SetTextField(Me.lblRecipient_Zip, "")
        'SetTextField(Me.lblRecipient_Phone1, "")
        'SetTextField(Me.lblRecipient_Phone2, "")
        'SetTextField(Me.lblRecipient_Email, "")

        'SetTextField(Me.lblDiscountAmount, "")
        'SetTextField(Me.LblDateEntered, "")

        ''Item Fields
        'LblItem1Qty.Text = ""
        'LblItem2Qty.Text = ""
        'LblItem3Qty.Text = ""
        'LblItem4Qty.Text = ""
        'LblItem5Qty.Text = ""


        'lblTotalAmount.Text = ""
        'lblDiscountAmount.Text = ""
        'lblDiscountCode.Text = ""

        'lblDeliveryOption.Text = ""
        'lblPointOfSale.Text = ""
        'lblPaymentMethod.Text = ""
        'LblJumprunCustomerStatusPurchaser.Text = "JumpRun customer Association Not Specified"
        'Panel2.Visible = False

        'SfDGPurchaser.DataSource = Nothing
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        '//Save Updated record
        ' If BlnChanged Then
        Dim GC As New ClsGiftCertificate2

            GC = CurrentGiftCertificate

            GC.PaymentMethod = CboPayment.SelectedValue

            Dim s1 = Module4.WhatRadioIsSelected(Me.Panel2)
            Select Case s1.ToLower
                Case "rdodeliveryemail"
                    GC.delivery = DeliveryOptions.Email
                Case "rdodeliveryusmail"
                    GC.delivery = DeliveryOptions.USMail
                Case "rdodeliveryusmaildiscrete"
                    GC.delivery = DeliveryOptions.USDiscreet
                Case "rdodeliveryinperson"
                    GC.delivery = DeliveryOptions.InOffice
            End Select

            Dim s2 = WhatRadioIsSelected(Me.Panel3)
            Select Case s2.ToLower
                Case "rdoposphone"
                    GC.PointOfSale = PointOfSale.PhoneInPerson

                Case "rdoposonline"
                    GC.PointOfSale = PointOfSale.Online
            End Select


            GC.HearAbout = CBoHearAbout.SelectedValue

            GC.PaymentNotes = TxtPaymentNotes.Text
            GC.Notes = TxtNotes.Text


            GC.Purchaser_FirstName = TxtPurchaserFirstName.Text
            GC.Purchaser_LastName = TxtPurchaserLastName.Text
            GC.Billing_Address.Address1 = TxtPurchaserAddress1.Text
            GC.Billing_Address.Address2 = TxtPurchaserAddress2.Text
            GC.Billing_Address.City = TxtPurchaserCity.Text
            GC.Billing_Address.State = TxtPurchaserState.Text
            GC.Billing_Address.Zip = TxtPurchaserZip.Text
            GC.Billing_Address.Phone1 = TxtPurchaserPhone1.Text
            GC.Billing_Address.Phone2 = TxtPurchaserPhone2.Text
            GC.Billing_Address.Email = TxtPurchaserEmail.Text

            GC.Shipping_Address.Address1 = TxtRecipientAddress1.Text
            GC.Shipping_Address.Address2 = TxtRecipientAddress2.Text
            GC.Shipping_Address.City = TxtRecipientCity.Text
            GC.Shipping_Address.State = TxtRecipientState.Text
            GC.Shipping_Address.Zip = TxtRecipientZip.Text
            GC.GC_Authorization = TxtAuthorization.Text
            GC.GC_DiscountCode = ComboBox6.Text.Trim.ToUpper

            '//These use individual values rather than numbers
            Dim Qty1 = CInt(CType(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty2 = CInt(CType(ComboBox2.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty3 = CInt(CType(ComboBox3.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty4 = CInt(CType(ComboBox4.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty5 = CInt(CType(ComboBox5.SelectedItem, KeyValuePair(Of Integer, String)).Key)


            GC.Item1.Quantity = Qty1
            GC.Item2.Quantity = Qty2
            GC.Item3.Quantity = Qty3
            GC.Item4.Quantity = Qty4
            GC.Item5.Quantity = Qty5

            Dim TotalAmount As Double = 0
            Dim TotalDiscount As Double = 0

            '//Calculate Amount and discount
            Dim Pricing1 = GetPricingForItemId(1)
            Dim DiscountAmount1 As Double = 0
            Dim Amount1 = GC.Item1.Quantity * Pricing1.Price
            If Pricing1.Discountable Then
                Dim dISCOUNT1 = GetPricingForDiscountId(GC.GC_DiscountCode)
                If dISCOUNT1 IsNot Nothing Then
                    DiscountAmount1 = GC.Item1.Quantity * dISCOUNT1.Price
                Else
                    DiscountAmount1 = 0
                End If

            End If

            Dim Pricing2 = GetPricingForItemId(2)
            Dim DiscountAmount2 As Double = 0
        Dim Amount2 = GC.Item3.Quantity * Pricing2.Price
        If Pricing2.Discountable Then
                Dim dISCOUNT2 = GetPricingForDiscountId(GC.GC_DiscountCode)
                If dISCOUNT2 IsNot Nothing Then
                    DiscountAmount2 = GC.Item2.Quantity * dISCOUNT2.Price
                Else
                    DiscountAmount2 = 0
                End If
            End If

            Dim Pricing3 = GetPricingForItemId(3)
            Dim DiscountAmount3 As Double = 0
        Dim Amount3 = GC.Item3.Quantity * Pricing3.Price
        If Pricing3.Discountable Then
                Dim dISCOUNT3 = GetPricingForDiscountId(GC.GC_DiscountCode)
                If dISCOUNT3 IsNot Nothing Then
                    DiscountAmount3 = GC.Item3.Quantity * dISCOUNT3.Price
                Else
                    DiscountAmount3 = 0
                End If
            End If

            Dim Pricing4 = GetPricingForItemId(4)
            Dim DiscountAmount4 As Double = 0
        Dim Amount4 = GC.Item4.Quantity * Pricing4.Price
        If Pricing4.Discountable Then
                Dim dISCOUNT4 = GetPricingForDiscountId(GC.GC_DiscountCode)
                If dISCOUNT4 IsNot Nothing Then
                    DiscountAmount4 = GC.Item4.Quantity * dISCOUNT4.Price
                Else
                    DiscountAmount4 = 0
                End If
            End If

            Dim Pricing5 = GetPricingForItemId(5)
            Dim DiscountAmount5 As Double = 0
        Dim Amount5 = GC.Item5.Quantity * Pricing5.Price
        If Pricing5.Discountable Then
                Dim dISCOUNT5 = GetPricingForDiscountId(GC.GC_DiscountCode)
                If dISCOUNT5 IsNot Nothing Then
                    DiscountAmount5 = GC.Item5.Quantity * dISCOUNT5.Price
                Else
                    DiscountAmount5 = 0
                End If
            End If

            TotalAmount = Amount1 + Amount2 + Amount3 + Amount4 + Amount5
            TotalDiscount = DiscountAmount1 + DiscountAmount2 + DiscountAmount3 + DiscountAmount4 + DiscountAmount5

            GC.GC_TotalAmount = TotalAmount
            GC.GC_TotalDiscount = TotalDiscount
            GC.GC_DiscountCode = ComboBox6.Text.Trim.ToUpper


            'Point Of Sale
            If RdoPOSOnline.Checked Then
                GC.PointOfSale = PointOfSale.Online
            Else
                If RdoPosPhone.Checked Then
                End If
            End If

            UpdateGiftCertificatesFromCertificateRecord(GC)



            MsgBox("Update Save TO BE IMPLEMENTED")
        'End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If BlnChanged Then
            '//Warn of unsaved changes
        End If
        Me.Close()

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

    Sub PopulateCurrentGiftCertificateDetails(gc As ClsGiftCertificate2)
        Try


            SfDateTimeEdit1.Value = gc.GC_DateEntered

            'Used to populate a form from the records        
            SetTextField(TxtPurchaserFirstName, gc.Purchaser_FirstName)
            SetTextField(TxtPurchaserLastName, gc.Purchaser_LastName)
            SetTextField(TxtPurchaserAddress1, gc.Billing_Address.Address1)
            SetTextField(TxtPurchaserAddress2, gc.Billing_Address.Address2)
            SetTextField(TxtPurchaserCity, gc.Billing_Address.City)
            SetTextField(TxtPurchaserState, gc.Billing_Address.State)
            SetTextField(TxtPurchaserZip, gc.Billing_Address.Zip)
            SetTextField(TxtPurchaserPhone1, gc.Billing_Address.Phone1)
            SetTextField(TxtPurchaserPhone2, gc.Billing_Address.Phone2)

            SetTextField(TxtPurchaserEmail, gc.Billing_Address.Email)


            'Recipient Fields
            SetTextField(Me.TxtRecipientAddress1, gc.Shipping_Address.Address1)
            SetTextField(Me.TxtRecipientAddress2, gc.Shipping_Address.Address2)
            SetTextField(Me.TxtRecipientCity, gc.Shipping_Address.City)
            SetTextField(Me.TxtRecipientState, gc.Shipping_Address.State)
            SetTextField(Me.TxtRecipientZip, gc.Shipping_Address.Zip)

            'SetTextField(Me.lblDiscountAmount, gc.GC_Authorization)
            'SetTextField(Me.LblDateEntered, gc.GC_DateEntered.ToShortDateString)

            ''Item Fields
            ComboBox1.SelectedValue = gc.Item1.Quantity
            ComboBox2.SelectedValue = gc.Item2.Quantity
            ComboBox3.SelectedValue = gc.Item3.Quantity
            ComboBox4.SelectedValue = gc.Item4.Quantity
            ComboBox5.SelectedValue = gc.Item5.Quantity


            Dim DeliveryOption As String = ""
            Select Case gc.delivery
                Case DeliveryOptions.Email
                    RdoDeliveryEmail.Checked = True

                Case DeliveryOptions.USMail
                    RdoDeliveryUSMail.Checked = True
                Case DeliveryOptions.USDiscreet
                    RdoDeliverUSMailDiscrete.Checked = True
                Case DeliveryOptions.InOffice
                    RdoDeliveryInPerson.Checked = True
            End Select


            Select Case gc.PointOfSale
                Case PointOfSale.Online
                    RdoPOSOnline.Checked = True
                Case PointOfSale.PhoneInPerson
                    RdoPosPhone.Checked = True
            End Select

            CboPayment.SelectedIndex = gc.PaymentMethod
            CBoHearAbout.SelectedIndex = gc.HearAbout

            ComboBox6.Text = gc.GC_DiscountCode.ToUpper

            TxtNotes.Text = gc.Notes
            TxtAuthorization.Text = gc.GC_Authorization

            '//Only allow modifying of certain items
            If gc.GC_Status = CertificateStatus.Entered Then
                Panel1.Enabled = True
                Panel2.Enabled = True
                Panel3.Enabled = True
                ComboBox6.Enabled = True
                TxtPersonalizedFrom.Enabled = True
                GroupBox1.Enabled = True
                TxtAuthorization.Enabled = False
                CBoHearAbout.Enabled = True
            Else
                Panel1.Enabled = False
                Panel2.Enabled = False
                Panel3.Enabled = False
                ComboBox6.Enabled = False
                TxtPersonalizedFrom.Enabled = False
                GroupBox1.Enabled = False
                TxtAuthorization.Enabled = False
                CBoHearAbout.Enabled = False
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged,
                                                                                     ComboBox2.SelectedIndexChanged,
                                                                                     ComboBox3.SelectedIndexChanged,
                                                                                     ComboBox4.SelectedIndexChanged,
                                                                                     ComboBox5.SelectedIndexChanged,
                                                                                     ComboBox6.SelectedIndexChanged


        CalculateTotals()
    End Sub


    Sub CalculateTotals()
        Try
            Dim TotalAmount As Double = 0
            Dim TotalDiscount As Double = 0
            Dim DiscountCode As String = ComboBox6.Text

            Dim Qty1 = CInt(CType(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty2 = CInt(CType(ComboBox2.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty3 = CInt(CType(ComboBox3.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty4 = CInt(CType(ComboBox4.SelectedItem, KeyValuePair(Of Integer, String)).Key)
            Dim Qty5 = CInt(CType(ComboBox5.SelectedItem, KeyValuePair(Of Integer, String)).Key)


            '//Calculate Amount and discount
            Dim Pricing1 = GetPricingForItemId(1)
            Dim DiscountAmount1 As Double = 0
            Dim Amount1 = Qty1 * Pricing1.Price
            If Pricing1.Discountable Then
                Dim dISCOUNT1 = GetPricingForDiscountId(DiscountCode)
                If dISCOUNT1 IsNot Nothing Then
                    DiscountAmount1 = Qty1 * dISCOUNT1.Price
                End If
                LblItem1Amount.Text = Amount1.ToString
            End If

            Dim Pricing2 = GetPricingForItemId(2)
            Dim DiscountAmount2 As Double = 0
            Dim Amount2 = Qty2 * Pricing2.Price
            If Pricing2.Discountable Then
                Dim dISCOUNT2 = GetPricingForDiscountId(DiscountCode)
                If dISCOUNT2 IsNot Nothing Then
                    DiscountAmount2 = Qty2 * dISCOUNT2.Price
                End If
                LblItem2Amount.Text = Amount2.ToString
            End If

            Dim Pricing3 = GetPricingForItemId(3)
            Dim DiscountAmount3 As Double = 0
            Dim Amount3 = Qty3 * Pricing3.Price
            If Pricing3.Discountable Then
                Dim dISCOUNT3 = GetPricingForDiscountId(DiscountCode)
                If dISCOUNT3 IsNot Nothing Then
                    DiscountAmount3 = Qty3 * dISCOUNT3.Price
                End If
                LblItem3Amount.Text = Amount3.ToString
            End If

            Dim Pricing4 = GetPricingForItemId(4)
            Dim DiscountAmount4 As Double = 0
            Dim Amount4 = Qty4 * Pricing4.Price
            If Pricing4.Discountable Then
                Dim dISCOUNT4 = GetPricingForDiscountId(DiscountCode)
                If dISCOUNT4 IsNot Nothing Then
                    DiscountAmount4 = Qty4 * dISCOUNT4.Price
                End If
                LblItem4Amount.Text = Amount4.ToString
            End If

            Dim Pricing5 = GetPricingForItemId(5)
            Dim DiscountAmount5 As Double = 0
            Dim Amount5 = Qty5 * Pricing5.Price
            If Pricing5.Discountable Then
                Dim dISCOUNT5 = GetPricingForDiscountId(DiscountCode)
                If dISCOUNT5 IsNot Nothing Then
                    DiscountAmount5 = Qty5 * dISCOUNT5.Price
                End If
                LblItem5Amount.Text = Amount5.ToString
            End If

            TotalAmount = Amount1 + Amount2 + Amount3 + Amount4 + Amount5
            TotalDiscount = DiscountAmount1 + DiscountAmount2 + DiscountAmount3 + DiscountAmount4 + DiscountAmount5

            LblCalculatedTotal.Text = TotalAmount.ToString
            LblCalculatedDiscount.Text = TotalDiscount.ToString
        Catch ex As Exception

        End Try



    End Sub


End Class