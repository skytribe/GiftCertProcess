Public Class FrmEntry



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ValidateEntry() Then

            CreateCertificate()


            Else
            Exit Sub
        End If


        Me.Close()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmEntry_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
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

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

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



    Function ValidateEntry() As Boolean
        Dim bValidationErrorOccured As Boolean = False
        Dim IErrorCount As Integer = 1
        Try

            'A item is selected
            ' Purchaser details are added 
            ' if credit card then 
            '     somedetails are present - I'd rather have simply an authorization than the card details
            '

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("Validation Errors Occured:")
            'sb.AppendLine("1. Need to provide a first and last name for purchaser")
            'sb.AppendLine("2. If email is selected then need to provide an email")
            'sb.AppendLine("3. At least one item needs to be selected")
            'sb.AppendLine("4. A contact detail  - phone/email needs to be provided")
            'sb.AppendLine("5. If credit card - a reference needs to be entered")

            'MsgBox(sb.ToString)

            If String.IsNullOrEmpty(TxtPurchaserFirstName.Text) Or
           String.IsNullOrEmpty(TxtPurchaserLastName.Text) Then
                sb.AppendLine(String.Format("{0}. Requires a purchase firstname and last name to be specified", IErrorCount))
                IErrorCount += 1
                bValidationErrorOccured = True
            End If

            If RdoDeliveryEmail.Checked Then
                If String.IsNullOrEmpty(TxtPurchaserEmail.Text) Then
                    sb.AppendLine(String.Format("{0}. Email for delivery is selected and no email has been provided.", IErrorCount))
                    IErrorCount += 1
                    bValidationErrorOccured = True
                End If
            End If

            If ComboBox1.Text = "0" AndAlso
               ComboBox2.Text = "0" AndAlso
               ComboBox3.Text = "0" AndAlso
               ComboBox4.Text = "0" AndAlso
                ComboBox5.Text = "0" Then
                sb.AppendLine(String.Format("{0}. At least one item needs to be selected for the gift certificate.", IErrorCount))
                IErrorCount += 1
                bValidationErrorOccured = True
            End If

            If String.IsNullOrEmpty(TxtPurchaserEmail.Text) AndAlso
           String.IsNullOrEmpty(TxtPurchaserPhone1.Text) AndAlso
           String.IsNullOrEmpty(TxtPurchaserPhone2.Text) Then
                sb.AppendLine(String.Format("{0}. Requires at least one phone or email contact to be specified for the purchaser", IErrorCount))
                IErrorCount += 1
                bValidationErrorOccured = True
            End If

            If CboPayment.SelectedValue = PaymentMethod.CreditCard Then
                If String.IsNullOrEmpty(TxtPaymentNotes.Text) Then
                    sb.AppendLine(String.Format("{0}. For credit card transactions a transaction reference should be supplied in the payment notes.", IErrorCount))
                    IErrorCount += 1
                    bValidationErrorOccured = True
                End If
            End If

            If bValidationErrorOccured Then
                Dim strErrors = sb.ToString
                MessageBox.Show(strErrors, "Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            Else
                Return True

            End If

        Catch ex As Exception

        End Try
    End Function

    Sub CreateCertificate()
        Try
            Dim GC As New ClsGiftCertificate2

            GC.PaymentMethod = 0
            GC.delivery = 0
            GC.PointOfSale = PointOfSale.PhoneInPerson
            GC.HearAbout = HearAbout.Unknown
            GC.PaymentNotes = TxtPaymentNotes.Text
            GC.Notes = TxtNotes.Text
            GC.GC_Status = CertificateStatus.Entered
            GC.GC_Username = ""
            GC.GC_Authorization = ""
            GC.GC_DateEntered = SfDateTimeEdit1.Value


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
            Dim Amount2 = GC.Item2.Quantity * Pricing2.Price
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

            'TODO
            GC.PersonalizedFrom = TxtPersonalizedFrom.Text
            '//Call InsertGiftCertificate
            CreateGiftCertificatesFromCertificateRecord(GC)
        Catch ex As Exception

        End Try

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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged,
                                                                                         ComboBox2.SelectedIndexChanged,
                                                                                         ComboBox3.SelectedIndexChanged,
                                                                                         ComboBox4.SelectedIndexChanged,
                                                                                         ComboBox5.SelectedIndexChanged,
                                                                                         ComboBox6.SelectedIndexChanged


        CalculateTotals()
    End Sub
End Class