Public Class FrmEntry



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ValidateEntry() Then

            CreateCertificate()


            Else
            Exit Sub
        End If

        If CheckBox6.Checked = False Then
            Me.Close()
        Else
            'Reset and leave purchaser details present
            MsgBox("Reset and leave purchaser details present to avoid retyping")
            ResetNonPurchaserFields()
        End If

    End Sub

    Private Sub UpdateCertificate()
        MsgBox("To Be Implemented")
    End Sub

    Sub ResetNonPurchaserFields()
        TxtRecipientFirstName.Text = ""
        TxtRecipientLastName.Text = ""
        TxtRecipientAddress1.Text = ""
        TxtRecipientAddress2.Text = ""
        TxtRecipientCity.Text = ""
        TxtRecipientState.Text = ""
        TxtRecipientZip.Text = ""
        TxtRecipientPhone1.Text = ""
        TxtRecipientPhone2.Text = ""
        TxtRecipientEmail.Text = ""

        TxtPersonalizedTo.Text = ""
        TxtOtherAmount.Text = "0"

        TxtUserName.Text = ""
        TxtAuthorization.Text = ""

        '//These use individual values rather than numbers
        ChkTandem10k.Checked = False


        ChkTandem12k.Checked = False
        ChkVideo.Checked = False

        ChkOther.Checked = False

        'Point Of Sale
        RdoPOSOnline.Checked = False
        RdoDeliveryEmail.Checked = True

        'TODO Other fields do we want to reset them.
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

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PopulateGiftCertificate()
        MsgBox("Populate Gift Certificate - To Be Implemented")
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
            If String.IsNullOrEmpty(TxtPurchaserEmail.Text) And String.IsNullOrEmpty(TxtRecipientEmail.Text) Then
                    sb.AppendLine(String.Format("{0}. Email for delivery is selected and no email has been provided.", IErrorCount))
                    IErrorCount += 1
                    bValidationErrorOccured = True
                End If
        End If

        If ChkTandem10k.Checked = False AndAlso
                ChkTandem12k.Checked = False AndAlso
                ChkVideo.Checked = False AndAlso
                ChkOther.Checked = False Then
                sb.AppendLine(String.Format("{0}. At least one item needs to be selected for the gift certificate.", IErrorCount))
                IErrorCount += 1
                bValidationErrorOccured = True
        End If

        If String.IsNullOrEmpty(TxtRecipientEmail.Text) AndAlso
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
        Dim GC As New ClsGiftCertificate

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
        GC.Purchaser_Address1 = TxtPurchaserAddress1.Text
        GC.Purchaser_Address2 = TxtPurchaserAddress2.Text
        GC.Purchaser_City = TxtPurchaserCity.Text
        GC.Purchaser_State = TxtPurchaserState.Text
        GC.Purchaser_Zip = TxtPurchaserZip.Text
        GC.Purchaser_Phones1 = TxtPurchaserPhone1.Text
        GC.Purchaser_Phones2 = TxtPurchaserPhone2.Text
        GC.Purchaser_Email = TxtPurchaserEmail.Text

        GC.Recipient_FirstName = TxtRecipientFirstName.Text
        GC.Recipient_LastName = TxtRecipientLastName.Text
        GC.Recipient_Address1 = TxtRecipientAddress1.Text
        GC.Recipient_Address2 = TxtRecipientAddress2.Text
        GC.Recipient_City = TxtRecipientCity.Text
        GC.Recipient_State = TxtRecipientState.Text
        GC.Recipient_Zip = TxtRecipientZip.Text
        GC.Recipient_Phones1 = TxtRecipientPhone1.Text
        GC.Recipient_Phones2 = TxtRecipientPhone2.Text
        GC.Recipient_Email = TxtRecipientEmail.Text

        GC.Recipient_PersonalizedFrom = TxtPersonalizedFrom.Text
        GC.Purchaser_PersonalizedTo = TxtPersonalizedTo.Text

        'GC.Item_OtherAmount = CDbl(TxtOtherAmount.Text)
        GC.GC_Username = TxtUserName.Text
        GC.GC_Authorization = TxtAuthorization.Text

        '//These use individual values rather than numbers
        If ChkTandem10k.Checked Then
            GC.Item_Tandem10k = 1
        Else
            GC.Item_Tandem10k = 0
        End If

        If ChkTandem12k.Checked Then
            GC.Item_Tandem12k = 1
        Else
            GC.Item_Tandem12k = 0
        End If

        If ChkVideo.Checked Then
            GC.Item_Video = 1
        Else
            GC.Item_Video = 0
        End If

        If ChkOther.Checked Then
            GC.Item_Other = 1
            GC.Item_OtherAmount = 1 'TODO Realtextbox
        Else
            GC.Item_Video = 0
            GC.Item_OtherAmount = 0
        End If

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
                GC.delivery = DeliveryOptions.USMailDiscrete
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


        '//Call InsertGiftCertificate
        InsertNewGiftCertRecord(GC)
    End Sub




    Private Sub ChkTandem10k_CheckedChanged(sender As Object, e As EventArgs) Handles ChkTandem10k.CheckedChanged, ChkTandem12k.CheckedChanged, ChkVideo.CheckedChanged, ChkOther.CheckedChanged
        'TODO: 
        '//cALCULATE NEW TOTAL
        'Based upon Price Items we can calculate totals - Do I just want to store these in a config or use the database

        LblCalculatedTotal.Text = "TODO: Calculated total"
    End Sub
End Class