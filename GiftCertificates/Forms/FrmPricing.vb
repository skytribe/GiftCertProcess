Imports System.Reflection
Imports System.Text

Public Class FrmPricing
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Dim CurrentPromoID As Integer = 0


    Private Sub FrmPricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            GetPromoPricingListItems()

            PopulatePricing()
            PopulateDiscounts()

            Dim Pricing1 = GetPricingForItemId(1)
            If Pricing1 IsNot Nothing Then
                LblItemDescription1.Text = Pricing1.Description
                TxtItemPrice1.Text = Pricing1.Price
                TxtJRItemCode1.Text = Pricing1.JR_ItemID
                ChkDiscountable1.Checked = Pricing1.Discountable
                LblItemDescription1.Tag = Pricing1.ID
                LblItemSku1.Text = Pricing1.SKU

            End If


            Dim Pricing2 = GetPricingForItemId(2)
            If Pricing2 IsNot Nothing Then
                LblItemDescription2.Text = Pricing2.Description
                TxtItemPrice2.Text = Pricing2.Price
                TxtJRItemCode2.Text = Pricing2.JR_ItemID
                ChkDiscountable2.Checked = Pricing2.Discountable
                LblItemDescription2.Tag = Pricing2.ID
                LblItemSku2.Text = Pricing2.SKU

            End If



            Dim Pricing3 = GetPricingForItemId(3)
            If Pricing3 IsNot Nothing Then
                LblItemDescription3.Text = Pricing3.Description
                TxtItemPrice3.Text = Pricing3.Price
                TxtJRItemCode3.Text = Pricing3.JR_ItemID
                ChkDiscountable3.Checked = Pricing3.Discountable
                LblItemDescription3.Tag = Pricing3.ID
                LblItemSku3.Text = Pricing3.SKU

            End If



            Dim Pricing4 = GetPricingForItemId(4)
            If Pricing4 IsNot Nothing Then
                LblItemDescription4.Text = Pricing4.Description
                TxtItemPrice4.Text = Pricing4.Price
                TxtJRItemCode4.Text = Pricing4.JR_ItemID
                ChkDiscountable4.Checked = Pricing4.Discountable
                LblItemDescription4.Tag = Pricing4.ID
                LblItemSku4.Text = Pricing4.SKU

            End If



            Dim Pricing5 = GetPricingForItemId(5)
            If Pricing5 IsNot Nothing Then
                LblItemDescription5.Text = Pricing5.Description
                TxtItemPrice5.Text = Pricing5.Price
                TxtJRItemCode5.Text = Pricing5.JR_ItemID
                ChkDiscountable5.Checked = Pricing5.Discountable
                LblItemDescription5.Tag = Pricing5.ID
                LblItemSku5.Text = Pricing5.SKU

            End If



            Dim dc = RetrieveDiscounts()
            If dc.Count > 4 Then
            Else
                If dc.Item(0) IsNot Nothing Then
                    Dim Discount1 = GetPricingForDiscountId(dc.Item(0).SKU)
                    TxtDiscountCode1.Text = Discount1.Price
                    TxtJRItemDiscountCode1.Text = Discount1.JR_ItemID
                    TxtDiscountSku1.Text = Discount1.SKU
                    lblDiscountDesc1.Tag = Discount1.ID
                End If

                If dc.Item(1) IsNot Nothing Then
                    Dim Discount2 = GetPricingForDiscountId(dc.Item(1).SKU)
                    TxtDiscountCode2.Text = Discount2.Price
                    TxtJRItemDiscountCode2.Text = Discount2.JR_ItemID
                    TxtDiscountSku2.Text = Discount2.SKU
                    lblDiscountDesc2.Tag = Discount2.ID
                End If


                If dc.Item(2) IsNot Nothing Then
                    Dim Discount3 = GetPricingForDiscountId(dc.Item(2).SKU)
                    TxtDiscountCode3.Text = Discount3.Price
                    TxtJRItemDiscountCode3.Text = Discount3.JR_ItemID
                    TxtDiscountSku3.Text = Discount3.SKU
                    lblDiscountDesc3.Tag = Discount3.ID
                End If
                If dc.Item(3) IsNot Nothing Then
                    Dim Discount4 = GetPricingForDiscountId(dc.Item(3).SKU)
                    TxtDiscountCodeFreeAlt.Text = Discount4.Price
                    TxtJRItemDiscountCodeFreeAlt.Text = Discount4.JR_ItemID
                    TxtDiscountFreeAlt.Text = Discount4.SKU
                    lblDiscountDescFreeAlt.Tag = Discount4.ID
                End If
            End If

            GetPricingForJumpRunItems()

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try


    End Sub

    Private Sub GetPromoPricingListItems()
        Dim x = RetrievePricingPromos(False, DisplayAll:=True)
        ComboBox1.DataSource = x
        ComboBox1.DisplayMember = "Value"
        ComboBox1.ValueMember = "Key"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim sb As New System.Text.StringBuilder
            Dim PricingProblem As Boolean = False

            '//Check that all TXtItemPrices1 item is not blank
            '//Check that JRPrice1 is set to a non zero value
            '//Ensure that the JR Price item actually exists in JR
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode1, JRPrice1, 1)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode2, JRPrice2, 2)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode3, JRPrice3, 3)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode4, JRPrice4, 4)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode5, JRPrice5, 5)

            If PricingProblem = False Then
                Dim Obj As New ClsPromoPricing
                With Obj
                    .ID = CurrentPromoID
                    .PromoDescription = ComboBox1.Text
                    .Status = ChkStatus.Checked
                    .DisplayInList = ChkDisplayInList.Checked
                    .DiscountCode = ""
                    .ItemCode1 = CInt(TxtJRItemCode1.Text)
                    .ItemCode2 = CInt(TxtJRItemCode2.Text)
                    .ItemCode3 = CInt(TxtJRItemCode3.Text)
                    .ItemCode4 = CInt(TxtJRItemCode4.Text)
                    .ItemCode5 = CInt(TxtJRItemCode5.Text)
                    .ItemPrice1 = CDbl(JRPrice1.Text)
                    .ItemPrice2 = CDbl(JRPrice2.Text)
                    .ItemPrice3 = CDbl(JRPrice3.Text)
                    .ItemPrice4 = CDbl(JRPrice4.Text)
                    .ItemPrice5 = CDbl(JRPrice5.Text)
                End With
                UpdatePromoPricing(Obj)
            Else
                MessageBox.Show("We have a problem with pricing" & Environment.NewLine & sb.ToString, "Pricing Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    'Private Function ValidateData() As Boolean
    '    Dim retval As Boolean = True

    '    If String.IsNullOrEmpty(TxtItemPrice1.Text) Or IsNumeric(TxtItemPrice1.Text) = False Then
    '        retval = False
    '        MsgBox("Item1 price is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtItemPrice2.Text) Or IsNumeric(TxtItemPrice2.Text) = False Then
    '        retval = False
    '        MsgBox("Item2 price is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtItemPrice3.Text) Or IsNumeric(TxtItemPrice3.Text) = False Then
    '        retval = False
    '        MsgBox("Item3 price is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtItemPrice4.Text) Or IsNumeric(TxtItemPrice4.Text) = False Then
    '        retval = False
    '        MsgBox("Item4 price is not valid")
    '    End If
    '    If String.IsNullOrEmpty(TxtItemPrice5.Text) Or IsNumeric(TxtItemPrice5.Text) = False Then
    '        retval = False
    '        MsgBox("Item5 price is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtJRItemCode1.Text) Or IsNumeric(TxtJRItemCode1.Text) = False Then
    '        retval = False
    '        MsgBox("Item1 JumpRun is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtJRItemCode2.Text) Or IsNumeric(TxtJRItemCode2.Text) = False Then
    '        retval = False
    '        MsgBox("Item2 JumpRun is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtJRItemCode3.Text) Or IsNumeric(TxtJRItemCode3.Text) = False Then
    '        retval = False
    '        MsgBox("Item3 JumpRun is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtJRItemCode4.Text) Or IsNumeric(TxtJRItemCode4.Text) = False Then
    '        retval = False
    '        MsgBox("Item4 JumpRun is not valid")
    '    End If

    '    If String.IsNullOrEmpty(TxtJRItemCode5.Text) Or IsNumeric(TxtJRItemCode5.Text) = False Then
    '        retval = False
    '        MsgBox("Item5 JumpRun is not valid")
    '    End If
    '    Return retval
    'End Function

    'Private Sub SavePricingInfo()
    '    Try
    '        Dim obj As New ClsPricing

    '        If Not String.IsNullOrEmpty(LblItemDescription1.Tag) Then
    '            With obj
    '                .ID = LblItemDescription1.Tag
    '                .Price = CDbl(TxtItemPrice1.Text)
    '                .JR_ItemID = CInt(TxtJRItemCode1.Text)
    '                .Discountable = ChkDiscountable1.Checked
    '                .SKU = LblItemSku1.Text
    '            End With
    '            UpdatePricing(obj, True)
    '        End If


    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(LblItemDescription2.Tag) Then
    '            With obj
    '                .ID = LblItemDescription2.Tag
    '                .Price = CDbl(TxtItemPrice2.Text)
    '                .JR_ItemID = CInt(TxtJRItemCode2.Text)
    '                .Discountable = ChkDiscountable2.Checked
    '                .SKU = LblItemSku2.Text
    '            End With
    '            UpdatePricing(obj, True)
    '        End If

    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(LblItemDescription3.Tag) Then
    '            With obj
    '                .ID = LblItemDescription3.Tag
    '                .Price = CDbl(TxtItemPrice3.Text)
    '                .JR_ItemID = CInt(TxtJRItemCode3.Text)
    '                .Discountable = ChkDiscountable3.Checked
    '                .SKU = LblItemSku3.Text
    '            End With
    '            UpdatePricing(obj, True)
    '        End If


    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(LblItemDescription4.Tag) Then
    '            With obj
    '                .ID = LblItemDescription4.Tag
    '                .Price = CDbl(TxtItemPrice4.Text)
    '                .JR_ItemID = CInt(TxtJRItemCode4.Text)
    '                .Discountable = ChkDiscountable4.Checked
    '                .SKU = LblItemSku4.Text
    '            End With
    '            UpdatePricing(obj, True)
    '        End If


    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(LblItemDescription5.Tag) Then
    '            With obj
    '                .ID = LblItemDescription5.Tag
    '                .Price = CDbl(TxtItemPrice5.Text)
    '                .JR_ItemID = CInt(TxtJRItemCode5.Text)
    '                .Discountable = ChkDiscountable5.Checked
    '                .SKU = LblItemSku5.Text
    '            End With
    '            UpdatePricing(obj, True)
    '        End If



    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(TxtDiscountSku1.Text) Then
    '            With obj
    '                .ID = lblDiscountDesc1.Tag
    '                .Price = CDbl(TxtDiscountCode1.Text)
    '                .JR_ItemID = CInt(TxtJRItemDiscountCode1.Text)
    '                .SKU = TxtDiscountSku1.Text
    '            End With
    '            UpdatePricing(obj, False)
    '        End If

    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(TxtDiscountSku2.Text) Then
    '            With obj
    '                .ID = lblDiscountDesc2.Tag
    '                .Price = CDbl(TxtDiscountCode2.Text)
    '                .JR_ItemID = CInt(TxtJRItemDiscountCode2.Text)
    '                .SKU = TxtDiscountSku2.Text
    '            End With
    '            UpdatePricing(obj, False)
    '        End If


    '        obj = New ClsPricing

    '        If Not String.IsNullOrEmpty(TxtDiscountSku3.Text) Then
    '            With obj
    '                .ID = lblDiscountDesc3.Tag
    '                .Price = CDbl(TxtDiscountCode3.Text)
    '                .JR_ItemID = CInt(TxtJRItemDiscountCode3.Text)
    '                .SKU = TxtDiscountSku3.Text
    '            End With
    '            UpdatePricing(obj, False)
    '        End If
    '    Catch ex As Exception
    '        Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
    '        Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
    '        LogError(methodName, ex)
    '    End Try
    'End Sub

    'Sub SavePromoPricing()
    '    Dim ObjPromo As New ClsPromoPricing

    '    With ObjPromo
    '        .ID = -1
    '        .PromoDescription = "New Item"
    '        .ItemCode1 = CInt(TxtJRItemCode1.Text)
    '        .ItemPrice1 = CDbl(JRPrice1.Text)
    '        .ItemCode2 = CInt(TxtJRItemCode2.Text)
    '        .ItemPrice2 = CDbl(JRPrice2.Text)
    '        .ItemCode3 = CInt(TxtJRItemCode3.Text)
    '        .ItemPrice3 = CDbl(JRPrice3.Text)
    '        .ItemCode4 = CInt(TxtJRItemCode4.Text)
    '        .ItemPrice4 = CDbl(JRPrice4.Text)
    '        .ItemCode5 = CInt(TxtJRItemCode5.Text)
    '        .ItemPrice5 = CDbl(JRPrice5.Text)
    '        .Status = CheckBox2.Checked

    '    End With

    '    InsertPromoPricing(ObjPromo)

    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            GetPricingForJumpRunItems()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub GetPricingForJumpRunItems()
        Try
            'for each of the items get the price.
            If IsNumeric(TxtJRItemCode1.Text) Then
                Dim p1 = GetJumpRunItemPrice(CInt(TxtJRItemCode1.Text))
                JRPrice1.Text = p1.ToString
            Else
                JRPrice1.Text = ""
            End If

            If IsNumeric(TxtJRItemCode2.Text) Then
                Dim p2 = GetJumpRunItemPrice(CInt(TxtJRItemCode2.Text))
                JRPrice2.Text = p2.ToString
            Else
                JRPrice2.Text = ""
            End If
            If IsNumeric(TxtJRItemCode3.Text) Then
                Dim p3 = GetJumpRunItemPrice(CInt(TxtJRItemCode3.Text))
                JRPrice3.Text = p3.ToString
            Else
                JRPrice3.Text = ""
            End If
            If IsNumeric(TxtJRItemCode4.Text) Then
                Dim p4 = GetJumpRunItemPrice(CInt(TxtJRItemCode4.Text))
                JRPrice4.Text = p4.ToString
            Else
                JRPrice4.Text = ""
            End If
            If IsNumeric(TxtJRItemCode5.Text) Then
                Dim p5 = GetJumpRunItemPrice(CInt(TxtJRItemCode5.Text))
                JRPrice5.Text = p5.ToString
            Else
                JRPrice5.Text = ""
            End If


            If IsNumeric(TxtJRItemDiscountCode1.Text) Then
                Dim pd1 = GetJumpRunItemPrice(CInt(TxtJRItemDiscountCode1.Text))
                JRDiscPrice1.Text = pd1.ToString
            Else
                JRDiscPrice1.Text = ""
            End If


            If IsNumeric(TxtJRItemDiscountCode2.Text) Then
                Dim pd2 = GetJumpRunItemPrice(CInt(TxtJRItemDiscountCode2.Text))
                JRDiscPrice2.Text = pd2.ToString
            Else
                JRDiscPrice1.Text = ""
            End If


            If IsNumeric(TxtJRItemDiscountCode3.Text) Then
                Dim pd3 = GetJumpRunItemPrice(CInt(TxtJRItemDiscountCode3.Text))
                JRDiscPrice3.Text = pd3.ToString
            Else
                JRDiscPrice3.Text = ""
            End If



            If IsNumeric(TxtJRItemDiscountCodeFreeAlt.Text) Then
                Dim pd4 = GetJumpRunItemPrice(CInt(TxtJRItemDiscountCodeFreeAlt.Text))
                JRDiscPriceFreeAlt.Text = pd4.ToString
            Else
                JRDiscPriceFreeAlt.Text = ""
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub BtnSelect1_Click(sender As Object, e As EventArgs) Handles BtnSelect1.Click
        '//We need to display a list of Items which the user can select
        Dim obj As New FrmJumprunItemsSelect
        obj.ShowDialog()
        If obj.CurrentSelectedItem <> -1 Then
            TxtItemPrice1.Text = obj.CurrentPrice
            TxtJRItemCode1.Text = obj.CurrentSelectedItem
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            ''//We need to determine that the Jumprun and web store pricing items match.  If they don't then we will have problems.
            Dim sb As New System.Text.StringBuilder
            Dim PricingProblem As Boolean = False

            '//Check that all TXtItemPrices1 item is not blank
            '//Check that JRPrice1 is set to a non zero value
            '//Ensure that the JR Price item actually exists in JR
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode1, JRPrice1, 1)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode2, JRPrice2, 2)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode3, JRPrice3, 3)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode4, JRPrice4, 4)
            PricingProblem = ValidatePriceItem(sb, PricingProblem, TxtJRItemCode5, JRPrice5, 5)

            If PricingProblem = True Then
                MessageBox.Show("We have a problem with pricing" & Environment.NewLine & sb.ToString, "Pricing Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Pricing Validation Is Successful" & Environment.NewLine & sb.ToString, "Pricing Validation", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Function ValidatePriceItem(ByRef sb As StringBuilder, PricingProblem As Boolean, TxtItemPrice As TextBox, JRPrice As Label, index As Integer) As Boolean
        If TxtItemPrice.Text.Trim = String.Empty Then
            PricingProblem = True
            sb.AppendLine(String.Format("No Item set for  Item {0}", index))
        End If
        If JRPrice.Text.Trim = String.Empty Or CDbl(JRPrice.Text) = 0 Then
            PricingProblem = True
            sb.AppendLine(String.Format("No Price set for  Item {0}", index))
        End If
        If IsNumeric(TxtItemPrice.Text.Trim) Then
            Dim pfjr1 = GetJumpRunItemPrice(CInt(TxtItemPrice.Text.Trim))
            If pfjr1 <> CDbl(JRPrice.Text.Trim) Then
                PricingProblem = True
                sb.AppendLine(String.Format("Price set is different from current price in Jumprun for this item {0} - {1}", pfjr1, JRPrice.Text))
            End If
        Else
            PricingProblem = True
            sb.AppendLine(String.Format("Non Numeric code set for item {0}", index))
        End If


        Return PricingProblem
    End Function

    Private Sub BtnSelect2_Click(sender As Object, e As EventArgs) Handles BtnSelect2.Click
        '//We need to display a list of Items which the user can select
        Dim obj As New FrmJumprunItemsSelect
        obj.ShowDialog()
        If obj.CurrentSelectedItem <> -1 Then
            TxtItemPrice2.Text = obj.CurrentPrice
            TxtJRItemCode2.Text = obj.CurrentSelectedItem
        End If
    End Sub

    Private Sub BtnSelect3_Click(sender As Object, e As EventArgs) Handles BtnSelect3.Click
        '//We need to display a list of Items which the user can select
        Dim obj As New FrmJumprunItemsSelect
        obj.ShowDialog()
        If obj.CurrentSelectedItem <> -1 Then
            TxtItemPrice3.Text = obj.CurrentPrice
            TxtJRItemCode3.Text = obj.CurrentSelectedItem
        End If
    End Sub

    Private Sub BtnSelect4_Click(sender As Object, e As EventArgs) Handles BtnSelect4.Click
        '//We need to display a list of Items which the user can select
        Dim obj As New FrmJumprunItemsSelect
        obj.ShowDialog()
        If obj.CurrentSelectedItem <> -1 Then
            TxtItemPrice4.Text = obj.CurrentPrice
            TxtJRItemCode4.Text = obj.CurrentSelectedItem
        End If
    End Sub

    Private Sub BtnSelect5_Click(sender As Object, e As EventArgs) Handles BtnSelect5.Click
        '//We need to display a list of Items which the user can select
        Dim obj As New FrmJumprunItemsSelect
        obj.ShowDialog()
        If obj.CurrentSelectedItem <> -1 Then
            TxtItemPrice5.Text = obj.CurrentPrice
            TxtJRItemCode5.Text = obj.CurrentSelectedItem
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            CurrentPromoID = CType(ComboBox1.SelectedItem, KeyValuePair(Of Integer, String)).Key

            Dim x = RetrievePricingPromoForId(CurrentPromoID)

            '//So lets populate the Items in the GroupBox and then get the latest from JumprUN
            TxtJRItemCode1.Text = x.ItemCode1.ToString
            TxtJRItemCode2.Text = x.ItemCode2.ToString
            TxtJRItemCode3.Text = x.ItemCode3.ToString
            TxtJRItemCode4.Text = x.ItemCode4.ToString
            TxtJRItemCode5.Text = x.ItemCode5.ToString

            JRPrice1.Text = x.ItemPrice1.ToString
            JRPrice2.Text = x.ItemPrice2.ToString
            JRPrice3.Text = x.ItemPrice3.ToString
            JRPrice4.Text = x.ItemPrice4.ToString
            JRPrice5.Text = x.ItemPrice5.ToString

            ChkStatus.Checked = x.Status
            ChkDisplayInList.Checked = x.DisplayInList

            GetPricingForJumpRunItems()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim x As New FrmNewPromo
            If x.ShowDialog = DialogResult.OK Then
                '//Get standard Pricing and create a new Promo Item
                Dim defaultPromo = RetrievePricingPromoForId(1)
                Dim Obj As New ClsPromoPricing
                With Obj
                    .ID = -1
                    .PromoDescription = x.PromoName
                    .Status = 1
                    .DisplayInList = 1
                    .DiscountCode = ""
                    .ItemCode1 = defaultPromo.ItemCode1
                    .ItemCode2 = defaultPromo.ItemCode2
                    .ItemCode3 = defaultPromo.ItemCode3
                    .ItemCode4 = defaultPromo.ItemCode4
                    .ItemCode5 = defaultPromo.ItemCode5
                    .ItemPrice1 = defaultPromo.ItemPrice1
                    .ItemPrice2 = defaultPromo.ItemPrice2
                    .ItemPrice3 = defaultPromo.ItemPrice3
                    .ItemPrice4 = defaultPromo.ItemPrice4
                    .ItemPrice5 = defaultPromo.ItemPrice5
                End With
                InsertPromoPricing(Obj)
                GetPromoPricingListItems()
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Sub
End Class