Imports System.Reflection

Public Class FrmPricing
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmPricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
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
            If dc.Count > 3 Then
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

            End If

            GetPricingForJumpRunItems()

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bln As Boolean = ValidateData()

        If bln = True Then
            SavePricingInfo()
        End If

    End Sub

    Private Function ValidateData() As Boolean
        Dim retval As Boolean = True

        If String.IsNullOrEmpty(TxtItemPrice1.Text) Or IsNumeric(TxtItemPrice1.Text) = False Then
            retval = False
            MsgBox("Item1 price is not valid")
        End If

        If String.IsNullOrEmpty(TxtItemPrice2.Text) Or IsNumeric(TxtItemPrice2.Text) = False Then
            retval = False
            MsgBox("Item2 price is not valid")
        End If

        If String.IsNullOrEmpty(TxtItemPrice3.Text) Or IsNumeric(TxtItemPrice3.Text) = False Then
            retval = False
            MsgBox("Item3 price is not valid")
        End If

        If String.IsNullOrEmpty(TxtItemPrice4.Text) Or IsNumeric(TxtItemPrice4.Text) = False Then
            retval = False
            MsgBox("Item4 price is not valid")
        End If
        If String.IsNullOrEmpty(TxtItemPrice5.Text) Or IsNumeric(TxtItemPrice5.Text) = False Then
            retval = False
            MsgBox("Item5 price is not valid")
        End If

        If String.IsNullOrEmpty(TxtJRItemCode1.Text) Or IsNumeric(TxtJRItemCode1.Text) = False Then
            retval = False
            MsgBox("Item1 JumpRun is not valid")
        End If

        If String.IsNullOrEmpty(TxtJRItemCode2.Text) Or IsNumeric(TxtJRItemCode2.Text) = False Then
            retval = False
            MsgBox("Item2 JumpRun is not valid")
        End If

        If String.IsNullOrEmpty(TxtJRItemCode3.Text) Or IsNumeric(TxtJRItemCode3.Text) = False Then
            retval = False
            MsgBox("Item3 JumpRun is not valid")
        End If

        If String.IsNullOrEmpty(TxtJRItemCode4.Text) Or IsNumeric(TxtJRItemCode4.Text) = False Then
            retval = False
            MsgBox("Item4 JumpRun is not valid")
        End If

        If String.IsNullOrEmpty(TxtJRItemCode5.Text) Or IsNumeric(TxtJRItemCode5.Text) = False Then
            retval = False
            MsgBox("Item5 JumpRun is not valid")
        End If
        Return retval
    End Function

    Private Sub SavePricingInfo()
        Try
            Dim obj As New ClsPricing

            If Not String.IsNullOrEmpty(LblItemDescription1.Tag) Then
                With obj
                    .ID = LblItemDescription1.Tag
                    .Price = CDbl(TxtItemPrice1.Text)
                    .JR_ItemID = CInt(TxtJRItemCode1.Text)
                    .Discountable = ChkDiscountable1.Checked
                    .SKU = LblItemSku1.Text
                End With
                UpdatePricing(obj, True)
            End If


            obj = New ClsPricing

            If Not String.IsNullOrEmpty(LblItemDescription2.Tag) Then
                With obj
                    .ID = LblItemDescription2.Tag
                    .Price = CDbl(TxtItemPrice2.Text)
                    .JR_ItemID = CInt(TxtJRItemCode2.Text)
                    .Discountable = ChkDiscountable2.Checked
                    .SKU = LblItemSku2.Text
                End With
                UpdatePricing(obj, True)
            End If

            obj = New ClsPricing

            If Not String.IsNullOrEmpty(LblItemDescription3.Tag) Then
                With obj
                    .ID = LblItemDescription3.Tag
                    .Price = CDbl(TxtItemPrice3.Text)
                    .JR_ItemID = CInt(TxtJRItemCode3.Text)
                    .Discountable = ChkDiscountable3.Checked
                    .SKU = LblItemSku3.Text
                End With
                UpdatePricing(obj, True)
            End If


            obj = New ClsPricing

            If Not String.IsNullOrEmpty(LblItemDescription4.Tag) Then
                With obj
                    .ID = LblItemDescription4.Tag
                    .Price = CDbl(TxtItemPrice4.Text)
                    .JR_ItemID = CInt(TxtJRItemCode4.Text)
                    .Discountable = ChkDiscountable4.Checked
                    .SKU = LblItemSku4.Text
                End With
                UpdatePricing(obj, True)
            End If


            obj = New ClsPricing

            If Not String.IsNullOrEmpty(LblItemDescription5.Tag) Then
                With obj
                    .ID = LblItemDescription5.Tag
                    .Price = CDbl(TxtItemPrice5.Text)
                    .JR_ItemID = CInt(TxtJRItemCode5.Text)
                    .Discountable = ChkDiscountable5.Checked
                    .SKU = LblItemSku5.Text
                End With
                UpdatePricing(obj, True)
            End If



            obj = New ClsPricing

            If Not String.IsNullOrEmpty(TxtDiscountSku1.Text) Then
                With obj
                    .ID = lblDiscountDesc1.Tag
                    .Price = CDbl(TxtDiscountCode1.Text)
                    .JR_ItemID = CInt(TxtJRItemDiscountCode1.Text)
                    .SKU = TxtDiscountSku1.Text
                End With
                UpdatePricing(obj, False)
            End If

            obj = New ClsPricing

            If Not String.IsNullOrEmpty(TxtDiscountSku2.Text) Then
                With obj
                    .ID = lblDiscountDesc2.Tag
                    .Price = CDbl(TxtDiscountCode2.Text)
                    .JR_ItemID = CInt(TxtJRItemDiscountCode2.Text)
                    .SKU = TxtDiscountSku2.Text
                End With
                UpdatePricing(obj, False)
            End If


            obj = New ClsPricing

            If Not String.IsNullOrEmpty(TxtDiscountSku3.Text) Then
                With obj
                    .ID = lblDiscountDesc3.Tag
                    .Price = CDbl(TxtDiscountCode3.Text)
                    .JR_ItemID = CInt(TxtJRItemDiscountCode3.Text)
                    .SKU = TxtDiscountSku3.Text
                End With
                UpdatePricing(obj, False)
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GetPricingForJumpRunItems()
    End Sub

    Private Sub GetPricingForJumpRunItems()
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
        '//We need to determine that the Jumprun and web store pricing items match.  If they don't then we will have problems.
        Dim sb As New System.Text.StringBuilder
        Dim PricingProblem As Boolean = False

        If CDbl(TxtItemPrice1.Text) <> CDbl(JRPrice1.Text) Then
            sb.AppendLine("Discrepency between WebStore and Jumprun item prices for Item 1")
            PricingProblem = True

        End If
        If CDbl(TxtItemPrice2.Text) <> CDbl(JRPrice2.Text) Then
            sb.AppendLine("Discrepency between WebStore and Jumprun item prices for Item 2")
            PricingProblem = True
        End If
        If CDbl(TxtItemPrice3.Text) <> CDbl(JRPrice3.Text) Then
            sb.AppendLine("Discrepency between WebStore and Jumprun item prices for Item 3")
            PricingProblem = True
        End If

        If CDbl(TxtItemPrice4.Text) <> CDbl(JRPrice4.Text) Then
            sb.AppendLine("Discrepency between WebStore and Jumprun item prices for Item 4")
            PricingProblem = True
        End If
        If CDbl(TxtItemPrice5.Text) <> CDbl(JRPrice5.Text) Then
            sb.AppendLine("Discrepency between WebStore and Jumprun item prices for Item 5")
            PricingProblem = True
        End If

        If PricingProblem = True Then
            MessageBox.Show("We have a problem with pricing" & Environment.NewLine & sb.ToString, "Pricing Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

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
End Class