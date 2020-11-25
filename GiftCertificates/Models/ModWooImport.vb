Imports System.Reflection

Public Module ModWooImport

    Const TypeToIncludeForImport As String = "processing"
    Public Sub ImportWooCSVFile(filename As String)
        Dim sb As New System.Text.StringBuilder
        Dim NotesSB As New System.Text.StringBuilder

        Dim i As Integer = 0
        Dim irow = 0
        Dim InvalidRecords As Integer = 0

        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(
                        filename)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()


            Dim LstOrderLineItems As New List(Of ClsWooLineItem)
            While Not MyReader.EndOfData
                Try
                    Dim ObjLineItem As New ClsWooLineItem
                    currentRow = MyReader.ReadFields()

                    If irow = 0 Then
                        'First row is a header row
                        'My.Computer.FileSystem.WriteAllText("fields.txt", sb.ToString, False)
                        'MsgBox(sb.ToString)
                    Else
                        'Only Process - processing status records
                        If String.IsNullOrEmpty(currentRow(1)) = False Then
                            If currentRow(1).Trim.ToLower = TypeToIncludeForImport Then
                                If String.IsNullOrEmpty(currentRow(0)) = False Then
                                    ObjLineItem.OrderID = currentRow(0)
                                End If

                                If String.IsNullOrEmpty(currentRow(2)) = False Then
                                    ObjLineItem.OrderDate = currentRow(2)
                                End If

                                If String.IsNullOrEmpty(currentRow(3)) = False Then
                                    ObjLineItem.PaymentMethod = currentRow(3)
                                End If
                                If String.IsNullOrEmpty(currentRow(3)) = False Then
                                    ObjLineItem.PaymentMethod = currentRow(3)
                                End If
                                If String.IsNullOrEmpty(currentRow(5)) = False Then
                                    ObjLineItem.Shippingmethod = currentRow(5)
                                End If

                                If String.IsNullOrEmpty(currentRow(6)) = False Then
                                    ObjLineItem.ShippingFirstName = currentRow(6)
                                End If
                                If String.IsNullOrEmpty(currentRow(7)) = False Then
                                    ObjLineItem.ShippingLastName = currentRow(7)
                                End If
                                If String.IsNullOrEmpty(currentRow(8)) = False Then
                                    ObjLineItem.ShippingAddress1 = currentRow(8)
                                End If
                                If String.IsNullOrEmpty(currentRow(9)) = False Then
                                    ObjLineItem.ShippingAddress2 = currentRow(9)
                                End If

                                If String.IsNullOrEmpty(currentRow(10)) = False Then
                                    ObjLineItem.ShippingCity = currentRow(10)
                                End If
                                If String.IsNullOrEmpty(currentRow(11)) = False Then
                                    ObjLineItem.ShippingState = currentRow(11)
                                End If

                                If String.IsNullOrEmpty(currentRow(12)) = False Then
                                    ObjLineItem.ShippingZip = currentRow(12)
                                End If
                                If String.IsNullOrEmpty(currentRow(13)) = False Then
                                    ObjLineItem.ShippingCountry = currentRow(13)
                                End If

                                If String.IsNullOrEmpty(currentRow(14)) = False Then
                                    ObjLineItem.BillingFirstName = currentRow(14)
                                End If
                                If String.IsNullOrEmpty(currentRow(15)) = False Then
                                    ObjLineItem.BillingLastName = currentRow(15)
                                End If
                                If String.IsNullOrEmpty(currentRow(16)) = False Then
                                    ObjLineItem.BillingAddress1 = currentRow(16)
                                End If
                                If String.IsNullOrEmpty(currentRow(17)) = False Then
                                    ObjLineItem.BillingAddress2 = currentRow(17)
                                End If

                                If String.IsNullOrEmpty(currentRow(18)) = False Then
                                    ObjLineItem.BillingCity = currentRow(18)
                                End If
                                If String.IsNullOrEmpty(currentRow(19)) = False Then
                                    ObjLineItem.BillingState = currentRow(19)
                                End If

                                If String.IsNullOrEmpty(currentRow(20)) = False Then
                                    ObjLineItem.BillingZip = currentRow(20)
                                End If
                                If String.IsNullOrEmpty(currentRow(21)) = False Then
                                    ObjLineItem.BillingCountry = currentRow(21)
                                End If

                                If String.IsNullOrEmpty(currentRow(22)) = False Then
                                    ObjLineItem.BillingPhone = currentRow(22)
                                End If
                                If String.IsNullOrEmpty(currentRow(23)) = False Then
                                    ObjLineItem.BillingEmail = currentRow(23)
                                End If

                                If String.IsNullOrEmpty(currentRow(24)) = False Then
                                    ObjLineItem.BillingOrderComments = currentRow(24)
                                End If

                                'Quantity of items purchased	
                                'Product Name	
                                'Product SKU	
                                'Product ID	
                                'Item price INCL. tax	
                                'Coupon Code	
                                'Order Discount	
                                'Order Total(Auth.net)	
                                'Paid Date	
                                'Transaction ID
                                If String.IsNullOrEmpty(currentRow(25)) = False Then
                                    ObjLineItem.LineitemQuantityPurchased = currentRow(25)
                                End If
                                If String.IsNullOrEmpty(currentRow(26)) = False Then
                                    ObjLineItem.LineitemName = currentRow(26)
                                End If


                                Dim SKU = DetermineSKUFromDescription(ObjLineItem.LineitemName)

                                If String.IsNullOrEmpty(SKU) = False Then
                                    ObjLineItem.LineitemSKU = SKU
                                End If
                                'If String.IsNullOrEmpty(currentRow(27)) = False Then
                                '    ObjLineItem.Pro = currentRow(27)
                                'End If
                                If String.IsNullOrEmpty(currentRow(28)) = False Then
                                    ObjLineItem.LineitemPrice = currentRow(28)
                                End If
                                If String.IsNullOrEmpty(currentRow(29)) = False Then
                                    ObjLineItem.LineitemCouponCode = currentRow(29)
                                End If
                                If String.IsNullOrEmpty(currentRow(30)) = False Then
                                    ObjLineItem.OrderDiscount = currentRow(30)
                                End If
                                If String.IsNullOrEmpty(currentRow(31)) = False Then
                                    ObjLineItem.OrderTotal = currentRow(31)
                                End If

                                If String.IsNullOrEmpty(currentRow(32)) = False Then
                                    ObjLineItem.PaidDate = currentRow(32)
                                End If
                                If String.IsNullOrEmpty(currentRow(33)) = False Then
                                    ObjLineItem.TransactionID = currentRow(33)
                                End If

                                If String.IsNullOrEmpty(currentRow(34)) = False Then
                                    ObjLineItem.Notes = "Import From Woo webstore" & Environment.NewLine & currentRow(34)
                                Else
                                    ObjLineItem.Notes = "Import From Woo webstore"
                                End If


                                LstOrderLineItems.Add(ObjLineItem)
                            Else
                                InvalidRecords += 1
                            End If
                        End If

                    End If

                        irow += 1
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                    "is not valid and will be skipped.")
                End Try
            End While


            '//Process the records and write to gift certificate table in 
            'So there could be multiple line items and process each.  ie. one for tandem. one for video.
            'These need to be combined so that they can be written as one record.
            Dim LstGCOrderQueue As New List(Of ClsGiftCertificate2)


            'Process Line Items into certificates
            Dim LastId As String = ""
            Dim Index = 0
            Dim ObjGiftCertificate As New ClsGiftCertificate2
            For Each lineitem In LstOrderLineItems
                Try
                    If lineitem.OrderID <> LastId AndAlso Index <> 0 Then
                        'Write Certificate to database

                        LstGCOrderQueue.Add(ObjGiftCertificate)
                        ObjGiftCertificate = New ClsGiftCertificate2
                    End If

                    If String.IsNullOrEmpty(lineitem.OrderID) = False Then
                        ObjGiftCertificate.OrderId = lineitem.OrderID
                    End If

                    ObjGiftCertificate.Purchaser_FirstName = lineitem.BillingFirstName
                    If String.IsNullOrEmpty(lineitem.BillingFirstName) = False Then
                        ObjGiftCertificate.Purchaser_FirstName = lineitem.BillingFirstName
                    End If

                    If String.IsNullOrEmpty(lineitem.BillingLastName) = False Then
                        ObjGiftCertificate.Purchaser_LastName = lineitem.BillingLastName
                    End If


                    If String.IsNullOrEmpty(lineitem.BillingAddress1) = False Then
                        ObjGiftCertificate.Billing_Address.Address1 = lineitem.BillingAddress1
                    End If
                    If String.IsNullOrEmpty(lineitem.BillingAddress2) = False Then
                        ObjGiftCertificate.Billing_Address.Address2 = lineitem.BillingAddress2
                    End If
                    If String.IsNullOrEmpty(lineitem.BillingCity) = False Then
                        ObjGiftCertificate.Billing_Address.City = lineitem.BillingCity
                    End If

                    If String.IsNullOrEmpty(lineitem.BillingState) = False Then
                        ObjGiftCertificate.Billing_Address.State = lineitem.BillingState
                    End If
                    If String.IsNullOrEmpty(lineitem.BillingCountry) = False Then
                        ObjGiftCertificate.Billing_Address.Country = lineitem.BillingCountry
                    End If
                    If String.IsNullOrEmpty(lineitem.BillingZip) = False Then
                        ObjGiftCertificate.Billing_Address.Zip = lineitem.BillingZip
                    End If

                    If String.IsNullOrEmpty(lineitem.BillingPhone) = False Then
                        ObjGiftCertificate.Billing_Address.Phone1 = lineitem.BillingPhone
                    End If

                    If String.IsNullOrEmpty(lineitem.BillingEmail) = False Then
                        ObjGiftCertificate.Billing_Address.Email = lineitem.BillingEmail
                    End If

                    If String.IsNullOrEmpty(lineitem.ShippingAddress1) = False Then
                        ObjGiftCertificate.Shipping_Address.Address1 = lineitem.ShippingAddress1
                    End If
                    If String.IsNullOrEmpty(lineitem.ShippingAddress2) = False Then
                        ObjGiftCertificate.Shipping_Address.Address2 = lineitem.ShippingAddress2
                    End If
                    If String.IsNullOrEmpty(lineitem.ShippingCity) = False Then
                        ObjGiftCertificate.Shipping_Address.City = lineitem.ShippingCity
                    End If

                    If String.IsNullOrEmpty(lineitem.ShippingState) = False Then
                        ObjGiftCertificate.Shipping_Address.State = lineitem.ShippingState
                    End If
                    If String.IsNullOrEmpty(lineitem.ShippingCountry) = False Then
                        ObjGiftCertificate.Shipping_Address.Country = lineitem.ShippingCountry
                    End If
                    If String.IsNullOrEmpty(lineitem.ShippingZip) = False Then
                        ObjGiftCertificate.Shipping_Address.Zip = lineitem.ShippingZip
                    End If

                    If String.IsNullOrEmpty(lineitem.Shippingmethod) = False Then
                        ObjGiftCertificate.delivery = DetermineShippingMethodFromString(lineitem.Shippingmethod)
                    End If

                    If String.IsNullOrEmpty(lineitem.PaidDate) = False Then
                        ObjGiftCertificate.GC_DateEntered = DateTime.Parse(lineitem.PaidDate)
                    End If

                    ObjGiftCertificate.PaymentMethod = PaymentMethod.Online_authorize_net_cim_credit_card

                    '//These Replicate for all records on the order
                    If String.IsNullOrEmpty(lineitem.OrderTotal) = False Then
                        ObjGiftCertificate.GC_TotalAmount = lineitem.OrderTotal
                    End If
                    If String.IsNullOrEmpty(lineitem.OrderDiscount) = False Then
                        ObjGiftCertificate.GC_TotalDiscount = lineitem.OrderDiscount
                    End If

                    If String.IsNullOrEmpty(lineitem.LineitemCouponCode) = False Then
                        ObjGiftCertificate.GC_DiscountCode = lineitem.LineitemCouponCode
                    End If

                    Dim Notes As String = ObjGiftCertificate.Notes & Environment.NewLine
                    If String.IsNullOrEmpty(lineitem.BillingOrderComments) = False Then
                        Notes = Notes & lineitem.BillingOrderComments
                    End If
                    ObjGiftCertificate.Notes = Notes
                    ObjGiftCertificate.PointOfSale = PointOfSale.Online

                    '//Places Quantity into product slots
                    '//TODO: Look at how this is possible from the entries
                    If String.IsNullOrEmpty(lineitem.LineitemSKU) = False Then
                        Select Case lineitem.LineitemSKU.Trim
                            Case "TDM10K"
                                ObjGiftCertificate.Item1.Quantity = lineitem.LineitemQuantityPurchased
                            Case "TDM12K"
                                ObjGiftCertificate.Item2.Quantity = lineitem.LineitemQuantityPurchased
                            Case "TDM10KVID"
                                ObjGiftCertificate.Item3.Quantity = lineitem.LineitemQuantityPurchased
                            Case "TDM12KVID"
                                ObjGiftCertificate.Item4.Quantity = lineitem.LineitemQuantityPurchased
                            Case "VID"
                                ObjGiftCertificate.Item5.Quantity = lineitem.LineitemQuantityPurchased
                        End Select
                    End If

                    ObjGiftCertificate.Online_OrderNumber = lineitem.OrderID
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                LastId = lineitem.OrderID
                Index = Index + 1
            Next

            If Index > 0 Then
                LstGCOrderQueue.Add(ObjGiftCertificate)
            End If


            '//Actually Write the Certificates to the Database if needed
            Try
                Dim iOrderCounter As Integer = 0
                Dim duplicateCertificatesFound As Integer = 0
                Dim newCertificatesCreated As Integer = 0

                For Each gc In LstGCOrderQueue
                    If DoesGCOrderExist(gc.Online_OrderNumber) = False Then
                        CreateGiftCertificatesFromCertificateRecord(gc)
                        newCertificatesCreated += 1
                    Else
                        duplicateCertificatesFound += 1
                    End If
                    iOrderCounter += 1
                Next

                '//Display a final message detailing records imported
                Dim sb1 As New System.Text.StringBuilder

                sb1.AppendLine(String.Format("{0} Order records are found.", LstGCOrderQueue.Count + InvalidRecords))
                sb1.AppendLine(String.Format("{0} Order records are Invalid Status.", InvalidRecords))
                sb1.AppendLine(String.Format("{0} will inserted into the queue.", newCertificatesCreated))
                sb1.AppendLine(String.Format("{0} are duplicates of record already in the queue.", duplicateCertificatesFound))
                MessageBox.Show(sb1.ToString, "Import Results", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
                Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
                LogError(methodName, ex)
            End Try


        End Using
    End Sub

    Function DetermineSKUFromDescription(lineitemName As String) As Object
        Dim COntainsTandem = False
        Dim Contains30Sec = False
        Dim Contains60Sec = False
        Dim ContainsVid = False
        Dim ProductSKU As String = ""

        Try
            Dim Desc As String = lineitemName.ToUpper.Trim
            COntainsTandem = Desc.Contains("TANDEM")
            Contains30Sec = Desc.Contains("30 SEC")
            Contains60Sec = Desc.Contains("60 SEC")
            ContainsVid = Desc.Contains("VID")

            If Contains30Sec And Contains60Sec Then
                Throw New Exception("Problem with Calculating SKU From Description")
            End If

            If COntainsTandem = False Then
                If ContainsVid Then
                    ProductSKU = "VID"
                End If
            Else
                '//Its a tandem
                If Contains30Sec AndAlso ContainsVid Then
                    ProductSKU = "TDM10KVID"
                ElseIf Contains30Sec AndAlso ContainsVid = False Then
                    ProductSKU = "TDM10K "
                ElseIf Contains60Sec AndAlso ContainsVid Then
                    ProductSKU = "TDM12KVID"
                ElseIf Contains60Sec AndAlso ContainsVid = False Then
                    ProductSKU = "TDM12K "
                End If
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
        Return ProductSKU
    End Function

    Function DetermineShippingMethodFromString(strShippingString As String) As DeliveryOptions
        Dim strCleanString = strShippingString.ToString.Trim
        Dim retvalue As DeliveryOptions = DeliveryOptions.USMail

        Try

            If strCleanString.StartsWith("EMAIL") Then
                retvalue = DeliveryOptions.Email
            ElseIf strCleanString.StartsWith("WILL CALL") Then
                retvalue = DeliveryOptions.InOffice
            ElseIf strCleanString.StartsWith("US MAIL") Then
                If strCleanString.Contains("DISCREET") Then
                    retvalue = DeliveryOptions.USDiscreet
                Else
                    retvalue = DeliveryOptions.USMail
                End If
            Else
                If strCleanString.StartsWith("PICKUP") Then
                    retvalue = DeliveryOptions.InOffice
                End If
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

        Return retvalue
    End Function

End Module
