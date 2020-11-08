Public Module ModShipify
    Public Sub ImportShopifyCSVFile(filename As String)
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer = 0
        Dim irow = 0

        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(
                        filename)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()


            Dim LstShopifyItems As New List(Of ClsShopifyLineItem)
            While Not MyReader.EndOfData
                Try
                    Dim ObjLineItem As New ClsShopifyLineItem

                    currentRow = MyReader.ReadFields()


                    If irow = 0 Then
                        'First row is a header row
                        'My.Computer.FileSystem.WriteAllText("fields.txt", sb.ToString, False)
                        'MsgBox(sb.ToString)
                    Else
                        If String.IsNullOrEmpty(currentRow(0)) = False Then
                            ObjLineItem.Name = currentRow(0)
                        End If
                        If String.IsNullOrEmpty(currentRow(1)) = False Then
                            ObjLineItem.BillingEmail = currentRow(1)
                        End If

                        If String.IsNullOrEmpty(currentRow(24)) = False Then
                            ObjLineItem.BillingName = currentRow(24)
                        End If
                        If String.IsNullOrEmpty(currentRow(26)) = False Then
                            ObjLineItem.BillingAddress1 = currentRow(26)
                        End If
                        If String.IsNullOrEmpty(currentRow(27)) = False Then
                            ObjLineItem.BillingAddress2 = currentRow(27)
                        End If
                        If String.IsNullOrEmpty(currentRow(29)) = False Then
                            ObjLineItem.BillingCity = currentRow(29)
                        End If
                        If String.IsNullOrEmpty(currentRow(31)) = False Then
                            ObjLineItem.BillingProvince = currentRow(31)
                        End If

                        If String.IsNullOrEmpty(currentRow(30)) = False Then
                            ObjLineItem.BillingZip = currentRow(30)
                        End If

                        If String.IsNullOrEmpty(currentRow(33)) = False Then
                            ObjLineItem.BillingPhone = currentRow(33)
                        End If



                        If String.IsNullOrEmpty(currentRow(17)) = False Then
                            ObjLineItem.LineitemName = currentRow(17)
                        End If
                        If String.IsNullOrEmpty(currentRow(18)) = False Then
                            ObjLineItem.LineitemPrice = currentRow(18)
                        End If

                        '// Determine
                        If String.IsNullOrEmpty(currentRow(16)) = False Then
                            ObjLineItem.LineitemQuantity = currentRow(16)
                        End If



                        If String.IsNullOrEmpty(currentRow(44)) = False Then
                            ObjLineItem.Notes = currentRow(44)
                        End If

                        If String.IsNullOrEmpty(currentRow(45)) = False Then
                            ObjLineItem.NotesAttributes = currentRow(45)
                        End If

                        If String.IsNullOrEmpty(currentRow(3)) = False Then
                            ObjLineItem.PaidAt = currentRow(3)
                        End If

                        If String.IsNullOrEmpty(currentRow(47)) = False Then
                            ObjLineItem.PaymentMethod = currentRow(47)
                        End If

                        If String.IsNullOrEmpty(currentRow(48)) = False Then
                            ObjLineItem.PaymentReference = currentRow(48)
                        End If
                        If String.IsNullOrEmpty(currentRow(8)) = False Then
                            ObjLineItem.Subtotal = currentRow(8)
                        End If
                        If String.IsNullOrEmpty(currentRow(34)) = False Then
                            ObjLineItem.ShippingName = currentRow(34)
                        End If
                        If String.IsNullOrEmpty(currentRow(36)) = False Then
                            ObjLineItem.ShippingAddress1 = currentRow(36)
                        End If
                        If String.IsNullOrEmpty(currentRow(37)) = False Then
                            ObjLineItem.ShippingAddress2 = currentRow(37)
                        End If
                        If String.IsNullOrEmpty(currentRow(39)) = False Then
                            ObjLineItem.ShippingCity = currentRow(39)
                        End If

                        If String.IsNullOrEmpty(currentRow(41)) = False Then
                            ObjLineItem.ShippingProvince = currentRow(41)
                        End If
                        If String.IsNullOrEmpty(currentRow(40)) = False Then
                            ObjLineItem.ShippingZip = currentRow(40)
                        End If

                        If String.IsNullOrEmpty(currentRow(43)) = False Then
                            ObjLineItem.ShippingPhone = currentRow(43)
                        End If

                        If String.IsNullOrEmpty(currentRow(9)) = False Then
                            ObjLineItem.Shippingmethod = currentRow(9)
                        End If

                        'TODO: Ensure all fields are added
                    End If

                    LstShopifyItems.Add(ObjLineItem)
                    'Read until a change in the Name field - its the 1st and contains a number
                    'Email 
                    'Paid at  - Paid Date
                    'Subtotal - amount paid
                    'Shipping method - Email: Delivery with 24 hours, US Mail
                    'Lineitem quantity
                    'Lineitem name  - Photography Services, Tandem Skydive - 60 sec Freefall, Your Skydive Photos!, Tandem Skydive - 30 sec Freefall
                    'Lineitem price
                    'Billing Name
                    'Billing Street
                    'Billing Address1
                    'Billing Address2
                    'Billing Company
                    'Billing City
                    'Billing Zip
                    'Billing Province
                    'Billing Country
                    'Billing Phone
                    'Shipping Name
                    'Shipping Street
                    'Shipping Address1
                    'Shipping Address2
                    'Shipping Company
                    'Shipping City
                    'Shipping Zip
                    'Shipping Province
                    'Shipping Country
                    'Shipping Phone
                    'Note Attributes - "recipients-name: Cyrelle Uesonoda
                    '                   special-occasion: Happy 45th Birthday!!"
                    'Payment Method - Authorize.net
                    'Payment Reference
                    'Vendor
                    'Id - looks like a unique code

                    irow += 1
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                    "is not valid and will be skipped.")
                End Try
            End While


            '//Process the records and write to gift certificate table in 
            'So there could be multiple line items and process each.  ie. one for tandem. one for video.
            'These need to be combined so that they can be written as one record.
            Dim LstGiftCertificatesToAddToQueue As New List(Of ClsGiftCertificate)


            'Process Line Items into certificates
            Dim LastId As String = ""
            Dim Index = 0
            Dim ObjGiftCertificate As New ClsGiftCertificate
            For Each lineitem In LstShopifyItems
                If lineitem.Name <> LastId AndAlso Index <> 0 Then
                    'Write Certificate to database

                    If String.IsNullOrEmpty(ObjGiftCertificate.Recipient_FirstName) Then
                        ObjGiftCertificate.Recipient_FirstName = ObjGiftCertificate.Purchaser_FirstName
                        ObjGiftCertificate.Recipient_LastName = ObjGiftCertificate.Purchaser_LastName
                    End If

                    LstGiftCertificatesToAddToQueue.Add(ObjGiftCertificate)
                    ObjGiftCertificate = New ClsGiftCertificate

                End If
                If String.IsNullOrEmpty(lineitem.Name) = False Then
                    ObjGiftCertificate.ShopifyName = lineitem.Name
                End If

                If String.IsNullOrEmpty(lineitem.BillingName) = False Then
                    'Split the name into first and last
                    Dim parts = lineitem.BillingName.Trim.Split(" "c)
                    Dim fn As String = ""
                    Dim ln As String = ""

                    If parts.Count = 2 Then
                        ObjGiftCertificate.Purchaser_FirstName = parts(0)
                        ObjGiftCertificate.Purchaser_LastName = parts(1)
                    Else
                        ObjGiftCertificate.Purchaser_FirstName = parts(0)
                        Dim result As String = String.Join(" ", parts)
                        Dim removefirstname = result.Replace(parts(0), "")
                        ObjGiftCertificate.Purchaser_LastName = removefirstname
                    End If
                End If
                If String.IsNullOrEmpty(lineitem.BillingStreet) = False Then
                    ObjGiftCertificate.Purchaser_Address1 = lineitem.BillingStreet
                End If
                If String.IsNullOrEmpty(lineitem.BillingAddress2) = False Then
                    ObjGiftCertificate.Purchaser_Address2 = lineitem.BillingAddress2
                End If
                If String.IsNullOrEmpty(lineitem.BillingCity) = False Then
                    ObjGiftCertificate.Purchaser_City = lineitem.BillingCity
                End If

                If String.IsNullOrEmpty(lineitem.BillingProvince) = False Then
                    ObjGiftCertificate.Purchaser_State = lineitem.BillingProvince
                End If

                If String.IsNullOrEmpty(lineitem.BillingZip) = False Then
                    ObjGiftCertificate.Purchaser_Zip = lineitem.BillingZip
                End If

                If String.IsNullOrEmpty(lineitem.BillingPhone) = False Then
                    ObjGiftCertificate.Purchaser_Phones1 = lineitem.BillingPhone
                End If

                If String.IsNullOrEmpty(lineitem.BillingEmail) = False Then
                    ObjGiftCertificate.Purchaser_Email = lineitem.BillingEmail
                End If

                If String.IsNullOrEmpty(lineitem.Shippingmethod) = False Then
                    If lineitem.Shippingmethod = "Email: Delivery with 24 hours" Then
                        ObjGiftCertificate.delivery = DeliveryOptions.Email
                    ElseIf lineitem.Shippingmethod = "US Mail" Then
                        ObjGiftCertificate.delivery = DeliveryOptions.USMail
                    ElseIf lineitem.Shippingmethod.StartsWith("Will Call/Pick-up") Then
                        ObjGiftCertificate.delivery = DeliveryOptions.InOffice
                    ElseIf String.IsNullOrEmpty(lineitem.Shippingmethod.trim) Then
                        ObjGiftCertificate.delivery = DeliveryOptions.Email
                    End If
                End If

                If String.IsNullOrEmpty(lineitem.PaidAt) = False Then
                    ObjGiftCertificate.GC_DateEntered = DateTime.Parse(lineitem.PaidAt)
                End If

                ObjGiftCertificate.PaymentMethod = PaymentMethod.Online

                Dim Notes As String = ""
                If String.IsNullOrEmpty(lineitem.Notes) = False Then
                    Notes = Notes & lineitem.Notes
                End If
                If String.IsNullOrEmpty(lineitem.NotesAttributes) = False Then
                    Notes = Notes & lineitem.NotesAttributes
                End If
                If String.IsNullOrEmpty(Notes) = False Then
                    ObjGiftCertificate.Notes = Notes
                End If

                ObjGiftCertificate.PointOfSale = PointOfSale.Online
                'Lineitem name  - Photography Services, Tandem Skydive - 60 sec Freefall, Your Skydive Photos!, Tandem Skydive - 30 sec Freefall

                '//Calculate Amount
                If String.IsNullOrEmpty(lineitem.LineitemName) = False Then
                    If lineitem.LineitemName = "Tandem Skydive - 60 sec Freefall" Then
                        ObjGiftCertificate.Item_Tandem12k = lineitem.LineitemQuantity
                        ObjGiftCertificate.Item_Tandem12kAmount = lineitem.LineitemQuantity * lineitem.LineitemPrice
                    ElseIf lineitem.LineitemName = "Tandem Skydive - 30 sec Freefall" Then
                        ObjGiftCertificate.Item_Tandem10k = lineitem.LineitemQuantity
                        ObjGiftCertificate.Item_Tandem10kAmount = lineitem.LineitemQuantity * lineitem.LineitemPrice
                    ElseIf lineitem.LineitemName = "Your Skydive Photos!" Then
                        ObjGiftCertificate.Item_Video = lineitem.LineitemQuantity
                        ObjGiftCertificate.Item_VideoAmount = lineitem.LineitemQuantity * lineitem.LineitemPrice
                    ElseIf lineitem.LineitemName = "Photography Services" Then
                        ObjGiftCertificate.Item_Video = lineitem.LineitemQuantity
                        ObjGiftCertificate.Item_VideoAmount = lineitem.LineitemQuantity * lineitem.LineitemPrice

                    End If
                End If
                Try
                    ObjGiftCertificate.GC_Number = lineitem.Name

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                '//Recipient Fields
                Dim RecipientName As String = ""
                If String.IsNullOrEmpty(lineitem.NotesAttributes) = False Then
                    'Split the name into first and last
                    Dim parts = lineitem.NotesAttributes.Trim.Split(vbLf)
                    For Each i1 As String In parts
                        If i1.Contains("recipients-name:") Then
                            RecipientName = i1.Replace("recipients-name:", "").Replace(vbLf, "").Trim
                            Exit For
                        End If
                    Next

                    If String.IsNullOrEmpty(RecipientName) = False Then
                        If String.IsNullOrEmpty(RecipientName) = False Then
                            'Split the name into first and last
                            Dim parts2 = RecipientName.Trim.Split(" "c)
                            Dim fn As String = ""
                            Dim ln As String = ""

                            If parts2.Count = 2 Then
                                ObjGiftCertificate.Recipient_FirstName = parts2(0)
                                ObjGiftCertificate.Recipient_LastName = parts2(1)
                            Else
                                ObjGiftCertificate.Recipient_FirstName = parts2(0)
                                Dim result As String = String.Join(" ", parts2)
                                Dim removefirstname1 = result.Replace(parts2(0), "")
                                ObjGiftCertificate.Recipient_LastName = removefirstname1
                            End If
                        End If

                    End If
                End If

                'Shipping method - Email: Delivery with 24 hours, US Mail
                'iF Cointent is not blank then add item to objGiftCertificate

                LastId = lineitem.Name
                Index = Index + 1
            Next

            If Index > 0 Then
                'ProcessLastItem
                'Write Certificate to database
                LstGiftCertificatesToAddToQueue.Add(ObjGiftCertificate)
            End If


            '//Actually Write the Certificates to the Database if needed
            Dim icounter As Integer = 0
            Dim duplicateCertificatesFound As Integer = 0
            Dim newCertificatesCreated As Integer = 0

            For Each gc In LstGiftCertificatesToAddToQueue
                'Insert New Gift Certificate3
                If icounter > 0 Then
                    If DoesGCExist(gc.GC_Number) = False Then
                        InsertNewGiftCertRecord(gc)
                        newCertificatesCreated += 1
                    Else
                        duplicateCertificatesFound += 1
                    End If
                Else

                End If
                icounter += 1
            Next

            '//Display a final message detailing records imported
            Dim sd = String.Format("{0} records are found. {1} will inserted into the queue, {2} are duplicates", LstGiftCertificatesToAddToQueue.Count, newCertificatesCreated, duplicateCertificatesFound)
            MsgBox(sd)

        End Using
    End Sub
End Module
