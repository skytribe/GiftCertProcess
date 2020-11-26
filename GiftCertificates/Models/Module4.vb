Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Module Module4

    Public sOpInsertUser As String = "GCProcess"

    Public _strConn As String = "Data Source=DESKTOP-LBU2SR9;Initial Catalog=JumpRunTraining;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"
    Private _sqlCon As SqlConnection
    Public Property ErrorString1 As String = ""
    Public BlnDevMode = False

    Public Const DEVMODE As Boolean = True

    Public Dct_Pricing As New Dictionary(Of Integer, ClsPricing)
    Public Dct_Discount As New Dictionary(Of Integer, ClsPricing)

    Public Sub PopulatePricing()
        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            'Dim sqlstring = String.Format("Select * from dbo.GiftCertificatePricing WHERE IsItem=1")
            cmd.CommandText = "dbo.GCO_RetrievePricing"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("IsItem", 1)
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Dct_Pricing.Clear()

            Do While reader.Read
                Dim x As New ClsPricing
                x.ID = reader("Id").ToString()
                x.SKU = reader("SKU").ToString()
                x.Description = reader("Description").ToString()
                x.Price = reader("Price").ToString()
                x.JR_ItemID = reader("JumpRunItemID").ToString()
                x.Discountable = reader("DiscountableItem").ToString()
                Dct_Pricing.Add(x.ID, x)
            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try


    End Sub
    Public Function RetrieveAuthorizers() As List(Of KeyValuePair(Of Integer, String))


        Dim LstPricing As New List(Of KeyValuePair(Of Integer, String))
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "dbo.GCO_RetrieveAuthorizers"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)


            Do While reader.Read
                Dim x As New KeyValuePair(Of Integer, String)
                Dim ikey = CInt(reader("Id").ToString())
                Dim sValue = String.Format("{0}   ({1})", reader("Name").ToString(), reader("Code").ToString())
                x = New KeyValuePair(Of Integer, String)(ikey, sValue)
                LstPricing.Add(x)
            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try

        Return LstPricing
    End Function

    Public Function RetrieveAuthorizerCode(Id As Integer) As String

        Dim sValue As String = ""

        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "Select * from dbo.GCO_Authorizer where ID=" & Id.ToString
            cmd.CommandType = CommandType.Text

            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)


            Do While reader.Read

                sValue = reader("Code").ToString()
                Exit Do

            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try

        Return sValue
    End Function
    Public Function RetrieveDiscounts() As List(Of ClsPricing)


        Dim LstPricing As New List(Of ClsPricing)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "dbo.GCO_RetrievePricing"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("IsItem", 0)
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Dct_Discount.Clear()

            Do While reader.Read
                Dim x As New ClsPricing
                x.ID = reader("Id").ToString()
                x.SKU = reader("SKU").ToString()
                x.Description = reader("Description").ToString()
                x.Price = reader("Price").ToString()
                x.JR_ItemID = reader("JumpRunItemID").ToString()
                x.Discountable = reader("DiscountableItem").ToString()
                LstPricing.Add(x)
            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try

        Return LstPricing
    End Function


    Public Function RetrieveOperators() As List(Of String)


        Dim LstPricing As New List(Of String)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "select * from tOperators"
            cmd.CommandType = CommandType.Text

            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Dct_Discount.Clear()

            Do While reader.Read

                Dim ID = reader("sOperID").ToString()
                Dim name = reader("sName").ToString()

                LstPricing.Add(ID)
            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try

        Return LstPricing
    End Function
    Public Sub PopulateDiscounts()
        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "dbo.GCO_RetrievePricing"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("IsItem", 0)
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Dct_Discount.Clear()

            Do While reader.Read
                Dim x As New ClsPricing
                x.ID = reader("Id").ToString()
                x.SKU = reader("SKU").ToString()
                x.Description = reader("Description").ToString()
                x.Price = reader("Price").ToString()
                x.JR_ItemID = reader("JumpRunItemID").ToString()
                x.Discountable = reader("DiscountableItem").ToString()
                Dct_Discount.Add(x.ID, x)
            Loop

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try


    End Sub


    Public Function GetJumpRunItemPrice(id As Integer) As Double
        Dim retValue As Double = 0
        Dim sqlCon1 As SqlConnection = Nothing
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            sqlCon1 = New SqlConnection(_strConn)

            cmd.CommandText = "PriceItemGet"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("pwId", id)
            cmd.Parameters.AddWithValue("dtBusiness", Now)
            cmd.Connection = sqlCon1

            sqlCon1.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)

            Do While reader.Read

                retValue = reader("cPrice").ToString()

            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            sqlCon1.Close()
        End Try
        Return retValue

    End Function
    Function PopulateQuantityItems() As List(Of KeyValuePair(Of Integer, String))
        Dim Dct_Qty As New List(Of KeyValuePair(Of Integer, String))
        Dct_Qty.Clear()
        For i = 0 To 10
            Dct_Qty.Add(New KeyValuePair(Of Integer, String)(i, i.ToString))
        Next
        Return Dct_Qty
    End Function
    'Public Function GetPricingForItemIdPrice(ItemId As Integer) As Double
    '    Try
    '        If Dct_Pricing.Count = 0 Then
    '            PopulatePricing()
    '        End If
    '        If ItemId >= 1 And ItemId <= 5 Then
    '            Dim p = Dct_Pricing(ItemId)
    '            Return p.Price
    '        Else
    '            Return 0
    '        End If
    '    Catch ex As Exception
    '        Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
    '        Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
    '        LogError(methodName, ex)
    '    End Try

    'End Function

    Public Function GetPricingForItemIdPrice(ItemId As Integer) As Double
        Try
            If Dct_Pricing.Count = 0 Then
                PopulatePricing()
            End If
            If ItemId >= 1 And ItemId <= 5 Then
                Dim p = Dct_Pricing(ItemId)
                Return p.Price
            Else
                Return 0
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function

    Public Function GetDescriptionForItemId(ItemId As Integer) As String
        Try
            If Dct_Pricing.Count = 0 Then
                PopulatePricing()
            End If

            If ItemId >= 1 And ItemId <= 5 Then
                Dim p = Dct_Pricing(ItemId)
                Return p.Description
            Else
                Return ""
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function
    Public Function GetJumpRunForItemId(ItemId As Integer) As Integer
        Try
            If Dct_Pricing.Count = 0 Then
                PopulatePricing()
            End If

            If ItemId >= 1 And ItemId <= 5 Then
                Dim p = Dct_Pricing(ItemId)
                Return p.JR_ItemID
            Else
                Return 0
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function

    Public Function GetPricingForItemId(ItemId As Integer) As ClsPricing
        Try
            If Dct_Pricing.Count = 0 Then
                PopulatePricing()
            End If

            If ItemId >= 1 And ItemId <= 5 Then
                Dim p = Dct_Pricing(ItemId)
                Return p
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try


    End Function
    Public Function GetJumpRunItemIdForDiscount(SKUCode As String) As Integer
        Try
            If Dct_Discount.Count = 0 Then
                PopulateDiscounts()
            End If

            For Each i In Dct_Discount
                If i.Value.SKU = SKUCode Then
                    Return i.Value.JR_ItemID
                End If
            Next
            Return 0
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try


    End Function

    Public Function GetPricingForDiscountId(SKUCode As String) As ClsPricing
        Try
            If Dct_Discount.Count = 0 Then
                PopulateDiscounts()
            End If


            For Each i In Dct_Discount
                If i.Value.SKU = SKUCode Then
                    Return i.Value
                End If
            Next
            Return Nothing
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try


    End Function

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

    Public Function GetFileFolderDocumentPath(filename As String) As String
        Try
            Dim appFolder = My.Application.Info.DirectoryPath
            Dim MasterFileFolder = System.IO.Path.Combine(appFolder, "Files")
            Dim CertificateToPrint = System.IO.Path.Combine(MasterFileFolder, filename)
            Dim PrintDocument = System.IO.Path.Combine(appFolder, CertificateToPrint)
            Return PrintDocument
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
        Return ""
    End Function

    Function GetAddressString(i As ClsAddress) As String
        Dim sb As New System.Text.StringBuilder
        Try
            If String.IsNullOrEmpty(i.Address1.Trim) = False Then
                sb.AppendLine(i.Address1.Trim)
            End If
            If String.IsNullOrEmpty(i.Address2.Trim) = False Then
                sb.AppendLine(i.Address2.Trim)
            End If
            If String.IsNullOrEmpty(i.City.Trim) = False Then
                sb.AppendLine(String.Format("{0} {1} {2}", i.City.Trim, i.State.Trim, i.Zip.Trim))
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return sb.ToString
    End Function

    Sub GetConnectionString()
        Try
            _strConn = My.Settings.DBConnectionString
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub





    Public Function RetrieveGCOrdersFromQueue(d As Date, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate2)
        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            If devmode Then
                cmd.CommandText = "dbo.GCO_GetOrders"
                cmd.CommandType = CommandType.StoredProcedure
            Else

                cmd.CommandText = "dbo.GCO_OrdersInDateRange"
                cmd.CommandType = CommandType.StoredProcedure
                Dim FromDate As DateTime = d.Date.ToShortDateString
                Dim ToDate As DateTime = d.Date.AddHours(23).AddMinutes(59).AddSeconds(59)

                cmd.Parameters.AddWithValue("FromDate", String.Format(FromDate, "YYYY-MM-DD HH: MM:SS"))
                cmd.Parameters.AddWithValue("ToDate", String.Format(ToDate, "YYYY-MM-DD HH:MM:SS"))

            End If

            cmd.Connection = _sqlCon

            _sqlCon.Open()
            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjOrderRecord As New ClsGiftCertificate2

                ObjOrderRecord.ID = reader("GCOrderId").ToString()
                ObjOrderRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                ObjOrderRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                ObjOrderRecord.Billing_Address.Address1 = reader("Purchaser_Address1").ToString()
                ObjOrderRecord.Billing_Address.Address2 = reader("Purchaser_Address2").ToString()
                ObjOrderRecord.Billing_Address.City = reader("Purchaser_City").ToString()
                ObjOrderRecord.Billing_Address.State = reader("Purchaser_State").ToString()
                ObjOrderRecord.Billing_Address.Zip = reader("Purchaser_Zip").ToString()
                ObjOrderRecord.Billing_Address.Phone1 = reader("Purchaser_Phone1").ToString()
                ObjOrderRecord.Billing_Address.Phone2 = reader("Purchaser_Phone2").ToString()
                ObjOrderRecord.Billing_Address.Email = reader("Purchaser_Email").ToString()
                ObjOrderRecord.Shipping_Address.Address1 = reader("Shipping_Address1").ToString()
                ObjOrderRecord.Shipping_Address.Address2 = reader("Shipping_Address2").ToString()
                ObjOrderRecord.Shipping_Address.City = reader("Shipping_City").ToString()
                ObjOrderRecord.Shipping_Address.State = reader("Shipping_State").ToString()
                ObjOrderRecord.Shipping_Address.Zip = reader("Shipping_Zip").ToString()
                ObjOrderRecord.Shipping_Address.Email = reader("Shipping_Email").ToString()
                ObjOrderRecord.Item1.Quantity = reader("Item1Qty")
                ObjOrderRecord.Item1.ItemId = Product_ItemType.Item_TDM10K

                ObjOrderRecord.Item2.Quantity = reader("Item2Qty")
                ObjOrderRecord.Item2.ItemId = Product_ItemType.Item_TDM12K

                ObjOrderRecord.Item3.Quantity = reader("Item3Qty")
                ObjOrderRecord.Item3.ItemId = Product_ItemType.Item_TDM10KVID

                ObjOrderRecord.Item4.Quantity = reader("Item4Qty")
                ObjOrderRecord.Item4.ItemId = Product_ItemType.Item_TDM12KVID

                ObjOrderRecord.Item5.Quantity = reader("Item5Qty")
                ObjOrderRecord.Item5.ItemId = Product_ItemType.Item_VID

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    ObjOrderRecord.JR_PurchaserID = reader("JR_PurchaseID")
                End If

                ObjOrderRecord.GC_DateEntered = reader("DateEntered").ToString()
                ObjOrderRecord.HearAbout = reader("HearAbout").ToString()
                ObjOrderRecord.delivery = reader("DeliveryOption").ToString()
                ObjOrderRecord.PointOfSale = reader("PointOfSales").ToString()
                ObjOrderRecord.Notes = reader("Notes").ToString()
                ObjOrderRecord.GC_Authorization = reader("Authorization").ToString()
                ObjOrderRecord.GC_Username = reader("UserName").ToString()
                ObjOrderRecord.GC_TotalAmount = reader("Item_CalculatedTotal").ToString()
                ObjOrderRecord.GC_TotalDiscount = reader("Item_DiscountTotal").ToString()
                ObjOrderRecord.GC_DiscountCode = reader("DiscountCode").ToString

                ObjOrderRecord.GC_Status = reader("Status")
                ObjOrderRecord.Online_OrderNumber = reader("Online_Certificate_Number").ToString
                ObjOrderRecord.PaymentMethod = reader("PaymentMethod").ToString
                LstPossibles.Add(ObjOrderRecord)
            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles
    End Function


    Public Function RetrieveGCOrdersFromQueue(status As FilterStates, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate2)
        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            If devmode Then
                cmd.CommandText = "dbo.GCO_GetOrders"
                cmd.CommandType = CommandType.StoredProcedure
            Else

                cmd.CommandText = "GCO_OrdersFromStatus"
                cmd.CommandType = CommandType.StoredProcedure

                '//Either FilterStates.Entered, FilterStates.Processing or FilterStates.Incomplete
                cmd.Parameters.AddWithValue("Status", status)


            End If

            cmd.Connection = _sqlCon

            _sqlCon.Open()
            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjOrderRecord As New ClsGiftCertificate2

                ObjOrderRecord.ID = reader("GCOrderId").ToString()
                ObjOrderRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                ObjOrderRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                ObjOrderRecord.Billing_Address.Address1 = reader("Purchaser_Address1").ToString()
                ObjOrderRecord.Billing_Address.Address2 = reader("Purchaser_Address2").ToString()
                ObjOrderRecord.Billing_Address.City = reader("Purchaser_City").ToString()
                ObjOrderRecord.Billing_Address.State = reader("Purchaser_State").ToString()
                ObjOrderRecord.Billing_Address.Zip = reader("Purchaser_Zip").ToString()
                ObjOrderRecord.Billing_Address.Phone1 = reader("Purchaser_Phone1").ToString()
                ObjOrderRecord.Billing_Address.Phone2 = reader("Purchaser_Phone2").ToString()
                ObjOrderRecord.Billing_Address.Email = reader("Purchaser_Email").ToString()
                ObjOrderRecord.Shipping_Address.Address1 = reader("Shipping_Address1").ToString()
                ObjOrderRecord.Shipping_Address.Address2 = reader("Shipping_Address2").ToString()
                ObjOrderRecord.Shipping_Address.City = reader("Shipping_City").ToString()
                ObjOrderRecord.Shipping_Address.State = reader("Shipping_State").ToString()
                ObjOrderRecord.Shipping_Address.Zip = reader("Shipping_Zip").ToString()
                ObjOrderRecord.Shipping_Address.Email = reader("Shipping_Email").ToString()
                ObjOrderRecord.Item1.Quantity = reader("Item1Qty")
                ObjOrderRecord.Item1.ItemId = Product_ItemType.Item_TDM10K

                ObjOrderRecord.Item2.Quantity = reader("Item2Qty")
                ObjOrderRecord.Item2.ItemId = Product_ItemType.Item_TDM12K

                ObjOrderRecord.Item3.Quantity = reader("Item3Qty")
                ObjOrderRecord.Item3.ItemId = Product_ItemType.Item_TDM10KVID

                ObjOrderRecord.Item4.Quantity = reader("Item4Qty")
                ObjOrderRecord.Item4.ItemId = Product_ItemType.Item_TDM12KVID

                ObjOrderRecord.Item5.Quantity = reader("Item5Qty")
                ObjOrderRecord.Item5.ItemId = Product_ItemType.Item_VID

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    ObjOrderRecord.JR_PurchaserID = reader("JR_PurchaseID")
                End If

                ObjOrderRecord.GC_DateEntered = reader("DateEntered").ToString()
                ObjOrderRecord.HearAbout = reader("HearAbout").ToString()
                ObjOrderRecord.delivery = reader("DeliveryOption").ToString()
                ObjOrderRecord.PointOfSale = reader("PointOfSales").ToString()
                ObjOrderRecord.Notes = reader("Notes").ToString()
                ObjOrderRecord.GC_Authorization = reader("Authorization").ToString()
                ObjOrderRecord.GC_Username = reader("UserName").ToString()
                ObjOrderRecord.GC_TotalAmount = reader("Item_CalculatedTotal").ToString()
                ObjOrderRecord.GC_TotalDiscount = reader("Item_DiscountTotal").ToString()
                ObjOrderRecord.GC_DiscountCode = reader("DiscountCode").ToString

                ObjOrderRecord.GC_Status = reader("Status")
                ObjOrderRecord.Online_OrderNumber = reader("Online_Certificate_Number").ToString
                ObjOrderRecord.PaymentMethod = reader("PaymentMethod").ToString
                LstPossibles.Add(ObjOrderRecord)
            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles
    End Function


    Public Function RetrieveJumpRunItems(Optional SearchString As String = "") As List(Of ClsJumpRunItem)
        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsJumpRunItem)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)


            cmd.CommandText = "dbo.GCO_GetJumpRunItems"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = _sqlCon

            _sqlCon.Open()
            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjOrderRecord As New ClsJumpRunItem

                If String.IsNullOrEmpty(SearchString) Then
                    ObjOrderRecord.ID = reader("wItemId").ToString()
                    ObjOrderRecord.Description = reader("sItem").ToString()
                    ObjOrderRecord.Price = reader("cPrice").ToString()
                    LstPossibles.Add(ObjOrderRecord)
                Else
                    Dim sItemUpper = reader("sItem").ToString().ToUpper
                    If sItemUpper.Contains(SearchString.ToUpper) Then
                        ObjOrderRecord.ID = reader("wItemId").ToString()
                        ObjOrderRecord.Description = reader("sItem").ToString()
                        ObjOrderRecord.Price = reader("cPrice").ToString()
                        LstPossibles.Add(ObjOrderRecord)
                    End If


                End If


            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles
    End Function

    Public Function InsertPayment(pDate As Date, pCustID As Integer, pAmt As Decimal, sComments As String) As String
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.PmtInsert"
            sqlComm.CommandType = CommandType.StoredProcedure

            Dim DTStamp As Date = Now

            sqlComm.Parameters.AddWithValue("@ReturnVal_1", 0)
            sqlComm.Parameters.AddWithValue("@dtProc_2", pDate)
            sqlComm.Parameters.AddWithValue("@dtBus_3", pDate)
            sqlComm.Parameters.AddWithValue("@wId_4", 0)
            sqlComm.Parameters.AddWithValue("@wCustId_5", pCustID)
            sqlComm.Parameters.AddWithValue("@wBillToId_6", pCustID)
            sqlComm.Parameters.AddWithValue("@nTransType_7", 1)
            sqlComm.Parameters.AddWithValue("@nMethodId_8", 9)
            sqlComm.Parameters.AddWithValue("@sComment_9", sComments)
            sqlComm.Parameters.AddWithValue("@cAmount_10", pAmt * -1)
            sqlComm.Parameters.AddWithValue("@wReemItemID_11", 0)
            sqlComm.Parameters.AddWithValue("@sSerialNo_12", 0)
            sqlComm.Parameters.AddWithValue("@dtInsert_13", DTStamp)
            sqlComm.Parameters.AddWithValue("@sOperInsert_14", sOpInsertUser)
            sqlComm.Parameters.AddWithValue("@dtUpdate_15", DTStamp)
            sqlComm.Parameters.AddWithValue("@sOperUpdate_16", sOpInsertUser)
            sqlComm.Parameters.AddWithValue("@wRelatedTo_17", 0)
            sqlComm.Parameters.AddWithValue("@cNewBal_18", 0)
            sqlComm.Parameters.AddWithValue("@cNewToday_19", 0)
            sqlComm.Parameters.AddWithValue("@sMachine_20", My.Computer.Name)

            sqlComm.ExecuteNonQuery()
            SuccessState = True
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return SuccessState
    End Function

    Public Function RetrieveGCOrdersFromQueue(name As String, Optional devmode As Boolean = False) As List(Of ClsGiftCertificate2)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)
            Dim sqlstring As String = ""
            If devmode = False Then
                sqlstring = String.Format("Select * from dbo.GiftCertificate where Purchaser_FirstName like  '%{0}%' OR Purchaser_LastName like  '%{1}%'", name.Trim, name.Trim)
                cmd.CommandText = sqlstring
            Else
                '//Devmode
                cmd.CommandText = "Select * from dbo.GiftCertificates"
            End If

            cmd.CommandType = CommandType.Text
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim PossibleRecord As New ClsGiftCertificate2

                PossibleRecord.ID = reader("GCOrderId").ToString()

                PossibleRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                PossibleRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                PossibleRecord.Billing_Address.Address1 = reader("Purchaser_Address1").ToString()
                PossibleRecord.Billing_Address.Address2 = reader("Purchaser_Address2").ToString()
                PossibleRecord.Billing_Address.City = reader("Purchaser_City").ToString()
                PossibleRecord.Billing_Address.State = reader("Purchaser_State").ToString()
                PossibleRecord.Billing_Address.Zip = reader("Purchaser_Zip").ToString()
                PossibleRecord.Billing_Address.Phone1 = reader("Purchaser_Phone1").ToString()
                PossibleRecord.Billing_Address.Phone2 = reader("Purchaser_Phone2").ToString()
                PossibleRecord.Billing_Address.Email = reader("Purchaser_Email").ToString()

                'PossibleRecord.Recipient_FirstName = reader("Recipient_FirstName").ToString()
                'PossibleRecord.Recipient_LastName = reader("Recipient_LastName").ToString()
                PossibleRecord.Shipping_Address.Address1 = reader("Shipping_Address1").ToString()
                PossibleRecord.Shipping_Address.Address2 = reader("Shipping_Address2").ToString()
                PossibleRecord.Shipping_Address.City = reader("Shipping_City").ToString()
                PossibleRecord.Shipping_Address.State = reader("Shipping_State").ToString()
                PossibleRecord.Shipping_Address.Zip = reader("Shipping_Zip").ToString()
                PossibleRecord.Shipping_Address.Email = reader("Shipping_Email").ToString()
                PossibleRecord.Item1.Quantity = reader("Item1Qty")
                PossibleRecord.Item1.ItemId = 1

                PossibleRecord.Item2.Quantity = reader("Item2Qty")
                PossibleRecord.Item2.ItemId = 2

                PossibleRecord.Item3.Quantity = reader("Item3Qty")
                PossibleRecord.Item3.ItemId = 3

                PossibleRecord.Item4.Quantity = reader("Item4Qty")
                PossibleRecord.Item4.ItemId = 4

                PossibleRecord.Item5.Quantity = reader("Item5Qty")
                PossibleRecord.Item5.ItemId = 5

                If IsDBNull(reader("JR_PurchaseID")) = False Then
                    PossibleRecord.JR_PurchaserID = reader("JR_PurchaseID")
                End If

                PossibleRecord.GC_DateEntered = reader("DateEntered").ToString()
                PossibleRecord.HearAbout = reader("HearAbout").ToString()
                PossibleRecord.delivery = reader("DeliveryOption").ToString()
                PossibleRecord.PointOfSale = reader("PointOfSales").ToString()
                PossibleRecord.Notes = reader("Notes").ToString()
                PossibleRecord.GC_Authorization = reader("Authorization").ToString()
                PossibleRecord.GC_Username = reader("UserName").ToString()
                PossibleRecord.GC_TotalAmount = reader("Item_CalculatedTotal").ToString()
                PossibleRecord.GC_TotalDiscount = reader("Item_DiscountTotal").ToString()
                PossibleRecord.GC_Status = reader("Status")
                PossibleRecord.Online_OrderNumber = reader("Online_Certificate_Number").ToString
                PossibleRecord.PaymentMethod = reader("PaymentMethod").ToString

                ' PossibleRecord.GC_TotalDiscount = reader("DiscountAmount").ToString
                PossibleRecord.GC_DiscountCode = reader("DiscountCode").ToString

                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles

    End Function

    Public Function RetrieveGCOrdersLineItems(gc As ClsGiftCertificate2) As List(Of ClsPrintCertificateDetails)

        Dim sfield As String = ""

        Dim LstPrintCerts As New List(Of ClsPrintCertificateDetails)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            cmd.CommandText = "dbo.GCO_RetrieveLineItemsForOrder"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = _sqlCon
            cmd.Parameters.AddWithValue("Id", gc.ID)
            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjPrintCertificateRecord As New ClsPrintCertificateDetails
                ObjPrintCertificateRecord.ID = reader("Id").ToString()
                Dim itemId = reader("ItemId").ToString()
                ObjPrintCertificateRecord.Description = GetPricingForItemId(itemId).Description
                ObjPrintCertificateRecord.JumpRunCertificateNumber = reader("JumpRunCertificateNumber").ToString()
                ObjPrintCertificateRecord.OrderDate = gc.GC_DateEntered
                ObjPrintCertificateRecord.PersoanlizedFrom = gc.Purchaser_Name
                ObjPrintCertificateRecord.Authorizer = gc.GC_Authorization
                LstPrintCerts.Add(ObjPrintCertificateRecord)
            Loop
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPrintCerts
    End Function





    Public Function GCOrdersReport(d1 As Date, d2 As Date) As List(Of ClsGiftCertificate2)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsGiftCertificate2)
        GetConnectionString()
        Try
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            Dim FromDate = d1.Date.ToShortDateString
            Dim ToDate = d2.Date.AddHours(23).AddMinutes(59).AddSeconds(59)

            cmd.CommandText = "Dbo.GCO_OrdersInDateRange"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("FromDate", FromDate)
            cmd.Parameters.AddWithValue("ToDate", ToDate)
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjOrderRecord As New ClsGiftCertificate2
                Try

                    ' ObjOrderRecord.OrderId = reader("GCOrderId").ToString()
                    ObjOrderRecord.OrderId = reader("GCOrderId").ToString()
                    ObjOrderRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                    ObjOrderRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                    ObjOrderRecord.Billing_Address.Address1 = reader("Purchaser_Address1").ToString()
                    ObjOrderRecord.Billing_Address.Address2 = reader("Purchaser_Address2").ToString()
                    ObjOrderRecord.Billing_Address.City = reader("Purchaser_City").ToString()
                    ObjOrderRecord.Billing_Address.State = reader("Purchaser_State").ToString()
                    ObjOrderRecord.Billing_Address.Zip = reader("Purchaser_Zip").ToString()
                    ObjOrderRecord.Billing_Address.Phone1 = reader("Purchaser_Phone1").ToString()
                    ObjOrderRecord.Billing_Address.Phone2 = reader("Purchaser_Phone2").ToString()
                    ObjOrderRecord.Billing_Address.Email = reader("Purchaser_Email").ToString()
                    ObjOrderRecord.Shipping_Address.Address1 = reader("Shipping_Address1").ToString()
                    ObjOrderRecord.Shipping_Address.Address2 = reader("Shipping_Address2").ToString()
                    ObjOrderRecord.Shipping_Address.City = reader("Shipping_City").ToString()
                    ObjOrderRecord.Shipping_Address.State = reader("Shipping_State").ToString()
                    ObjOrderRecord.Shipping_Address.Zip = reader("Shipping_Zip").ToString()
                    ObjOrderRecord.Shipping_Address.Email = reader("Shipping_Email").ToString()
                    ObjOrderRecord.Item1.Quantity = reader("Item1Qty")
                    ObjOrderRecord.Item1.ItemId = Product_ItemType.Item_TDM10K

                    ObjOrderRecord.Item2.Quantity = reader("Item2Qty")
                    ObjOrderRecord.Item2.ItemId = Product_ItemType.Item_TDM12K

                    ObjOrderRecord.Item3.Quantity = reader("Item3Qty")
                    ObjOrderRecord.Item3.ItemId = Product_ItemType.Item_TDM10KVID

                    ObjOrderRecord.Item4.Quantity = reader("Item4Qty")
                    ObjOrderRecord.Item4.ItemId = Product_ItemType.Item_TDM12KVID

                    ObjOrderRecord.Item5.Quantity = reader("Item5Qty")
                    ObjOrderRecord.Item5.ItemId = Product_ItemType.Item_VID

                    If IsDBNull(reader("JR_PurchaseID")) = False Then
                        ObjOrderRecord.JR_PurchaserID = reader("JR_PurchaseID")
                    End If

                    ObjOrderRecord.GC_DateEntered = reader("DateEntered").ToString()
                    ObjOrderRecord.GC_ProcessedDate = reader("ProcessDate").ToString()
                    ObjOrderRecord.HearAbout = reader("HearAbout").ToString()
                    ObjOrderRecord.delivery = reader("DeliveryOption").ToString()
                    ObjOrderRecord.PointOfSale = reader("PointOfSales").ToString()
                    ObjOrderRecord.Notes = reader("Notes").ToString()
                    ObjOrderRecord.GC_Authorization = reader("Authorization").ToString()
                    ObjOrderRecord.GC_Username = reader("UserName").ToString()
                    ObjOrderRecord.GC_TotalAmount = reader("Item_CalculatedTotal").ToString()
                    ObjOrderRecord.GC_TotalDiscount = reader("Item_DiscountTotal").ToString()
                    ObjOrderRecord.GC_DiscountCode = reader("DiscountCode").ToString

                    ObjOrderRecord.GC_Status = reader("Status")
                    ObjOrderRecord.Online_OrderNumber = reader("Online_Certificate_Number").ToString
                    ObjOrderRecord.PaymentMethod = reader("PaymentMethod").ToString


                Catch ex As Exception
                    '//TODO: Verify conditions this could occur
                End Try

                LstPossibles.Add(ObjOrderRecord)
            Loop
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles

    End Function
    Public Function GCOrders_Search(lastname As String) As List(Of ClsGiftCertificate2)

        Dim LstPossibles As New List(Of ClsGiftCertificate2)

        Try
            GetConnectionString()

            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader

            _sqlCon = New SqlConnection(_strConn)

            Dim SQLString = "dbo.GCO_FindCertificatesWithLastName"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = SQLString
            cmd.Parameters.AddWithValue("lastname", lastname)
            cmd.Connection = _sqlCon

            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read
                Dim ObjOrderRecord As New ClsGiftCertificate2
                Try
                    ObjOrderRecord.ID = reader("GCOrderId").ToString()
                    ObjOrderRecord.Purchaser_FirstName = reader("Purchaser_FirstName").ToString()
                    ObjOrderRecord.Purchaser_LastName = reader("Purchaser_LastName").ToString()
                    ObjOrderRecord.Billing_Address.Address1 = reader("Purchaser_Address1").ToString()
                    ObjOrderRecord.Billing_Address.Address2 = reader("Purchaser_Address2").ToString()
                    ObjOrderRecord.Billing_Address.City = reader("Purchaser_City").ToString()
                    ObjOrderRecord.Billing_Address.State = reader("Purchaser_State").ToString()
                    ObjOrderRecord.Billing_Address.Zip = reader("Purchaser_Zip").ToString()
                    ObjOrderRecord.Billing_Address.Phone1 = reader("Purchaser_Phone1").ToString()
                    ObjOrderRecord.Billing_Address.Phone2 = reader("Purchaser_Phone2").ToString()
                    ObjOrderRecord.Billing_Address.Email = reader("Purchaser_Email").ToString()
                    ObjOrderRecord.Shipping_Address.Address1 = reader("Shipping_Address1").ToString()
                    ObjOrderRecord.Shipping_Address.Address2 = reader("Shipping_Address2").ToString()
                    ObjOrderRecord.Shipping_Address.City = reader("Shipping_City").ToString()
                    ObjOrderRecord.Shipping_Address.State = reader("Shipping_State").ToString()
                    ObjOrderRecord.Shipping_Address.Zip = reader("Shipping_Zip").ToString()
                    ObjOrderRecord.Shipping_Address.Email = reader("Shipping_Email").ToString()
                    ObjOrderRecord.Item1.Quantity = reader("Item1Qty")
                    ObjOrderRecord.Item1.ItemId = Product_ItemType.Item_TDM10K

                    ObjOrderRecord.Item2.Quantity = reader("Item2Qty")
                    ObjOrderRecord.Item2.ItemId = Product_ItemType.Item_TDM12K

                    ObjOrderRecord.Item3.Quantity = reader("Item3Qty")
                    ObjOrderRecord.Item3.ItemId = Product_ItemType.Item_TDM10KVID

                    ObjOrderRecord.Item4.Quantity = reader("Item4Qty")
                    ObjOrderRecord.Item4.ItemId = Product_ItemType.Item_TDM12KVID

                    ObjOrderRecord.Item5.Quantity = reader("Item5Qty")
                    ObjOrderRecord.Item5.ItemId = Product_ItemType.Item_VID

                    If IsDBNull(reader("JR_PurchaseID")) = False Then
                        ObjOrderRecord.JR_PurchaserID = reader("JR_PurchaseID")
                    End If

                    ObjOrderRecord.GC_DateEntered = reader("DateEntered").ToString()
                    ObjOrderRecord.HearAbout = reader("HearAbout").ToString()
                    ObjOrderRecord.delivery = reader("DeliveryOption").ToString()
                    ObjOrderRecord.PointOfSale = reader("PointOfSales").ToString()
                    ObjOrderRecord.Notes = reader("Notes").ToString()
                    ObjOrderRecord.GC_Authorization = reader("Authorization").ToString()
                    ObjOrderRecord.GC_Username = reader("UserName").ToString()
                    ObjOrderRecord.GC_TotalAmount = reader("Item_CalculatedTotal").ToString()
                    ObjOrderRecord.GC_TotalDiscount = reader("Item_DiscountTotal").ToString()
                    ObjOrderRecord.GC_DiscountCode = reader("DiscountCode").ToString

                    ObjOrderRecord.GC_Status = reader("Status")
                    ObjOrderRecord.Online_OrderNumber = reader("Online_Certificate_Number").ToString
                    ObjOrderRecord.PaymentMethod = reader("PaymentMethod").ToString
                Catch ex As Exception
                    '//TODO: Verify conditions this could occur
                End Try

                LstPossibles.Add(ObjOrderRecord)
            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Return LstPossibles
    End Function


    Public Function LoadMatchData(c As ClsPersonSearch) As List(Of ClsJumpRunPossibleCustomers)

        Dim sfield As String = ""

        Dim LstPossibles As New List(Of ClsJumpRunPossibleCustomers)
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
            cmd.Parameters("@PhoneNumber").Value = c.Phone.Trim & ""
            cmd.Parameters.Add("@PreferedName", SqlDbType.NVarChar)
            cmd.Parameters("@PreferedName").Value = String.Format("{0},{1}", c.LastName, c.FirstName)
            _sqlCon.Open()

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While reader.Read

                Dim PossibleRecord As New ClsJumpRunPossibleCustomers
                sfield = "WCustid"


                PossibleRecord.wCustId = reader(0)
                sfield = "sCust"
                PossibleRecord.sCust = getString(reader(1))
                sfield = "sFirstName"
                PossibleRecord.sFirstName = getString(reader(2)).Trim

                If c.FirstName.ToLower.Trim = PossibleRecord.sFirstName.ToLower.Trim And String.IsNullOrEmpty(c.FirstName.Trim) = False Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 15
                End If
                sfield = "sLastName"
                PossibleRecord.sLastName = getString(reader(3)).Trim
                If c.LastName.ToLower.Trim = PossibleRecord.sLastName.ToLower.Trim And String.IsNullOrEmpty(c.LastName.Trim) = False Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 20
                End If

                sfield = "sEmail"
                PossibleRecord.sEmail = getString(reader(6)).Trim
                If String.IsNullOrEmpty(c.Email) = False AndAlso c.Email.ToLower.Trim = PossibleRecord.sEmail.ToLower.Trim Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 20
                End If

                sfield = "sEmail2"
                PossibleRecord.sEmail2 = getString(reader(7)).Trim
                If String.IsNullOrEmpty(c.Email) = False AndAlso c.Email.ToLower.Trim = PossibleRecord.sEmail2.ToLower.Trim Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                sfield = "sPhone"
                PossibleRecord.sPhone1 = getString(reader(8)).Trim
                If String.IsNullOrEmpty(c.Phone) = False AndAlso c.Phone.Trim = PossibleRecord.sPhone1.Trim Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 20
                ElseIf String.IsNullOrEmpty(c.Phone) = False AndAlso "1" & c.Phone.Trim = PossibleRecord.sPhone1.Trim Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 20
                ElseIf String.IsNullOrEmpty(c.Phone) = False AndAlso c.Phone.Trim = "1" & PossibleRecord.sPhone1.Trim Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 20

                End If

                sfield = "sZip"
                PossibleRecord.sZip = getString(reader(15)).Trim
                If String.IsNullOrEmpty(c.Zip) = False AndAlso PossibleRecord.sZip.StartsWith(c.Zip.Trim) Then
                    PossibleRecord.PercentageMatch = PossibleRecord.PercentageMatch + 10
                End If

                LstPossibles.Add(PossibleRecord)
            Loop


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An Error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon.Close()
        End Try

        Dim matchingRecords = From i In LstPossibles
                              Where i.PercentageMatch >= 20
                              Order By i.PercentageMatch Descending
                              Select i

        Return matchingRecords.ToList
    End Function

    Function GetBusinessDate() As DateTime
        Dim DteBusDate As DateTime
        Dim sqlCon As SqlConnection
        Try
            GetConnectionString()

            sqlCon = New SqlConnection(_strConn)
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.CommandText = "Select sValue From tConfig Where SKey = 'BusDate'"
            sqlComm.CommandType = CommandType.Text
            sqlComm.Connection = sqlCon

            Dim Reader As SqlDataReader = sqlComm.ExecuteReader()

            ' Data is accessible through the DataReader object here.
            ' Use Read method (true/false) to see if reader has records and advance to next record
            ' You can use a While loop for multiple records (While reader.Read() ... End While)
            Do While Reader.Read


                DteBusDate = DateTime.Parse(Reader("sValue").ToString())
                Exit Do
            Loop
            Return DteBusDate

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlCon.Close()
        End Try

    End Function

    Public Function InsertJumpRunInvRecord(c As ClsCertificateItems, InsertBy As String, dtInsert As DateTime) As Boolean
        Dim SuccessState As Boolean = False

        Try
            GetConnectionString()

            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.InvInsert"

            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.Add("ReturnVal_1", SqlDbType.Int)
            sqlComm.Parameters("ReturnVal_1").Direction = ParameterDirection.Output

            sqlComm.Parameters.AddWithValue("dtProc_2", dtInsert)
            sqlComm.Parameters.AddWithValue("dtBus_3", dtInsert)

            sqlComm.Parameters.AddWithValue("siSerialized_4", 2)  '//Its a serialized Item
            sqlComm.Parameters.AddWithValue("bInventory_5", 0)
            sqlComm.Parameters.Add("wId_6", SqlDbType.Int)
            sqlComm.Parameters("wId_6").Direction = ParameterDirection.Output

            sqlComm.Parameters.AddWithValue("nManiId_7", 0) '????
            sqlComm.Parameters.AddWithValue("wCustId_8", c.JumpRunCustomerID)
            sqlComm.Parameters.AddWithValue("wBillToId_9", c.JumpRunCustomerID)

            Dim JRItemID = GetPricingForItemId(c.ItemId)

            sqlComm.Parameters.AddWithValue("wItemId_10", c.JumpRunItemId)

            sqlComm.Parameters.AddWithValue("sSerialNo_11", c.JRCertificateNumber)
            sqlComm.Parameters.AddWithValue("hQty_12", 1)
            sqlComm.Parameters.AddWithValue("nBodyCnt_13", SqlDbType.SmallInt)
            sqlComm.Parameters("nBodyCnt_13").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("cTax1_14", 0)
            sqlComm.Parameters.AddWithValue("cTax2_15", 0)
            sqlComm.Parameters.AddWithValue("cTax3_16", 0)
            sqlComm.Parameters.AddWithValue("cPrice_17", c.Amount)
            sqlComm.Parameters.AddWithValue("bManualPrice_18", 0)
            sqlComm.Parameters.AddWithValue("cTimeAdj_19", 0)
            sqlComm.Parameters.AddWithValue("cWeekdayAdj_20", 0)
            sqlComm.Parameters.AddWithValue("cGroupAdj_21", 0)
            sqlComm.Parameters.AddWithValue("cDZAdj_22", 0)
            sqlComm.Parameters.AddWithValue("cCashAdj_23", 0)
            sqlComm.Parameters.AddWithValue("cTeamAdj_24", 0)
            sqlComm.Parameters.AddWithValue("cCategAdj_25", 0)
            sqlComm.Parameters.AddWithValue("cPersAdj_26", 0)
            sqlComm.Parameters.AddWithValue("hWeight_27", -1)
            sqlComm.Parameters.AddWithValue("nSeat_28", 0)
            'sqlComm.Parameters.AddWithValue("hWeight_27", SqlDbType.Decimal)
            'sqlComm.Parameters("hWeight_27").Direction = ParameterDirection.Output
            'sqlComm.Parameters.AddWithValue("nSeat_28", SqlDbType.SmallInt)
            'sqlComm.Parameters("nSeat_28").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("sComment_29", "GC " & c.JRCertificateNumber)
            sqlComm.Parameters.AddWithValue("nTeamNo_30", 0)
            sqlComm.Parameters.AddWithValue("dtInsert_31", Now)
            sqlComm.Parameters.AddWithValue("sOperInsert_32", sOpInsertUser)
            sqlComm.Parameters.AddWithValue("dtUpdate_33", DBNull.Value)
            sqlComm.Parameters.AddWithValue("sOperUpdate_34", "")
            sqlComm.Parameters.AddWithValue("wRelatedTo_35", 0)
            sqlComm.Parameters.AddWithValue("nItemType_36", 0)
            sqlComm.Parameters.AddWithValue("bCalcCG_37", 0)
            sqlComm.Parameters.AddWithValue("nNewRiders_38", SqlDbType.SmallInt)
            sqlComm.Parameters("nNewRiders_38").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("hNewWt_39", SqlDbType.Decimal)
            sqlComm.Parameters("hNewWt_39").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("hNewCG_40", SqlDbType.Decimal)
            sqlComm.Parameters("hNewCG_40").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("cNewBal_41", SqlDbType.Money)
            sqlComm.Parameters("cNewBal_41").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("sMachine_42", My.Computer.Name)

            sqlComm.ExecuteNonQuery()
            SuccessState = True

            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)

            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            _sqlCon = Nothing
        End Try
        Return SuccessState
    End Function


    Public Function InsertJumpRunInvDiscountRecord(c As ClsCertificateItems, InsertBy As String, dtInsert As DateTime) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.InvInsert"

            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.Add("ReturnVal_1", SqlDbType.Int)
            sqlComm.Parameters("ReturnVal_1").Direction = ParameterDirection.Output

            sqlComm.Parameters.AddWithValue("dtProc_2", dtInsert)
            sqlComm.Parameters.AddWithValue("dtBus_3", dtInsert) '???

            sqlComm.Parameters.AddWithValue("siSerialized_4", 0)  '//Its a serialized Item
            sqlComm.Parameters.AddWithValue("bInventory_5", 0)
            sqlComm.Parameters.Add("wId_6", SqlDbType.Int)
            sqlComm.Parameters("wId_6").Direction = ParameterDirection.Output

            sqlComm.Parameters.AddWithValue("nManiId_7", 0) '????
            sqlComm.Parameters.AddWithValue("wCustId_8", c.JumpRunCustomerID)
            sqlComm.Parameters.AddWithValue("wBillToId_9", c.JumpRunCustomerID)

            Dim JRItemID = GetPricingForItemId(c.ItemId)

            sqlComm.Parameters.AddWithValue("wItemId_10", c.JumpRunItemId)

            sqlComm.Parameters.AddWithValue("sSerialNo_11", "DISC FOR- " & c.JRCertificateNumber)
            sqlComm.Parameters.AddWithValue("hQty_12", 1)
            sqlComm.Parameters.AddWithValue("nBodyCnt_13", SqlDbType.SmallInt)
            sqlComm.Parameters("nBodyCnt_13").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("cTax1_14", 0)
            sqlComm.Parameters.AddWithValue("cTax2_15", 0)
            sqlComm.Parameters.AddWithValue("cTax3_16", 0)
            sqlComm.Parameters.AddWithValue("cPrice_17", c.Amount)
            sqlComm.Parameters.AddWithValue("bManualPrice_18", 0)
            sqlComm.Parameters.AddWithValue("cTimeAdj_19", 0)
            sqlComm.Parameters.AddWithValue("cWeekdayAdj_20", 0)
            sqlComm.Parameters.AddWithValue("cGroupAdj_21", 0)
            sqlComm.Parameters.AddWithValue("cDZAdj_22", 0)
            sqlComm.Parameters.AddWithValue("cCashAdj_23", 0)
            sqlComm.Parameters.AddWithValue("cTeamAdj_24", 0)
            sqlComm.Parameters.AddWithValue("cCategAdj_25", 0)
            sqlComm.Parameters.AddWithValue("cPersAdj_26", 0)
            sqlComm.Parameters.AddWithValue("hWeight_27", -1)
            sqlComm.Parameters.AddWithValue("nSeat_28", 0)
            'sqlComm.Parameters.AddWithValue("hWeight_27", SqlDbType.Decimal)
            'sqlComm.Parameters("hWeight_27").Direction = ParameterDirection.Output
            'sqlComm.Parameters.AddWithValue("nSeat_28", SqlDbType.SmallInt)
            'sqlComm.Parameters("nSeat_28").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("sComment_29", "Discount for GC " & c.JRCertificateNumber & Environment.NewLine & "Order Reference: " & c.CertificateOrderReference)
            sqlComm.Parameters.AddWithValue("nTeamNo_30", 0)
            sqlComm.Parameters.AddWithValue("dtInsert_31", Now)
            sqlComm.Parameters.AddWithValue("sOperInsert_32", InsertBy)
            sqlComm.Parameters.AddWithValue("dtUpdate_33", DBNull.Value)
            sqlComm.Parameters.AddWithValue("sOperUpdate_34", "")
            sqlComm.Parameters.AddWithValue("wRelatedTo_35", 0)
            sqlComm.Parameters.AddWithValue("nItemType_36", 0)
            sqlComm.Parameters.AddWithValue("bCalcCG_37", 0)
            sqlComm.Parameters.AddWithValue("nNewRiders_38", SqlDbType.SmallInt)
            sqlComm.Parameters("nNewRiders_38").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("hNewWt_39", SqlDbType.Decimal)
            sqlComm.Parameters("hNewWt_39").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("hNewCG_40", SqlDbType.Decimal)
            sqlComm.Parameters("hNewCG_40").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("cNewBal_41", SqlDbType.Money)
            sqlComm.Parameters("cNewBal_41").Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("sMachine_42", My.Computer.Name)


            '(@ReturnVal_1   [int] OUTPUT,
            ' @dtProc_2	[datetime],             CALL: Date To Process (Now)
            ' @dtBus_3	[datetime],             CALL: ?????  I belive this Is the current business date - this along with above determine which tables are used for inventory / manifest updates
            ' @siSerialized_4 [smallint],            CALL:= 0 
            ' @bInventory_5  [bit],                  CALL:= 0 (Its Not an inventory controlled item)
            ' @wId_6 	[int] OUTPUT,           It returns next ID for inventory 
            ' @nManiId_7 	[smallint],
            ' @wCustId_8 	[int],                  CALL: CustomerID to use 
            ' @wBillToId_9 	[int],
            ' @wItemId_10 	[int],                  CALL: The Item to use (5 for different products)
            ' @sSerialNo_11 	[nvarchar](30),         CALL: With New serialnumber To use
            ' @hQty_12 	[decimal] (9, 2),       CALL WITH QTY =1  
            ' @nBodyCnt_13 	[smallint] OUTPUT,      CALL with 0
            ' @cTax1_14 	[money],                CALL with 0
            ' @cTax2_15 	[money],                CALL with 0
            ' @cTax3_16 	[money],                CALL with 0
            ' @cPrice_17 	[money],                 CALL: This will be the price of the item (I believe this should match with Item price)
            ' @bManualPrice_18 	[bit],           CALL WITH False
            ' @cTimeAdj_19 	[money],                 0
            ' @cWeekdayAdj_20 	[money],         0
            ' @cGroupAdj_21 	[money],                 0
            ' @cDZAdj_22 	[money],                 0 
            ' @cCashAdj_23 	[money],                 0
            ' @cTeamAdj_24 	[money],                 0
            ' @cCategAdj_25 	[money],                 0              
            ' @cPersAdj_26 	[money],                 0
            ' @hWeight_27 	[decimal] (9, 2) OUTPUT, 0
            ' @nSeat_28 	[smallint] OUTPUT ,      -1
            ' @sComment_29 	[nvarchar](30),          CALL WITH "GC xxxxx"
            ' @nTeamNo_30 	[smallint],              0
            ' @dtInsert_31 	[datetime],              CALL  with NOw
            ' @sOperInsert_32 	[nvarchar](10),  CALL: This will be the name of the person inserting
            ' @dtUpdate_33 	[datetime],              Call with Null
            ' @sOperUpdate_34 	[nvarchar](10),  Call with Null
            ' @wRelatedTo_35 	[int],           0
            ' @nItemType_36	[smallint],              ????? 
            ' @bCalcCG_37	[bit],                   0
            ' @nNewRiders_38	[smallint] OUTPUT,       ????? Output
            ' @hNewWt_39	[decimal] (9, 2)	OUTPUT,    ?????? Output
            ' @hNewCG_40	[decimal] (9, 2) OUTPUT,   ???? Output
            ' @cNewBal_41	[money] OUTPUT,            ???? Output
            ' @sMachine_42	[nvarchar](31))             CALL: Machine Name being used
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
    Public Sub GenerateIndividualCertificateRecordsFromGCOrder(GCOrder As ClsGiftCertificate2, Optional DEVTEST As Boolean = False)
        'Determine Qty and Items
        Try
            Dim dtInsert As DateTime = GetBusinessDate()

            Dim CertificateOrderId As Integer = GCOrder.ID
            PopulatePricing()

            '//Count Tandem Itesm - ie.  Exclude Video Only Item
            'Exclude Item3 Qty
            Dim IndividualDiscountAmount As Integer = CalculateItemDiscountAmountFromOrder(GCOrder)

            'Determine The Gift Certificates I need to produced from the quantities of each product that have been orders
            Dim Items As New List(Of ClsCertificateItems)
            'AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetPricingForItemIdPrice(1), GCOrder.Item1, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            'AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetPricingForItemIdPrice(2), GCOrder.Item2, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            'AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetPricingForItemIdPrice(3), GCOrder.Item3, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            'AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetPricingForItemIdPrice(4), GCOrder.Item4, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            'AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetPricingForItemIdPrice(5), GCOrder.Item5, Items, IndividualDiscountAmount, "")

            Dim s As String = "TDM10K"
            Dim o = GetPricingForItemId(1).JR_ItemID

            AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetJumpRunItemPrice(CInt(GetPricingForItemId(1).JR_ItemID)), GCOrder.Item1, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetJumpRunItemPrice(CInt(GetPricingForItemId(2).JR_ItemID)), GCOrder.Item2, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetJumpRunItemPrice(CInt(GetPricingForItemId(3).JR_ItemID)), GCOrder.Item3, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetJumpRunItemPrice(CInt(GetPricingForItemId(4).JR_ItemID)), GCOrder.Item4, Items, IndividualDiscountAmount, GCOrder.GC_DiscountCode)
            AddItemToVirtualCart(GCOrder.ID, GCOrder.JR_PurchaserID, GetJumpRunItemPrice(CInt(GetPricingForItemId(5).JR_ItemID)), GCOrder.Item5, Items, IndividualDiscountAmount, "")

            If Items.Count > 0 Then
                Dim iindex As Integer = 1
                For Each i In Items
                    '//Create LIne Items for the Order
                    'Get Next Girft Certificate ID
                    Dim strGCReference = String.Format("{0}-{1}", i.GCOrderId, iindex)
                    Dim NewId = InsertGCOrdersReference(strGCReference)
                    i.JRCertificateNumber = NewId.ToString
                    i.CertificateOrderReference = strGCReference

                    '//1. Write a record in GiftCertificateOrders_Config which will assign a new Certificate Number which we will use in 
                    '//JumpRun rather than have it assign the numbers - this way we can ensure continuity with existing system,
                    '//2. Update GiftCertificateItems record with new assigned number
                    i.JumpRunItemId = GetJumpRunForItemId(i.ItemId)
                    'Verify If discountCode is set
                    'If it is then we cam specify a Discount Amount
                    If String.IsNullOrEmpty(i.DiscountCode) = False Then
                        i.DiscountCode = i.DiscountCode
                        i.DiscountAmount = IndividualDiscountAmount
                    End If

                    '//This will write a GCOrder Line Item record - which has GCReference Number for each certificate being generated
                    InsertGiftCertificateLineItem(i)

                    '//This generates the Jumprun Inventory and Redeem tables records
                    '//For Gift Certificate and Discount
                    InsertJumpRunInvRecord(i, sOpInsertUser, dtInsert)  '?/Insert Item Record

                    '//Generate a discount record if needed. 
                    If String.IsNullOrEmpty(i.DiscountCode) = False Then
                        Dim ObjDiscount = GetPricingForDiscountId(i.DiscountCode.ToUpper)
                        If ObjDiscount IsNot Nothing Then
                            Dim iDisc As New ClsCertificateItems

                            iDisc.GCOrderId = i.GCOrderId
                            iDisc.JumpRunCustomerID = i.JumpRunCustomerID
                            iDisc.JumpRunItemId = ObjDiscount.JR_ItemID
                            iDisc.Amount = ObjDiscount.Price
                            iDisc.JRCertificateNumber = i.JRCertificateNumber
                            InsertJumpRunInvDiscountRecord(iDisc, sOpInsertUser, dtInsert)
                        Else
                            MsgBox("Discount Not Found")
                        End If
                    End If
                    iindex += 1
                Next
                'Insert a Payment Record For Customer
                Dim RetValue = InsertPayment(dtInsert, GCOrder.JR_PurchaserID, GCOrder.GC_TotalAmount, "Payment For GC Order" & GCOrder.Online_OrderNumber)
            Else
                MessageBox.Show("LineItems Required to be created " & Items.Count)
            End If




        Catch ex As Exception

        End Try


    End Sub

    Private Function CalculateItemDiscountAmountFromOrder(GCOrder As ClsGiftCertificate2) As Integer
        Dim IndividualDiscountAmount = 0
        Try
            Dim TdmQty = GCOrder.Item1.Quantity + GCOrder.Item2.Quantity + GCOrder.Item3.Quantity + GCOrder.Item4.Quantity
            If TdmQty > 0 Then
                IndividualDiscountAmount = GCOrder.GC_TotalDiscount / TdmQty
            End If
        Catch ex As Exception
            MsgBox("Calculating discount for order")
        End Try

        Return IndividualDiscountAmount
    End Function

    Private Sub AddItemToVirtualCart(CertificateId As Integer, JumpRunCustomerID As Integer, ItemAmount As Double, Item As ClsItem, ByRef Items As List(Of ClsCertificateItems), DIscountAmount As Double, DiscountCode As String)
        Try
            If Item.Quantity > 0 Then
                For i = 1 To Item.Quantity
                    Dim objLineItem As New ClsCertificateItems
                    objLineItem.GCOrderId = CertificateId
                    objLineItem.JumpRunCustomerID = JumpRunCustomerID
                    objLineItem.ItemId = Item.ItemId
                    objLineItem.Amount = ItemAmount
                    '//If not discountable then these should be 0
                    Dim objpricing = GetPricingForItemId(Item.ItemId)
                    If objpricing.Discountable Then
                        objLineItem.DiscountAmount = DIscountAmount
                        objLineItem.DiscountCode = DiscountCode
                    Else
                        objLineItem.DiscountAmount = 0
                        objLineItem.DiscountCode = ""
                    End If


                    Items.Add(objLineItem)
                Next
            End If
        Catch ex As Exception
            MsgBox("Problem adding items to virtual cart")
        End Try

    End Sub

    Public Function InsertGCOrdersReference(GCOrderReference As String) As Integer
        Dim SuccessState As Boolean = False
        Dim NewId As Integer = 0
        GetConnectionString()
        Try

            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.GCO_InsertGCOrdersReference"

            sqlComm.CommandType = CommandType.StoredProcedure


            sqlComm.Parameters.Add("Id", SqlDbType.Int).Direction = ParameterDirection.Output
            sqlComm.Parameters.AddWithValue("GCOrderReference", GCOrderReference)
            sqlComm.ExecuteNonQuery()
            NewId = Convert.ToInt32(sqlComm.Parameters("Id").Value)
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
        Return NewId
    End Function



    Public Function CreateGiftCertificatesFromCertificateRecord(c As ClsGiftCertificate2) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.GCO_InsertGiftCertificate"

            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("DateEntered", Now)
            sqlComm.Parameters.AddWithValue("HearAbout", c.HearAbout)
            sqlComm.Parameters.AddWithValue("DeliveryOption", c.delivery)
            sqlComm.Parameters.AddWithValue("PointOfSales", c.PointOfSale)
            sqlComm.Parameters.AddWithValue("Authorization", c.GC_Authorization)
            sqlComm.Parameters.AddWithValue("UserName", c.GC_Username)
            sqlComm.Parameters.AddWithValue("Status", c.GC_Status)
            sqlComm.Parameters.AddWithValue("Purchaser_FirstName", c.Purchaser_FirstName)
            sqlComm.Parameters.AddWithValue("Purchaser_LastName", c.Purchaser_LastName)
            sqlComm.Parameters.AddWithValue("BillingAddress_Address1", c.Billing_Address.Address1)
            sqlComm.Parameters.AddWithValue("BillingAddress_Address2", c.Billing_Address.Address2)
            sqlComm.Parameters.AddWithValue("BillingAddress_City", c.Billing_Address.City)
            sqlComm.Parameters.AddWithValue("BillingAddress_State", c.Billing_Address.State)
            sqlComm.Parameters.AddWithValue("BillingAddress_Zip", c.Billing_Address.Zip)
            sqlComm.Parameters.AddWithValue("BillingAddress_Phone1", c.Billing_Address.Phone1)
            sqlComm.Parameters.AddWithValue("BillingAddress_Phone2", c.Billing_Address.Phone2)
            sqlComm.Parameters.AddWithValue("BillingAddress_Email", c.Billing_Address.Email)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Address1", c.Shipping_Address.Address1)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Address2", c.Shipping_Address.Address2)
            sqlComm.Parameters.AddWithValue("ShippingAddress_City", c.Shipping_Address.City)
            sqlComm.Parameters.AddWithValue("ShippingAddress_State", c.Shipping_Address.State)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Zip", c.Shipping_Address.Zip)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Email", c.Shipping_Address.Email)
            sqlComm.Parameters.AddWithValue("Item1Qty", c.Item1.Quantity)
            sqlComm.Parameters.AddWithValue("Item2Qty", c.Item2.Quantity)
            sqlComm.Parameters.AddWithValue("Item3Qty", c.Item3.Quantity)
            sqlComm.Parameters.AddWithValue("Item4Qty", c.Item4.Quantity)
            sqlComm.Parameters.AddWithValue("Item5Qty", c.Item5.Quantity)
            sqlComm.Parameters.AddWithValue("Item_CalculatedTotal", c.GC_TotalAmount)
            sqlComm.Parameters.AddWithValue("Item_TotalDiscount", c.GC_TotalDiscount)
            sqlComm.Parameters.AddWithValue("JR_PurchaserID", c.JR_PurchaserID)
            sqlComm.Parameters.AddWithValue("GC_Number", c.Online_OrderNumber)
            sqlComm.Parameters.AddWithValue("PaymentMethod", c.PaymentMethod)
            If c.GC_DiscountCode Is Nothing Then
                sqlComm.Parameters.AddWithValue("DiscountCode", "")
            Else
                sqlComm.Parameters.AddWithValue("DiscountCode", c.GC_DiscountCode)
            End If


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

    Public Function UpdateGiftCertificatesFromCertificateRecord(c As ClsGiftCertificate2) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            '        LogError(String.Format(RecordDetail, c.FirstName, c.LastName, c.wCustId, c.SmartWaiverID), "")


            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.GCO_UpdateGiftCertificate"

            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("Id", c.ID)
            sqlComm.Parameters.AddWithValue("HearAbout", c.HearAbout)
            sqlComm.Parameters.AddWithValue("DeliveryOption", c.delivery)
            sqlComm.Parameters.AddWithValue("PointOfSales", c.PointOfSale)
            sqlComm.Parameters.AddWithValue("Purchaser_FirstName", c.Purchaser_FirstName)
            sqlComm.Parameters.AddWithValue("Purchaser_LastName", c.Purchaser_LastName)
            sqlComm.Parameters.AddWithValue("BillingAddress_Address1", c.Billing_Address.Address1)
            sqlComm.Parameters.AddWithValue("BillingAddress_Address2", c.Billing_Address.Address2)
            sqlComm.Parameters.AddWithValue("BillingAddress_City", c.Billing_Address.City)
            sqlComm.Parameters.AddWithValue("BillingAddress_State", c.Billing_Address.State)
            sqlComm.Parameters.AddWithValue("BillingAddress_Zip", c.Billing_Address.Zip)
            sqlComm.Parameters.AddWithValue("BillingAddress_Phone1", c.Billing_Address.Phone1)
            sqlComm.Parameters.AddWithValue("BillingAddress_Phone2", c.Billing_Address.Phone2)
            sqlComm.Parameters.AddWithValue("BillingAddress_Email", c.Billing_Address.Email)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Address1", c.Shipping_Address.Address1)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Address2", c.Shipping_Address.Address2)
            sqlComm.Parameters.AddWithValue("ShippingAddress_City", c.Shipping_Address.City)
            sqlComm.Parameters.AddWithValue("ShippingAddress_State", c.Shipping_Address.State)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Zip", c.Shipping_Address.Zip)
            sqlComm.Parameters.AddWithValue("ShippingAddress_Email", c.Shipping_Address.Email)
            sqlComm.Parameters.AddWithValue("Item1Qty", c.Item1.Quantity)
            sqlComm.Parameters.AddWithValue("Item2Qty", c.Item2.Quantity)
            sqlComm.Parameters.AddWithValue("Item3Qty", c.Item3.Quantity)
            sqlComm.Parameters.AddWithValue("Item4Qty", c.Item4.Quantity)
            sqlComm.Parameters.AddWithValue("Item5Qty", c.Item5.Quantity)
            sqlComm.Parameters.AddWithValue("Item_CalculatedTotal", c.GC_TotalAmount)
            sqlComm.Parameters.AddWithValue("Item_TotalDiscount", c.GC_TotalDiscount)
            sqlComm.Parameters.AddWithValue("JR_PurchaserID", c.JR_PurchaserID)
            sqlComm.Parameters.AddWithValue("GC_Number", c.Online_OrderNumber)
            sqlComm.Parameters.AddWithValue("PaymentMethod", c.PaymentMethod)
            If c.GC_DiscountCode Is Nothing Then
                sqlComm.Parameters.AddWithValue("DiscountCode", "")
            Else
                sqlComm.Parameters.AddWithValue("DiscountCode", c.GC_DiscountCode)
            End If


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
                Dim SimilarFirstNames As New List(Of ClsPreferredNameFirstNamePair)
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

                            SimilarFirstNames.Add(New ClsPreferredNameFirstNamePair(CInt(sqlReader.GetInt32(0)), sqlReader.GetString(1), sqlReader.GetString(2)))
                        End While
                    End If

                    If sqlCon1.State = ConnectionState.Open Then sqlCon1.Close()

                    '//New
                    '//Determine the non matching firstname with number and adjust the preferredName

                    For i = 1 To 100
                        If SimilarFirstNames.Count > 0 Then
                            c.PreferredName = String.Format("{0},{1}{2}", c.LastName.Trim, actualfirstname.Trim, i.ToString)
                            If DoesCustomerIndexExist(c.PreferredName.Trim) = False Then
                                Exit For
                            End If
                        Else
                            If DoesCustomerIndexExist(c.PreferredName.Trim) = False Then
                                Exit For
                            End If
                        End If
                    Next
                Else

                End If

            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return c
    End Function


    Public Function InsertNewCustomerRecord(c As ClsJumpRunCustomer, ByRef newID As Integer) As Boolean
        Dim SuccessState As Boolean = False

        Try
            GetConnectionString()

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
            sqlComm.CommandText = "dbo.GCO_InsertNewUserCustomer"
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
            sqlComm.Parameters.AddWithValue("sOpInsert", sOpInsertUser)
            sqlComm.Parameters.AddWithValue("dtInsert", c.dtInsert)
            sqlComm.Parameters.Add("new_identity", SqlDbType.Int).Direction = ParameterDirection.Output

            sqlComm.ExecuteNonQuery()
            Dim id = Convert.ToInt32(sqlComm.Parameters("new_identity").Value)
            newID = id
            SuccessState = True

            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon = Nothing
        End Try
        Return SuccessState
    End Function

    Public Function InsertGiftCertificateLineItem(c As ClsCertificateItems) As Boolean
        Dim SuccessState As Boolean = False
        GetConnectionString()
        Try
            _sqlCon = New SqlConnection(_strConn)
            If _sqlCon.State = ConnectionState.Closed Then _sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = _sqlCon

            sqlComm.CommandText = "dbo.GCO_InsertGiftCertificateLineItem"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("CertificateId", c.GCOrderId)
            sqlComm.Parameters.AddWithValue("JumpRunCertificateNumber", c.JRCertificateNumber) '//TODO Get a real next GiftCertificate Id
            sqlComm.Parameters.AddWithValue("GCOrderLineReference", c.CertificateOrderReference) '//TODO Get a real next GiftCertificate Id
            sqlComm.Parameters.AddWithValue("JumpRunCustomerId", c.JumpRunCustomerID)
            sqlComm.Parameters.AddWithValue("ItemID", c.ItemId)
            sqlComm.Parameters.AddWithValue("Amount", c.Amount)
            sqlComm.Parameters.AddWithValue("DiscountAmount", c.DiscountAmount)
            sqlComm.Parameters.AddWithValue("DiscountCode", c.DiscountCode)
            sqlComm.ExecuteNonQuery()

            SuccessState = True

            If _sqlCon.State = ConnectionState.Open Then _sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon = Nothing
        End Try
        Return SuccessState
    End Function



    Function DevelopmentOnly_GetListOfGCCustomers() As IEnumerable(Of Integer)
        Dim LstCustomers As New List(Of Integer)

        GetConnectionString()
        Try
            Dim sqlCon = New SqlConnection(_strConn)
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm1 As New SqlCommand()
            sqlComm1.Connection = sqlCon
            sqlComm1.CommandText = "dbo.GCO_GetCustomers"
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
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _sqlCon = Nothing
        End Try
        Return LstCustomers.AsEnumerable
    End Function



    Sub DevelopmentOnly_DeleteCustomer(Id As String)
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
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)


            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("RemoveFromJumpRun:", ex)
        Finally

            sqlCon = Nothing
        End Try
    End Sub


    '//Verifies if the People table contains this string which is a unique Key
    Public Function DoesCustomerIndexExist(customerKeyString As String) As Boolean
        Dim sqlCommand As New SqlCommand
        Dim sqlconnect As SqlConnection = Nothing

        Try
            GetConnectionString()

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
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            sqlconnect.Close()

        End Try
        Return False
    End Function

    Public Sub DeleteImportedGiftCertificates()
        Dim sqlCon As SqlConnection
        Try
            GetConnectionString()
            sqlCon = New SqlConnection(_strConn)

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_DeleteImportGiftCertificates"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            sqlCon = Nothing
        End Try

    End Sub
    Public Sub ResetImportedGiftCertificatesStatus()
        Dim sqlCon As SqlConnection
        Try
            GetConnectionString()
            sqlCon = New SqlConnection(_strConn)

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_ResetImportGiftCertificates"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            sqlCon = Nothing
        End Try

    End Sub
    Public Sub DeleteAllGCOData()
        Dim sqlCon As SqlConnection
        Try
            GetConnectionString()
            sqlCon = New SqlConnection(_strConn)

            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon

            sqlComm.CommandText = "dbo.GCO_ResetAllData"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.ExecuteNonQuery()

            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            sqlCon = Nothing
        End Try

    End Sub

    Public Function GetAllGCOData() As DataSet
        GetConnectionString()
        Dim sqlCommand As New SqlCommand
        Dim sqlconnect As SqlConnection = Nothing

        Try
            sqlconnect = New SqlConnection(_strConn)
            sqlconnect.Open()

            sqlCommand.CommandText = "dbo.GCO_GetAllGCOData"
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlconnect

            Dim ds = New DataSet()
            Dim adap = New SqlDataAdapter(sqlCommand)
            adap.Fill(ds, "Certs")

            If ds.Tables(0).Rows.Count > 0 Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'LogError("DoesIndexExist:", ex)
        Finally
            sqlconnect.Close()

        End Try


    End Function

    Public Function DoesGCOrderExist(GCNumber As String) As Boolean
        GetConnectionString()
        Dim sqlCommand As New SqlCommand
        Dim sqlconnect As SqlConnection = Nothing

        Try
            sqlconnect = New SqlConnection(_strConn)
            sqlconnect.Open()

            sqlCommand.CommandText = "dbo.GCO_DoesItAlreadyExist"
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlconnect
            sqlCommand.Parameters.AddWithValue("@GC_Number", GCNumber)
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



    Public Function ExecuteSQL(sText As String) As Boolean
        Dim SuccessState As Boolean = False
        Dim Con1 As SqlConnection

        Try
            GetConnectionString()

            Con1 = New SqlConnection(_strConn)
            If Con1.State = ConnectionState.Closed Then Con1.Open()



            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = Con1

            '    
            sqlComm.CommandText = sText
            sqlComm.CommandType = CommandType.Text



            sqlComm.ExecuteNonQuery()


            If Con1.State = ConnectionState.Open Then Con1.Close()
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Con1 = Nothing
        End Try
        Return SuccessState
    End Function
End Module
