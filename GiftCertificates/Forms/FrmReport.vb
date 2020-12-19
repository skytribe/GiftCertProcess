Imports System.Reflection

Public Class FrmReport
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BtnGo.Click
        Try
            Dim LstRecords = GCOrdersReport(SfDateTimeEdit1.Value, SfDateTimeEdit2.Value)
            Dim LstReport As New List(Of ClsReportItem)
            For Each itemRecord In LstRecords
                Dim ci As New ClsReportItem
                ci.GCOrderId = itemRecord.OrderId
                Dim blankdate As DateTime
                If itemRecord.GC_ProcessedDate = blankdate Then
                Else
                    ci.ProcessedDate = itemRecord.GC_ProcessedDate.ToShortDateString
                End If

                ci.EnteredDate = itemRecord.GC_DateEntered.ToShortDateString
                ci.Name = itemRecord.Purchaser_Name
                ci.EmailAddress = itemRecord.Billing_Address.Email
                ci.Phone1 = itemRecord.Billing_Address.Phone1
                ci.BillingAddress = GetAddressString(itemRecord.Billing_Address)
                ci.ShippingAddress = GetAddressString(itemRecord.Shipping_Address)
                ci.Authorized = itemRecord.GC_Authorization

                '//Items
                ci.Item1.ItemId = 1
                ci.Item2.ItemId = 2
                ci.Item3.ItemId = 3
                ci.Item4.ItemId = 4
                ci.Item5.ItemId = 5

                ci.TotalAmount = itemRecord.GC_TotalAmount
                ci.TotalDiscount = itemRecord.GC_TotalDiscount
                ci.DiscountCode = itemRecord.GC_DiscountCode

                ci.DeliveryMethod = GetDeliveryOptionString(itemRecord.delivery)
                ci.Status = GetStatusString(itemRecord.GC_Status)

                If itemRecord.JR_PurchaserID > 0 Then
                    ci.JRCustomerID = itemRecord.JR_PurchaserID.ToString
                End If
                ci.OnlineOrderID = itemRecord.Online_OrderNumber
                LstReport.Add(ci)
            Next

            SfDataGrid1.DataSource = LstReport
            SfDataGrid1.AllowDraggingColumns = True
            For Each c In SfDataGrid1.Columns
                c.AllowDragging = True
            Next

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnExport.Click
        Try
            Dim x = GCOrdersReport(SfDateTimeEdit1.Value, SfDateTimeEdit2.Value)
            Dim LstReport As New List(Of ClsReportItem)
            Dim sb As New System.Text.StringBuilder
            sb.Append("DateEntered,")
            sb.Append("Date Processed,")
            sb.Append("Purchaser Name,")
            sb.Append("Billing Email,")
            sb.Append("Billing Phone,")
            sb.Append("Billing Address,")
            sb.Append("Shipping Address,")
            sb.Append("GC Authorizer,")
            sb.Append("10k Tdm qty,")
            sb.Append("12k Tdm qty,")
            sb.Append("10k Tdm With Vid Qty,")
            sb.Append("12k Tdm With Vid Qty,")
            sb.Append("Vid Qty,")
            sb.Append("Amount,")
            sb.Append("Discount Amount ,")
            sb.Append("Discount Code,")
            sb.Append("Delivery Option,")
            sb.Append("Status,")
            sb.Append("JR Customer ID")
            sb.AppendLine()
            For Each i In x
                Dim ci As New ClsReportItem
                sb.Append(Chr(34) & i.GC_DateEntered & Chr(34) & ",")
                sb.Append(Chr(34) & i.GC_ProcessedDate & Chr(34) & ",")
                sb.Append(Chr(34) & i.Purchaser_Name & Chr(34) & ",")
                sb.Append(Chr(34) & i.Billing_Address.Email & Chr(34) & ",")
                sb.Append(Chr(34) & i.Billing_Address.Phone1 & Chr(34) & ",")
                sb.Append(Chr(34) & GetAddressString(i.Billing_Address) & Chr(34) & ",")
                sb.Append(Chr(34) & GetAddressString(i.Shipping_Address) & Chr(34) & ",")
                sb.Append(Chr(34) & i.GC_Authorization & Chr(34) & ",")
                sb.Append(i.Item1.Quantity & ",")
                sb.Append(i.Item2.Quantity & ",")
                sb.Append(i.Item3.Quantity & ",")
                sb.Append(i.Item4.Quantity & ",")
                sb.Append(i.Item5.Quantity & ",")
                sb.Append(i.GC_TotalAmount & ",")
                sb.Append(i.GC_TotalDiscount & ",")
                sb.Append(i.GC_DiscountCode & ",")

                sb.Append(Chr(34) & GetDeliveryOptionString(i.delivery) & Chr(34) & ",")
                sb.Append(Chr(34) & GetStatusString(i.GC_Status) & Chr(34) & ",")
                If i.JR_PurchaserID > 0 Then
                    sb.Append(i.JR_PurchaserID.ToString)

                Else
                    sb.Append("")
                End If
                sb.AppendLine()

            Next
            Dim savefiledialog1 As New SaveFileDialog
            savefiledialog1.Filter = "CSV Files (*.csv*)|*.csv"
            savefiledialog1.AddExtension = True

            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK _
       Then
                My.Computer.FileSystem.WriteAllText(savefiledialog1.FileName, sb.ToString, False)
            End If


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FrmReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initializegrid()
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        Using file = System.IO.File.Create("FrmReport.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using

    End Sub

    Public Sub initializegrid()
        If IO.File.Exists("FrmReport.xml") Then
            Try
                Using file = System.IO.File.Open("FrmReport.xml", System.IO.FileMode.Open)
                    Me.SfDataGrid1.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If

    End Sub
End Class