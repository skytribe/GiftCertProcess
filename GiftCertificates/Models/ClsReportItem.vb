Public Class ClsReportItem
    Sub New()
        Item1 = New ClsItem(1)
        Item2 = New ClsItem(2)
        Item3 = New ClsItem(3)
        Item4 = New ClsItem(4)
        Item5 = New ClsItem(5)
    End Sub
    Public Property EnteredDate As DateTime
    Public Property GCOrderId As String

    Public Property Name As String
    Public Property BillingAddress As String
    Public Property ShippingAddress As String

    Public Property DeliveryMethod As String
    Public Property EmailAddress As String
    Public Property Phone1 As String
    Public Property Authorized As String

    Public Property Status As String

    Public Property Item1 As ClsItem
    Public Property Item2 As ClsItem
    Public Property Item3 As ClsItem

    Public Property Item4 As ClsItem
    Public Property Item5 As ClsItem

    Public Property TotalAmount As Double
    Public Property TotalDiscount As Double
    Public Property DiscountCode As String
    Public Property JRCustomerID As String

    Public Property OnlineOrderID As String
    Public Property ProcessedDate As String

    'Date   GCNumber,   Purchaser Name,   Purchase Address,    Item1 QTY, Item 2 Qty,  Item 3 Qty,   Total Amount,   Status, Delivery method, Email

End Class
