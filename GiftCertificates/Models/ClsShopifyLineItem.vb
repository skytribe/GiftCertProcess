Public Class ClsShopifyLineItem
    Public Property Name As String
    Public Property ItemName As String
    'Read until a change in the Name field - its the 1st and contains a number
    Public Property Email As String
    Public Property PaidAt As DateTime '- Paid Date
    Public Property Subtotal As Double '- amount paid
    Public Property Shippingmethod As String '- Email: Delivery with 24 hours, US Mail
    Public Property LineitemQuantity As Integer
    Public Property LineitemName As String '  - Photography Services, Tandem Skydive - 60 sec Freefall, Your Skydive Photos!, Tandem Skydive - 30 sec Freefall
    Public Property LineitemPrice As Double
    Public Property BillingName As String
    Public Property BillingStreet As String
    Public Property BillingAddress1 As String
    Public Property BillingAddress2 As String
    Public Property BillingCompany As String
    Public Property BillingCity As String
    Public Property BillingZip As String
    Public Property BillingProvince As String
    Public Property BillingCountry As String
    Public Property BillingPhone As String
    Public Property BillingEmail As String
    Public Property ShippingName As String
    Public Property ShippingStreet As String
    Public Property ShippingAddress1 As String
    Public Property ShippingAddress2 As String
    Public Property ShippingCompany As String
    Public Property ShippingCity As String
    Public Property ShippingZip As String
    Public Property ShippingProvince As String
    Public Property ShippingCountry As String
    Public Property ShippingPhone As String

    Public Property PaymentMethod As String '- Authorize.net
    Public Property PaymentReference As String

    Public Property Notes As String
    Public Property NotesAttributes As String
    'Note Attributes - "recipients-name: Cyrelle Uesonoda
    '                   special-occasion: Happy 45th Birthday!!"

    'Vendor
    'Id - looks like a unique code

    Public Property RecipientName As String = ""
    'TODO - So is recipient the shipper ????
End Class
