Public Class ClsWooLineItem

    Public Property OrderID As String
    Public Property OrderDate As DateTime
    Public Property WebStoreOrderStatus As String = ""   '//More used for rejected items
    Public Property PaymentMethod As String '- Authorize.net

    Public Property Captured_AuthNet As Boolean = True
    Public Property Shippingmethod As String '- Email: Delivery with 24 hours, US Mail    (Woo - Will Call/ Pickup: Daily 9am-5pm. (Free)
    'US Mail - Discreet labeling (Free)
    'US Mail - Discreet labeling (Free)
    'US Mail(Free)
    'US Mail(Free)
    'Email: Delivery with 24 hours. (Free)
    Public Property ShippingFirstName As String
    Public Property ShippingLastName As String

    Public Property ShippingAddress1 As String
    Public Property ShippingAddress2 As String

    Public Property ShippingCity As String

    Public Property ShippingState As String
    Public Property ShippingZip As String
    Public Property ShippingCountry As String
    Public Property BillingFirstName As String
    Public Property BillingLastName As String

    Public Property BillingAddress1 As String
    Public Property BillingAddress2 As String

    Public Property BillingCity As String

    Public Property BillingState As String
    Public Property BillingZip As String
    Public Property BillingCountry As String
    Public Property BillingPhone As String
    Public Property BillingEmail As String
    Public Property BillingOrderComments As String
    Public Property PaidDate As DateTime '- Paid Date
    Public Property OrderTotal As Double '- amount paid
    Public Property TransactionID As String 'Payment Reference
    Public Property OrderDiscount As Double

    Public Property LineitemQuantityPurchased As Integer
    Public Property LineitemName As String '  - Photography Services, Tandem Skydive - 60 sec Freefall, Your Skydive Photos!, Tandem Skydive - 30 sec Freefall

    Public Property LineitemSKU As String
    Public Property LineitemPrice As Double
    Public Property LineitemCouponCode As String
    Public Property Notes As String

    'HEADER - 'OrderID, 
    'HEADER - 'OrderDate,
    'HEADER - 'PaymentMethod, 
    'HEADER - 'Captured(Auth.net),  
    'HEADER - 'ShippingMethodTitle,
    'HEADER - 'ShippingFirstName,
    'HEADER - 'ShippingLastName, 
    'HEADER - 'ShippingAddressLine1,
    'HEADER - 'ShippingAddressLine2,
    'HEADER - 'ShippingCity,
    'HEADER - 'ShippingState, 
    'HEADER - 'ShippingZip/Postcode,
    'HEADER - 'ShippingCountry, 
    'HEADER - 'BillingFirstName,
    'HEADER - 'BillingLastName,
    'HEADER - 'BillingAddressLine1,
    'HEADER - 'BillingAddressLine2,
    'HEADER - 'BillingCity, 
    'HEADER - 'BillingState, 
    'HEADER - 'BillingZip/Postcode,
    'HEADER - 'BillingCountry, 
    'HEADER - 'BillingPhoneNumber,
    'HEADER - 'BillingEmail, 
    'HEADER - 'Billingordercomments,
    ''QuantityOfitemspurchased,
    'ProductName, 
    'ProductSKU,
    'ProductID, 
    'ItempriceINCL.tax,
    'CouponCode, 
    'OrderDiscount, 
    'HEADER - 'OrderTotal(Auth.net), 
    'HEADER - 'PaidDate,
    'HEADER - 'Transaction ID

    'Shared Function FindShippingMethod(WooString) As shippingmethod
    '    'This will do a match for the online web store description to the internal shippingmethod
    'End Function
End Class

