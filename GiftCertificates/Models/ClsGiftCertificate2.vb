Public Class ClsGiftCertificate2
    Sub New()
        Billing_Address = New ClsAddress
        Shipping_Address = New ClsAddress
        Item1 = New ClsItem(1)
        Item2 = New ClsItem(2)
        Item3 = New ClsItem(3)
        Item4 = New ClsItem(4)
        Item5 = New ClsItem(5)
    End Sub

    Public Property ID As Integer
    Public Property Online_OrderNumber As String = "" 'Online Number
    Public Property GC_DateEntered As Date
    Public Property HearAbout As HearAbout
    Public Property delivery As DeliveryOptions  'EMail, USMail, US Mail Discrete, In-Office
    Public Property PointOfSale As PointOfSale  'On-Line, Phone/Inperson
    Public Property Notes As String = ""
    Public Property GC_Authorization As String = ""
    Public Property GC_Username As String = ""
    Public Property GC_Status As CertificateStatus
    Public Property OrderId As String = ""
    Public Property PaymentMethod As PaymentMethod
    Public Property PaymentNotes As String = ""

    '// Purchaser
    Public Property Purchaser_FirstName As String
    Public Property Purchaser_LastName As String
    Public Property Billing_Address As ClsAddress
    'Shipping Address
    Public Property Shipping_Address As ClsAddress


    Public ReadOnly Property Purchaser_Name() As String
        Get
            Return String.Format("{0} {1}", Purchaser_FirstName.Trim, Purchaser_LastName.Trim)
        End Get
    End Property

    Public Property PersonalizedFrom As String = ""
    Public Property PersonalizedTo As String = ""


    'Items  - storing the amount which allows for recalculation of totals
    'based upon price when transaction was original made to determine gift certificate value
    'as online allows multiple quantities of each item

    Public Property Item1 As ClsItem
    Public Property Item2 As ClsItem
    Public Property Item3 As ClsItem
    Public Property Item4 As ClsItem
    Public Property Item5 As ClsItem

    Public Property GC_TotalAmount As Double
    Public Property GC_TotalDiscount As Double
    Public Property GC_DiscountCode As String

    'JR Items
    Public Property JR_PurchaserID As Integer

    Public Property GC_ProcessedDate As Date



End Class
