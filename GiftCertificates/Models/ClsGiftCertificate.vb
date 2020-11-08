Public Class ClsGiftCertificate
    Public Property ID As Integer
    Public Property GC_Number As String = "" 'Online Number
    Public Property GC_DateEntered As Date
    Public Property HearAbout As HearAbout
    Public Property delivery As DeliveryOptions  'EMail, USMail, US Mail Discrete, In-Office
    Public Property PointOfSale As PointOfSale  'On-Line, Phone/Inperson
    Public Property Notes As String = ""
    Public Property GC_Authorization As String = ""
    Public Property GC_Username As String = ""
    Public Property GC_Status As CertificateStatus
    Public Property ShopifyName As String = ""

    Public Property PaymentMethod As PaymentMethod
    Public Property PaymentNotes As String = ""

    '// Purchaser
    Public Property Purchaser_FirstName As String
    Public Property Purchaser_LastName As String
    Public Property Purchaser_Address1 As String = ""
    Public Property Purchaser_Address2 As String = ""
    Public Property Purchaser_City As String = ""
    Public Property Purchaser_State As String = ""
    Public Property Purchaser_Zip As String = ""
    Public Property Purchaser_Phones1 As String = ""
    Public Property Purchaser_Phones2 As String = ""
    Public Property Purchaser_Email As String = ""
    Public Property Purchaser_PersonalizedTo As String = ""

    Public ReadOnly Property Purchaser_Name() As String
        Get
            Return String.Format("{0} {1}", Purchaser_FirstName.Trim, Purchaser_LastName.Trim)
        End Get
    End Property

    ''Recipient
    Public Property Recipient_FirstName As String = ""
    Public Property Recipient_LastName As String = ""
    Public Property Recipient_Address1 As String = ""
    Public Property Recipient_Address2 As String = ""
    Public Property Recipient_City As String = ""
    Public Property Recipient_State As String = ""
    Public Property Recipient_Zip As String = ""
    Public Property Recipient_Phones1 As String = ""
    Public Property Recipient_Phones2 As String = ""
    Public Property Recipient_Email As String = ""
    Public Property Recipient_PersonalizedFrom As String = ""
    Public ReadOnly Property Recipient_Name() As String
        Get
            Return String.Format("{0} {1}", Recipient_FirstName.Trim, Recipient_LastName.Trim)
        End Get
    End Property

    'Items  - storing the amount which allows for recalculation of totals
    'based upon price when transaction was original made to determine gift certificate value
    'as online allows multiple quantities of each item
    Public Property Item_Tandem10k As Integer
    Public Property Item_Tandem10kAmount As Double
    Public Property Item_Tandem12k As Integer
    Public Property Item_Tandem12kAmount As Double
    Public Property Item_Video As Integer
    Public Property Item_VideoAmount As Double
    Public Property Item_Other As Boolean
    Public Property Item_OtherAmount As Double
    Public Property GC_CalculateTotal As Boolean


    'JR Items
    Public Property JR_PurchaseID As Integer
    Public Property JR_RecipientID As Integer

    Sub RecalculateTotal()
        Try
            Dim totalamount As Double = 0

            totalamount = Item_Tandem10kAmount + Item_Tandem12kAmount + Item_VideoAmount + Item_OtherAmount

            Me.GC_CalculateTotal = totalamount
        Catch ex As Exception

        End Try

    End Sub
End Class
