

Public Enum DeliveryOptions
    Email
    USMail
    USDiscreet
    InOffice
End Enum

Public Enum PointOfSale
    Online
    PhoneInPerson
End Enum

Public Enum PaymentMethod
    Cash
    CreditCard
    Online
    Online_authorize_net_cim_credit_card

End Enum

Public Enum HearAbout
    Unknown
    WordOfMouth
    Facebook
    Email
    Tripadvisor
    Yelp
    Google
End Enum

Public Enum FilterStates

    Entered
    Processing
    Completed
    ProcessAndCompleteOnly
    All
    Unspecified
    Incomplete
End Enum

Public Enum CertificateStatus
    Entered = 0
    Processing = 1
    Completed = 2
End Enum

Public Enum Product_ItemType
    Item_TDM10K = 1
    Item_TDM12K
    Item_TDM10KVID
    Item_TDM12KVID
    Item_VID
End Enum

Public Enum PrintLabelTypes
    Address
    ReturnAddress
    ReturnAddressDiscreet
End Enum


Public Module Lists
    Public Const NullDate As String = "12/30/1899 12:00:00 AM"
    Public Const BCC_Email = "gcprocess@spottysworld.com"

    Function RetrievePaymentMethodList() As List(Of KeyValuePair(Of Integer, String))
        Dim L As New List(Of KeyValuePair(Of Integer, String))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.Cash, "Cash"))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.CreditCard, "Credit Card"))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.Online, "Online"))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.Online_authorize_net_cim_credit_card, "authorize_net_cim_credit_card"))
        Return L
    End Function

    Function RetrieveHearAboutList() As List(Of KeyValuePair(Of Integer, String))
        Dim L As New List(Of KeyValuePair(Of Integer, String))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Unknown, "Unknown"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.WordOfMouth, "Word Of Mouth"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Facebook, "Facebook"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Email, "Email"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Tripadvisor, "Tripadvisor"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Google, "Google"))
        L.Add(New KeyValuePair(Of Integer, String)(HearAbout.Yelp, "Yelp"))
        Return L
    End Function

    Function RetrievePointOfSale() As List(Of KeyValuePair(Of Integer, String))
        Dim L As New List(Of KeyValuePair(Of Integer, String))
        L.Add(New KeyValuePair(Of Integer, String)(PointOfSale.Online, "Online"))
        L.Add(New KeyValuePair(Of Integer, String)(PointOfSale.PhoneInPerson, "Office"))
        Return L
    End Function

    Function RetrieveDelivery() As List(Of KeyValuePair(Of Integer, String))
        Dim L As New List(Of KeyValuePair(Of Integer, String))
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.Email, "Email"))
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.USMail, "US Mail"))
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.USDiscreet, "US Mail Discrete"))
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.InOffice, "In Person"))

        Return L
    End Function

    Function GetDeliveryOptionString(i As Integer) As String
        Dim DeliveryOption As String = ""
        Select Case i
            Case DeliveryOptions.Email
                DeliveryOption = "Email"
            Case DeliveryOptions.USMail
                DeliveryOption = "US Mail"
            Case DeliveryOptions.USDiscreet
                DeliveryOption = "US Mail Discrete"
            Case DeliveryOptions.InOffice
                DeliveryOption = "In Office"
        End Select
        Return DeliveryOption

    End Function

    Function GetPointOfSaleString(i As Integer) As String
        Dim POS As String = ""
        Select Case i
            Case PointOfSale.Online
                POS = "Online"
            Case PointOfSale.PhoneInPerson
                POS = "In Person"
        End Select
        Return POS
    End Function

    Function GetStatusString(i As Integer) As String
        Dim Status As String = ""
        Select Case i
            Case CertificateStatus.Entered
                Status = "Entered"
            Case CertificateStatus.Processing
                Status = "In Process"
            Case CertificateStatus.Completed
                Status = "Completed"
        End Select
        Return Status
    End Function

End Module