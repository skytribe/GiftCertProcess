

Public Enum DeliveryOptions
    Email
    USMail
    USMailDiscrete
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
    Processed
    Completed
    ProcessAndCompleteOnly
    All
    Unspecified
End Enum

Public Enum CertificateStatus
    Entered = 0
    Processed = 1
    Completed = 2
End Enum

Module Lists
    Function RetrievePaymentMethodList() As List(Of KeyValuePair(Of Integer, String))

        Dim L As New List(Of KeyValuePair(Of Integer, String))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.Cash, "Cash"))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.CreditCard, "Credit Card"))
        L.Add(New KeyValuePair(Of Integer, String)(PaymentMethod.Online, "Online"))
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
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.USMailDiscrete, "US Mail Discrete"))
        L.Add(New KeyValuePair(Of Integer, String)(DeliveryOptions.InOffice, "In Person"))

        Return L
    End Function


End Module