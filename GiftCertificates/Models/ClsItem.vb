''' <summary>
''' Used to identify a single Product Line Item on a gift certificate
''' </summary>
Public Class ClsItem
    Sub New(id As Integer)
        Me.ItemId = id
        Me.Quantity = 0
        Me.JRCustomerID = 0
        Me.JRCertificateID = 0
        Me.JRItemId = 0
        Me.Price = 0

    End Sub

    Public Property ItemId As Integer
    Public Property Quantity As Integer
    Public Property Price As Double
    Public Property JRCustomerID As Integer
    Public Property JRCertificateID As Integer
    Public Property JRItemId As Integer

End Class
