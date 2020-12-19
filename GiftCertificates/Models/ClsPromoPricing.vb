Public Class ClsPromoPricing
    Property ID As Integer
    Property PromoDescription As String

    Property ItemCode1 As Integer
    Property ItemPrice1 As Double
    Property ItemCode2 As Integer
    Property ItemPrice2 As Double
    Property ItemCode3 As Integer
    Property ItemPrice3 As Double
    Property ItemCode4 As Integer
    Property ItemPrice4 As Double
    Property ItemCode5 As Integer
    Property ItemPrice5 As Double

    Property ItemDiscount1 As Double
    Property ItemDiscount2 As Double
    Property ItemDiscount3 As Double
    Property ItemDiscount4 As Double
    Property ItemDiscount5 As Double


    Public ReadOnly Property TotalPrice1() As Double
        Get
            Return ItemPrice1 - Math.Abs(ItemDiscount1)
        End Get
    End Property

    Public ReadOnly Property TotalPrice2() As Double
        Get
            Return ItemPrice2 - Math.Abs(ItemDiscount2)
        End Get
    End Property

    Public ReadOnly Property TotalPrice3() As Double
        Get
            Return ItemPrice3 - Math.Abs(ItemDiscount3)
        End Get
    End Property

    Public ReadOnly Property TotalPrice4() As Double
        Get
            Return ItemPrice4 - Math.Abs(ItemDiscount4)
        End Get
    End Property

    Public ReadOnly Property TotalPrice5() As Double
        Get
            Return ItemPrice5 - Math.Abs(ItemDiscount5)
        End Get
    End Property

    Property Status As Integer

End Class
