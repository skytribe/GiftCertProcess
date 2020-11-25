
Public Class ClsJumpRunItem
    Public Property ID As Integer
    Public Property Description As String
    Public Property Price As Double

    Sub New()

    End Sub
    Sub New(i As Integer, d As String, p As Double)
        ID = i
        Description = d
        Price = p
    End Sub
End Class
