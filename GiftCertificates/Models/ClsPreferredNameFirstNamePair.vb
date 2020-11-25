Public Class ClsPreferredNameFirstNamePair
    Public Property PreferredName As String
    Public Property FirstName As String
    Public Property sCustId As Integer

    Public Sub New(c As Integer, p As String, f As String)
        sCustId = c
        PreferredName = p
        FirstName = f
    End Sub

End Class
