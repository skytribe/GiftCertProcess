Public Class PreferredNameFirstNamePair
    Public Property preferedname As String
    Public Property firstname As String
    Public Property scust As Integer

    Public Sub New(c As Integer, p As String, f As String)
        scust = c
        preferedname = p
        firstname = f
    
    End Sub

End Class
