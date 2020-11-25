Module ModLogging

    Public Sub LogError(msg As String, ex As Exception)
        Dim newtext = ""
        If ex IsNot Nothing Then
            newtext = Environment.NewLine & String.Format("{0:g}:{1}", Now, msg & Environment.NewLine & ex.Message) & "**********************************************************"

        Else
            newtext = Environment.NewLine & String.Format("{0:g}:{1}", Now, msg & Environment.NewLine) & "**********************************************************"
        End If

        My.Computer.FileSystem.WriteAllText("Log.txt", newtext, True)
    End Sub

    Public Sub LogError(msg As String, st As String)
        Dim newtext = ""

        newtext = Environment.NewLine & String.Format("{0:g}:{1}", Now, msg & Environment.NewLine & st) & "**********************************************************"


        My.Computer.FileSystem.WriteAllText("Log.txt", newtext, True)
    End Sub

End Module
