

Imports System.Drawing.Printing

Public Class FrmSetDefaultPrinter

    Private Sub FrmSetDefaultPrinter_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim strInstalledPrinters As String

        'check if there is installed printer
        If PrinterSettings.InstalledPrinters.Count = 0 Then
            MsgBox("No printer installed")
            Exit Sub
        End If

        'display installed printer into combobox list item
        For Each strInstalledPrinters In PrinterSettings.InstalledPrinters
            ComboBox1.Items.Add(strInstalledPrinters)
        Next strInstalledPrinters

        If String.IsNullOrEmpty(My.Settings.DefaultPrinterToUse) Then
            'Display current default printer on combobox texts
            ComboBox1.Text = GetDefaultPrinter()
        Else
            ComboBox1.Text = My.Settings.DefaultPrinterToUse
        End If
        Button2.Text = "Get Tray Options"


        Button1.Text = "Set Default Printer"


    End Sub


    'Function to set a printer as default
    Function SetDefaulPrinter(ByVal strPrinterName As String, Optional strDefaultTray As String = "") As Boolean
        Dim strCurrPrinter As String
        Dim WsNetwork As Object
        Dim prntDoc As New PrintDocument

        strCurrPrinter = prntDoc.PrinterSettings.PrinterName
        WsNetwork = Microsoft.VisualBasic.CreateObject("WScript.Network")

        Try
            WsNetwork.SetDefaultPrinter(strPrinterName)
            WsNetwork.set
            prntDoc.PrinterSettings.PrinterName = strPrinterName
            If strDefaultTray = "" Then
            Else
                ' step through the available paper sources and set it to the right one
                For psrc = 0 To prntDoc.PrinterSettings.PaperSources.Count - 1
                    If prntDoc.PrinterSettings.PaperSources(psrc).SourceName = strDefaultTray Then
                        prntDoc.PrinterSettings.DefaultPageSettings.PaperSource = prntDoc.PrinterSettings.PaperSources.Item(psrc)
                    End If
                Next

            End If
            'set default if selected printer name is a valid installed printer
            If prntDoc.PrinterSettings.IsValid Then
                Return True
            Else
                WsNetwork.SetDefaultPrinter(strCurrPrinter)
                Return False
            End If
        Catch ex As Exception
            WsNetwork.SetDefaultPrinter(strCurrPrinter)
            Return False
        Finally
            WsNetwork = Nothing
            prntDoc = Nothing
        End Try
    End Function
    Function GetDefaultPrinter() As String
        Dim CurrentDefault As String = ""

        Dim prntDoc As New PrintDocument

        'check if there is installed printer
        If PrinterSettings.InstalledPrinters.Count = 0 Then
            MsgBox("No printer installed")
            Return CurrentDefault

        End If


        Return prntDoc.PrinterSettings.PrinterName

    End Function
    Private Sub Button1_Click(ByVal sender As System.Object,
                ByVal e As System.EventArgs) Handles Button1.Click
        If SetDefaulPrinter(ComboBox1.Text) = True Then
            MsgBox("Printer default " & ComboBox1.Text)
            My.Settings.DefaultPrinterToUse = ComboBox1.Text
            My.Settings.Save()
        Else
            MsgBox("Printer name " & ComboBox1.Text & " is not valid!")
        End If

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            Dim x As New PrinterSettings
            Dim printDoc As New PrintDocument
            Dim pkSource As PaperSource
            ComboBox2.Items.Clear()
            For i = 0 To printDoc.PrinterSettings.PaperSources.Count - 1
                pkSource = printDoc.PrinterSettings.PaperSources(i)
                ComboBox2.Items.Add(pkSource.SourceName)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try




        'For Each d As PaperSource In x.PaperSources

        '    Debug.Print(d.SourceName & " " & d.RawKind)
        '    ComboBox2.Items.Add(pkSource)


        'Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            My.Settings.Reload()
            My.Settings.DefaultPrinterTray = ComboBox2.Text
            My.Settings.Save()
            Dim printer = GetDefaultPrinter()
            SetDefaulPrinter(printer, ComboBox2.Text)
        Catch ex As Exception

        End Try

    End Sub
End Class