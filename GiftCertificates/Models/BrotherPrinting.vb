Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Publisher
Imports Microsoft.Win32
Imports System.Drawing.Printing
Imports System.Reflection
Imports System.Runtime.InteropServices
Module BrotherPrinting
    Const BrotherPrinterType As String = "Brother QL-700"

    '********************************************************
    '   Open and Print a spcified file.
    '********************************************************
    Public Sub PrintLabel_BrotherPrinter(Name As String, Address As String, PrintLabelType As PrintLabelTypes)
        Dim PrintDocument As String = GetFileFolderDocumentPath("Address.lbx")
        If PrintLabelType = PrintLabelTypes.Address Then
            PrintDocument = GetFileFolderDocumentPath("Address.lbx")
        ElseIf PrintLabelType = PrintLabelTypes.ReturnAddress Then
            PrintDocument = GetFileFolderDocumentPath("ReturnAddress.lbx")
        ElseIf PrintLabelType = PrintLabelTypes.ReturnAddressDiscreet Then
            PrintDocument = GetFileFolderDocumentPath("ReturnAddressDiscreet.lbx")
        End If

        Try
            Dim objDoc As bpac.Document
            objDoc = CreateObject("bpac.Document")
            If (objDoc.Open(PrintDocument) <> False) Then

                Dim test = objDoc.SetPrinter(BrotherPrinterType, True)
                objDoc.GetObject("Name").Text = Name
                objDoc.GetObject("Address").Text = Address
                objDoc.StartPrint("", bpac.PrintOptionConstants.bpoDefault)
                objDoc.PrintOut(1, bpac.PrintOptionConstants.bpoDefault)
                objDoc.EndPrint()
                objDoc.Close()
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            Dim Sb As New System.Text.StringBuilder
            Sb.AppendLine("A problem occured trying to print to the label printer")
            Sb.AppendLine("Exception: " & ex.Message)
            MessageBox.Show(Sb.ToString, "Brother Printing", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub



    'Dim strPrinterAddress As String = "domain\machinename"
    'Dim strPath As String = "192.168.1.45" + " /D" + strPrinterAddress

    '  Private Declare Auto Function DocumentProperties Lib "winspool.drv" _
    '   (ByVal hWnd As IntPtr, ByVal hPrinter As IntPtr, ByVal pDeviceName As String,
    '    ByVal pDevModeOutput As IntPtr, ByVal pDevModeInput As IntPtr, ByVal fMode As Int32) As Integer

    '  Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterW" _
    '   (ByVal hPrinter As IntPtr, ByVal Level As Integer, ByVal pPrinter As IntPtr,
    '    ByVal cbBuf As Integer, ByRef pcbNeeded As Integer) As Integer

    '  Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" _
    '   (ByVal hPrinter As IntPtr, ByVal level As Integer, ByVal pPrinterInfoIn As IntPtr,
    '    ByVal command As Int32) As Boolean

    '  <DllImport("winspool.drv", EntryPoint:="OpenPrinterA", ExactSpelling:=True,
    '     SetLastError:=True, CallingConvention:=CallingConvention.StdCall,
    '     CharSet:=CharSet.Ansi)>
    '  Private Shared Function OpenPrinter(ByVal pPrinterName As String,
    'ByRef hPrinter As IntPtr, ByRef pDefault As PRINTER_DEFAULTS) As Boolean
    '  End Function

    '  <DllImport("winspool.drv", EntryPoint:="ClosePrinter", SetLastError:=True, ExactSpelling:=True,
    '   CallingConvention:=CallingConvention.StdCall)>
    '  Private Shared Function ClosePrinter(ByVal hPrinter As Int32) As Boolean
    '  End Function

    '  Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" _
    '   (ByVal pszBuffer As System.Text.StringBuilder, ByRef pcchBuffer As Int32) As Boolean

    '  Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" _
    '   (ByVal pszPrinter As String) As Boolean

    '  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    '   (ByVal hpvDest As IntPtr, ByVal hpvSource As IntPtr, ByVal cbCopy As Long)


    '  Private Structure PRINTER_DEFAULTS
    '      Dim pDatatype As String
    '      Dim pDevMode As Long
    '      Dim pDesiredAccess As Long
    '  End Structure

    '  Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
    '  Private Const PRINTER_ACCESS_ADMINISTER = &H4
    '  Private Const PRINTER_ACCESS_USE = &H8
    '  Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

    '  Private Const DM_IN_BUFFER As Integer = 8
    '  Private Const DM_IN_PROMPT As Integer = 4
    '  Private Const DM_OUT_BUFFER As Integer = 2

    '  Private Structure PRINTER_INFO_9
    '      Dim pDevMode As IntPtr
    '  End Structure



    '  <StructLayout(LayoutKind.Sequential)>
    '  Private Structure PRINTER_INFO_2
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pServerName As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pPrinterName As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pShareName As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pPortName As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pDriverName As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pComment As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pLocation As String

    '      Public pDevMode As IntPtr

    '      <MarshalAs(UnmanagedType.LPTStr)> Public pSepFile As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pPrintProcessor As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pDatatype As String
    '      <MarshalAs(UnmanagedType.LPTStr)> Public pParameters As String

    '      Public pSecurityDescriptor As IntPtr
    '      Public Attributes As Integer
    '      Public Priority As Integer
    '      Public DefaultPriority As Integer
    '      Public StartTime As Integer
    '      Public UntilTime As Integer
    '      Public Status As Integer
    '      Public cJobs As Integer
    '      Public AveragePPM As Integer
    '  End Structure


    '  <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    '  Public Structure DEVMODE
    '      <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public pDeviceName As String
    '      Public dmSpecVersion As Short
    '      Public dmDriverVersion As Short
    '      Public dmSize As Short
    '      Public dmDriverExtra As Short
    '      Public dmFields As Integer
    '      Public dmOrientation As Short
    '      Public dmPaperSize As Short
    '      Public dmPaperLength As Short
    '      Public dmPaperWidth As Short
    '      Public dmScale As Short
    '      Public dmCopies As Short
    '      Public dmDefaultSource As Short
    '      Public dmPrintQuality As Short
    '      Public dmColor As Short
    '      Public dmDuplex As Short
    '      Public dmYResolution As Short
    '      Public dmTTOption As Short
    '      Public dmCollate As Short
    '      <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmFormName As String
    '      Public dmUnusedPadding As Short
    '      Public dmBitsPerPel As Integer
    '      Public dmPelsWidth As Integer
    '      Public dmPelsHeight As Integer
    '      Public dmNup As Integer
    '      Public dmDisplayFrequency As Integer
    '      Public dmICMMethod As Integer
    '      Public dmICMIntent As Integer
    '      Public dmMediaType As Integer
    '      Public dmDitherType As Integer
    '      Public dmReserved1 As Integer
    '      Public dmReserved2 As Integer
    '      Public dmPanningWidth As Integer
    '      Public dmPanningHeight As Integer

    '  End Structure



    '  Private pOriginalDEVMODE As IntPtr


    'Public Sub SavePrinterSettings(ByVal printerName As String)
    '    Dim Needed As Integer
    '    Dim hPrinter As IntPtr
    '    If printerName = "" Then Exit Sub

    '    Try
    '        If OpenPrinter(printerName, hPrinter, Nothing) = False Then Exit Sub
    '        'Save original printer settings data (DEVMODE structure)
    '        Needed = DocumentProperties(Me.Handle, hPrinter, printerName, Nothing, Nothing, 0)
    '        Dim pFullDevMode As IntPtr = Marshal.AllocHGlobal(Needed) 'buffer for DEVMODE structure
    '        DocumentProperties(Me.Handle, hPrinter, printerName, pFullDevMode, Nothing, DM_OUT_BUFFER)
    '        pOriginalDEVMODE = Marshal.AllocHGlobal(Needed)
    '        CopyMemory(pOriginalDEVMODE, pFullDevMode, Needed)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    'Public Sub RestorePrinterSettings(ByVal printerName As String)
    '    Dim hPrinter As IntPtr
    '    If printerName = "" Then Exit Sub

    '    Try
    '        If OpenPrinter(printerName, hPrinter, Nothing) = False Then Exit Sub
    '        Dim PI9 As New PRINTER_INFO_9
    '        PI9.pDevMode = pOriginalDEVMODE
    '        Dim pPI9 As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(PI9))
    '        Marshal.StructureToPtr(PI9, pPI9, True)
    '        SetPrinter(hPrinter, 9, pPI9, 0&)
    '        Marshal.FreeHGlobal(pPI9) 'pOriginalDEVMODE will be free too
    '        ClosePrinter(hPrinter)


    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try


    'End Sub

    'Function GetPrinterName() As String
    '    Dim buffer As New System.Text.StringBuilder(256)
    '    Dim PrinterName As String = String.Empty
    '    'Get default printer's name
    '    GetDefaultPrinter(buffer, 256)
    '    PrinterName = buffer.ToString
    '    If PrinterName = "" Then
    '        MsgBox("Can't find default printer.")
    '    End If
    '    Return PrinterName
    'End Function

    'Sub SetTray(ByVal printerName As String, ByVal trayNumber As Integer)
    '    Dim hPrinter As IntPtr
    '    Dim Needed As Integer

    '    OpenPrinter(printerName, hPrinter, Nothing)

    '    'Get original printer settings data (DEVMODE structure)
    '    Needed = DocumentProperties(IntPtr.Zero, hPrinter, printerName, Nothing, Nothing, 0)
    '    Dim pFullDevMode As IntPtr = Marshal.AllocHGlobal(Needed) 'buffer for DEVMODE structure
    '    DocumentProperties(IntPtr.Zero, hPrinter, printerName, pFullDevMode, Nothing, DM_OUT_BUFFER)

    '    Dim pDevMode9 As DEVMODE = Marshal.PtrToStructure(pFullDevMode, GetType(DEVMODE))

    '    ' Tray change
    '    pDevMode9.dmDefaultSource = trayNumber

    '    Marshal.StructureToPtr(pDevMode9, pFullDevMode, True)

    '    Dim PI9 As New PRINTER_INFO_9
    '    PI9.pDevMode = pFullDevMode

    '    Dim pPI9 As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(PI9))
    '    Marshal.StructureToPtr(PI9, pPI9, True)
    '    SetPrinter(hPrinter, 9, pPI9, 0&)
    '    Marshal.FreeHGlobal(pPI9) 'pFullDevMode will be free too

    '    ClosePrinter(hPrinter)
    'End Sub


    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    ''Dim pbPrinter As Printer
    '    ''Dim pbMyPrinter As Printer

    '    Dim path = "c:\temp\files\TandemCertificate.pub"


    '    'Dim appDoc As New Microsoft.Office.Interop.Publisher.Document
    '    Dim appPub As Microsoft.Office.Interop.Publisher.Application = New Microsoft.Office.Interop.Publisher.Application
    '    Dim appDoc = appPub.Open(path)

    '    ''appPub.ActiveWindow.Visible = True
    '    ''appPub.ActiveDocument.PageSetup.PageHeight = appPub.MillimetersToPoints(55.7)
    '    ''appPub.ActiveDocument.PageSetup.PageWidth = appPub.MillimetersToPoints(39.4)


    '    ''pbPrinter.IsActivePrinter = True

    '    Dim Printer = "Canon MF4500w Series"
    '    Dim port = String.Empty
    '    'Try

    '    'Get collection of printers installed on users PC
    '    Dim objPrinters As Publisher.InstalledPrinters
    '    Dim intCurrentActiveIndex As Long
    '    Dim intWantedActiveIndex As Long

    '    'Redirect output to special printer if it exists for user
    '    objPrinters = appPub.InstalledPrinters
    '    Dim boolPDFFactoryExists As Boolean = False
    '    For Each objPr As Publisher.Printer In objPrinters
    '        If objPr.PrinterName.ToString = Printer Then '\\devsrv\Dave B Lexmark E352dn (MS)" -Testing
    '            boolPDFFactoryExists = True
    '            intWantedActiveIndex = objPr.Index
    '        End If
    '        If objPr.IsActivePrinter = True Then
    '            intCurrentActiveIndex = objPr.Index
    '        End If
    '        'MsgBox(objPr.PrinterName.ToString)
    '    Next
    '    'Printing outputs to whatever printer has IsActivePrinter set - Default Printer is initially set
    '    If boolPDFFactoryExists = True Then
    '        objPrinters(intWantedActiveIndex).IsActivePrinter = True
    '    End If

    '    Try
    '        MsgBox("Default printer" & GetPrinterName())
    '        SavePrinterSettings(Printer)
    '        SetTray(Printer, 2)

    '        '//Print
    '        appDoc.PrintOutEx()
    '        SetTray(Printer, 1)
    '        appDoc.PrintOutEx()

    '        RestorePrinterSettings(Printer)
    '    Catch ex As Exception

    '    End Try



    '    '    'appDoc.AdvancedPrintOptions.




    '    '    Dim _paperSource = "TRAY 2" '; // Printer Tray
    '    '    Dim _paperName = "8x17" '; // Printer paper name



    '    '    'Dim PapSource As PaperSource
    '    '    'Dim intx As Integer = 0
    '    '    'With appDoc.PrinterSettings
    '    '    '    For Each PapSource In .PaperSources
    '    '    '        If PapSource.Kind = PaperSourceKind.Lower Then
    '    '    '            appDoc.DefaultPageSettings.PaperSource = appDoc.PrinterSettings.PaperSources(intx)
    '    '    '            .DefaultPageSettings.PaperSource = appDoc.PrinterSettings.PaperSources(intx)
    '    '    '            Exit For
    '    '    '        End If
    '    '    '        intx += 1
    '    '    '    Next
    '    '    'End With


    '    '    Dim x As New PrinterSettings

    '    '    For Each d As PaperSource In x.PaperSources

    '    '        MsgBox(d.SourceName & " " & d.RawKind)

    '    '    Next


    '    '    appDoc.PrintOutEx()
    '    '    appDoc.Save()
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message)
    '    'End Try



    '    ''' Set ActivePrinter if not already set
    '    ''If Not appPub.InstalledPrinters.StartsWith(Printer) Then
    '    ''    ' Get current concatenation string ('on' in enlgish, 'op' in dutch, etc..)
    '    ''    Dim split = appPub.ActivePrinter.Split(" "c)

    '    ''    If split.Length >= 3 Then
    '    ''        appPub.ActivePrinter = String.Format("{0} {1} {2}", Printer, split(split.Length - 2), port)
    '    ''    End If
    '    ''End If

    '    appDoc.Close()
    '    appPub = Nothing



    'End Sub


    '//This will be the module that contains the address printing methods

    Public Property Name As String = ""
    Public Property Address1 As String = ""
    Public Property Address2 As String = ""

    Public Property City As String = ""
    Public Property State As String = ""
    Public Property Zip As String = ""




End Module
