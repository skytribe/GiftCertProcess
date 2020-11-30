Imports System.Reflection
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Publisher

Public Module ModPublisherIntegration

    Public Const MSPublisherProcessName = "MSPUB"
    Public Const PublisherCertificateDocument = "SSGC10-2020.PUB"
    Public Const MailMergeSourceDocument = "gcMailmerge.txt"
    Public Sub PrintOrderCertificates(certificate As ClsGiftCertificate2,
                                      destination As Microsoft.Office.Interop.Publisher.PbMailMergeDestination,
                                      Optional ToBePrinted As Boolean = True,
                                      Optional IsReprint As Boolean = False)
        Try

            Dim PrintCertificateDocument As String = GetFullFilePathForFilesFolderContent(PublisherCertificateDocument)

            '//EACH  CERTIFICATES IN A SINGLE DOCUMENT FOR EMAIL ATTACHMENT PURPOSE
            If ToBePrinted = False Then
                '//AL CERTIFICATES IN A SINGLE DOCUMENT FOR EMAIL PURPOSE
                Dim certificatestoprint1 = GenerateMailMergeFileForOrder(certificate, MailMergeSourceDocument)

                KillExistingPublisherInstances()

                Dim Application As Publisher.Application = New Publisher.Application()
                Application.ActiveWindow.Visible = True
                Dim Path As String = PrintCertificateDocument


                Try
                    Application.Open(Path, False, True)

                    If destination = PbMailMergeDestination.pbSendToPrinter Then
                        Application.ActiveDocument.MailMerge.Execute(True, Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbSendToPrinter)
                    ElseIf destination = PbMailMergeDestination.pbMergeToNewPublication Then
                        Application.ActiveDocument.MailMerge.Execute(True, Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication)

                        Dim filename = String.Format("Certificate {0}-{1} {2} {3}.pdf", certificate.ID, "ALL", certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
                        Dim d1 As Document = Nothing
                        If Application.Documents.Count() > 1 Then
                            Dim MergedDocumentIndex As Integer = 1
                            For Each i As Document In Application.Documents
                                If i.Name.ToUpper().Contains(PublisherCertificateDocument.ToUpper) = False Then
                                    d1 = Application.Documents(MergedDocumentIndex)

                                    Exit For
                                End If
                                MergedDocumentIndex += 1
                            Next
                            If d1 IsNot Nothing Then
                                For i = 1 To certificatestoprint1


                                    Dim filename1 = String.Format("Certificate {0}-{1} {2} {3}.pdf", certificate.ID, i, certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
                                    d1.ExportAsFixedFormat(Format:=PbFixedFormatType.pbFixedFormatTypePDF, Filename:=System.IO.Path.Combine(My.Settings.PDFOutputFolder, filename1), From:=i, [To]:=i)
                                Next
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
                    Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
                    LogError(methodName, ex)
                End Try


            Else
                '//AL CERTIFICATES IN A SINGLE DOCUMENT FOR PRINT PURPOSE
                Dim certificatestoprint = GenerateMailMergeFileForOrder(certificate, MailMergeSourceDocument)

                KillExistingPublisherInstances()

                Dim Application As Publisher.Application = New Publisher.Application()
                Application.ActiveWindow.Visible = True
                Dim Path As String = PrintCertificateDocument

                Try
                    Application.Open(Path, False, True)
                    If destination = PbMailMergeDestination.pbSendToPrinter Then
                        Application.ActiveDocument.MailMerge.Execute(True, Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbSendToPrinter)
                    ElseIf destination = PbMailMergeDestination.pbMergeToNewPublication Then
                        Application.ActiveDocument.MailMerge.Execute(True, Microsoft.Office.Interop.Publisher.PbMailMergeDestination.pbMergeToNewPublication)
                        Dim filename = String.Format("Certificate {0}-{1} {2} {3}.pdf", certificate.ID, "ALL", certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
                        If Application.Documents.Count() = 2 Then
                            Dim MergedDocument As Document = Application.Documents(2)
                            MergedDocument.ExportAsFixedFormat(Format:=PbFixedFormatType.pbFixedFormatTypePDF, Filename:=System.IO.Path.Combine(My.Settings.PDFOutputFolder, filename))
                        End If
                    End If
                Catch ex As Exception
                    Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
                    Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
                    LogError(methodName, ex)
                End Try
            End If

            KillExistingPublisherInstances() '//When closing

            '//We don't want to update if its a reprint
            If IsReprint = False Then
                UpdateGCOrderStatus(certificate, CertificateStatus.Completed, GetBusinessDate)
            End If

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub KillExistingPublisherInstances()
        If CheckIfRunning(MSPublisherProcessName) Then
            Try
                For Each p In Process.GetProcessesByName(MSPublisherProcessName)
                    p.Kill()
                Next
            Catch ex As Exception
                Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
                Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
                LogError(methodName, ex)
            End Try
        End If
    End Sub

    Private Function GetFullFilePathForFilesFolderContent(FileName As String) As String
        Dim appFolder = My.Application.Info.DirectoryPath
        Dim MasterFileFolder = System.IO.Path.Combine(appFolder, "Files")
        Dim CertificateToPrint = System.IO.Path.Combine(MasterFileFolder, FileName)
        Dim PrintCertificateDocument = System.IO.Path.Combine(appFolder, CertificateToPrint)
        Return PrintCertificateDocument
    End Function

    Private Function GenerateMailMergeFileForOrder(certificate As ClsGiftCertificate2, mmsource As String) As Integer
        Dim icount As Integer = 0
        Try

            Dim sb As New StringBuilder

            Dim lstItems = RetrieveGCOrdersLineItems(certificate)

            sb.AppendLine("Number" & "," & "OrderDate" & "," & "From" & "," & "Description" & "," & "Authorized")
            Dim certificatestoprint As Integer = 0

            Dim ActualAltitude = ""
            Dim name As String = ""
            Dim GCFrom As String

            If String.IsNullOrEmpty(certificate.PersonalizedFrom) Then
                GCFrom = String.Format("{0} {1}", certificate.Purchaser_FirstName.Trim, certificate.Purchaser_LastName.Trim)
            Else
                GCFrom = certificate.PersonalizedFrom.Trim
            End If

            For Each i In lstItems
                Dim st1 = String.Format("{0},{1},{2},{3},{4}", i.JumpRunCertificateNumber,
                                                                i.OrderDate,
                                                                GCFrom,
                                                                i.Description,
                                                                GetAuthorizerCode(i.Authorizer)
)

                sb.AppendLine(st1)
                icount += 1
            Next

            mmsource = System.IO.Path.Combine(My.Application.Info.DirectoryPath, mmsource)
            Dim StdEncoder As New System.Text.ASCIIEncoding
            Dim orfWriter As System.IO.StreamWriter = New System.IO.StreamWriter(mmsource, False, StdEncoder)
            orfWriter.Write(sb.ToString)
            orfWriter.Close()


        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show(ex.Message)
        End Try

        Return icount
    End Function

    '    Private Function GenerateMailMergeFile(Cert As ClsPrintCertificateDetails, mmsource As String) As Integer
    '        Try
    '            Dim sb As New StringBuilder

    '            sb.AppendLine("Number" & "," & "OrderDate" & "," & "From" & "," & "Description" & "," & "Authorized")
    '            Dim certificatestoprint As Integer = 0

    '            Dim ActualAltitude = ""
    '            Dim name As String = ""
    '            Dim GCFrom As String


    '            GCFrom = Cert.PersoanlizedFrom.Trim


    '            Dim st1 = String.Format("{0},{1},{2},{3},{4}", Cert.JumpRunCertificateNumber,
    '                                                                Cert.OrderDate,
    '                                                                GCFrom,
    '                                                                Cert.Description,
    '                                                                Cert.Authorizer
    ')

    '            sb.AppendLine(st1)

    '            mmsource = System.IO.Path.Combine(My.Application.Info.DirectoryPath, mmsource)
    '            Dim StdEncoder As New System.Text.ASCIIEncoding
    '            Dim orfWriter As System.IO.StreamWriter = New System.IO.StreamWriter(mmsource, False, StdEncoder)
    '            orfWriter.Write(sb.ToString)
    '            orfWriter.Close()

    '        Catch ex As Exception
    '            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
    '            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
    '            LogError(methodName, ex)

    '            MessageBox.Show(ex.Message)
    '        End Try

    '        Return 1
    '    End Function

End Module
