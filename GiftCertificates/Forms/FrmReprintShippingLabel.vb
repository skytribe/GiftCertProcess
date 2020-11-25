Imports System.IO
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports Syncfusion.Windows.Forms.Grid
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events

Public Class FrmReprintShippingLabel

    Public Property CurrentGiftCertificate As ClsGiftCertificate2
    Public Property CurrentFilter As FilterStates

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All
                Case "rdoboth"
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

            End Select
            Dim dteEntry As Date = SfDateTimeEdit1.Value
            SfDataGrid1.SelectedItems.Clear()

            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            If fl IsNot Nothing AndAlso fl.Count > 0 Then
                SfDataGrid1.DataSource = fl
            Else
                'BlankDisplayFields()
                Button1.Enabled = False
                Button4.Enabled = False
                Button3.Enabled = False

                SfDataGrid1.DataSource = Nothing
            End If



        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Public Function GetCertificatesToProcess(entrydate As Date, Filter As FilterStates) As List(Of ClsGiftCertificate2)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate2)
            gclist1 = RetrieveGCOrdersFromQueue(entrydate)

            Dim FilteredList As List(Of ClsGiftCertificate2)

            Select Case Filter
                Case FilterStates.Entered
                    CurrentFilter = FilterStates.Entered
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Select i1).ToList

                Case FilterStates.Processing
                    CurrentFilter = FilterStates.Processing
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Processing Select i1).ToList



                Case FilterStates.Completed
                    CurrentFilter = FilterStates.Completed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Select i1).ToList

                Case FilterStates.ProcessAndCompleteOnly
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
                Case FilterStates.All
                    FilteredList = (From i1 In gclist1 Select i1).ToList
                Case FilterStates.Incomplete
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
            End Select

            Return FilteredList

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function

    Public Function GetCertificatesToProcess(name As String, Filter As FilterStates) As List(Of ClsGiftCertificate2)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate2)
            gclist1 = RetrieveGCOrdersFromQueue(name)

            Dim FilteredList As List(Of ClsGiftCertificate2)

            Select Case Filter
                Case FilterStates.Entered
                    CurrentFilter = FilterStates.Entered
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Select i1).ToList

                Case FilterStates.Processing
                    CurrentFilter = FilterStates.Processing
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Processing Select i1).ToList



                Case FilterStates.Completed
                    CurrentFilter = FilterStates.Completed
                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Select i1).ToList

                Case FilterStates.ProcessAndCompleteOnly
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Completed Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
                Case FilterStates.All
                    FilteredList = (From i1 In gclist1 Select i1).ToList
                Case FilterStates.Incomplete
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly

                    FilteredList = (From i1 In gclist1 Where i1.GC_Status = CertificateStatus.Entered Or
                                                          i1.GC_Status = CertificateStatus.Processing
                                    Select i1).ToList
            End Select

            Return FilteredList

        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try

    End Function



    Private Sub FrmReprintShippingLabel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SfDateTimeEdit1.Value = Now
        Button1.Enabled = False
        Button4.Enabled = False
        Button3.Enabled = False

        initializegrid()

    End Sub

    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate2)
                Button1.Enabled = True
                Button4.Enabled = True
                Button3.Enabled = True
            Else
                Button1.Enabled = False
                Button4.Enabled = False
                Button3.Enabled = False
            End If
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sb As New System.Text.StringBuilder
        Dim addressCityStateZip As String = ""
        Try
            If String.IsNullOrEmpty(CurrentGiftCertificate.Shipping_Address.Address1) = False Then
                sb.AppendLine(CurrentGiftCertificate.Shipping_Address.Address1.Trim)
                If String.IsNullOrEmpty(CurrentGiftCertificate.Shipping_Address.Address2) = False Then
                    sb.AppendLine(CurrentGiftCertificate.Shipping_Address.Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", CurrentGiftCertificate.Shipping_Address.City.Trim.ToProperCase, CurrentGiftCertificate.Shipping_Address.State.Trim.ToProperCase, CurrentGiftCertificate.Shipping_Address.Zip.Trim)
                sb.AppendLine(addressCityStateZip)
            Else
                sb.AppendLine(CurrentGiftCertificate.Billing_Address.Address1.Trim)
                If String.IsNullOrEmpty(CurrentGiftCertificate.Billing_Address.Address2) = False Then
                    sb.AppendLine(CurrentGiftCertificate.Billing_Address.Address2.Trim)
                End If
                addressCityStateZip = String.Format("{0}  {1}  {2}", CurrentGiftCertificate.Billing_Address.City.Trim.ToProperCase, CurrentGiftCertificate.Billing_Address.State.Trim.ToProperCase, CurrentGiftCertificate.Billing_Address.Zip.Trim)
                sb.AppendLine(addressCityStateZip)
            End If

            PrintLabel_BrotherPrinter(CurrentGiftCertificate.Purchaser_Name, sb.ToString, PrintLabelTypes.Address)



        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            PrintLabel_BrotherPrinter("Sender", My.Settings.ReturnAddress, PrintLabelTypes.ReturnAddress)
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            PrintLabel_BrotherPrinter("Sender", My.Settings.ReturnAddressDiscreet, PrintLabelTypes.ReturnAddressDiscreet)
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoall"
                    CurrentFilter = FilterStates.All
                Case "rdoboth"
                    CurrentFilter = FilterStates.Incomplete

            End Select
            SfDataGrid1.SelectedItems.Clear()

            Dim fl = GetCertificatesToProcess(TextBox1.Text, CurrentFilter)
            SfDataGrid1.DataSource = fl
        Catch ex As Exception
            Dim m1 As MethodBase = MethodBase.GetCurrentMethod()
            Dim methodName = String.Format("{0}.{1}", m1.ReflectedType.Name, m1.Name)
            LogError(methodName, ex)

            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged

        If RadioButton1.Checked Then
            Panel4.Enabled = True
            Panel5.Enabled = False

        Else
            Panel4.Enabled = False
            Panel5.Enabled = True

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub




    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        Using file = System.IO.File.Create("FrmShippingReprint.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using
    End Sub

    Public Sub initializegrid()
        If File.Exists("FrmShippingReprint.xml") Then
            Try
                Using file = System.IO.File.Open("FrmShippingReprint.xml", System.IO.FileMode.Open)
                    Me.SfDataGrid1.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If
    End Sub

End Class