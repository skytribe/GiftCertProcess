﻿Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events

Public Class FrmReprint

    Public Property CurrentGiftCertificate As ClsGiftCertificate2
    Public Property CurrentFilter As FilterStates

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Dim dteEntry As Date = SfDateEntry.Value
            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            SfDataGrid1.DataSource = fl

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate2)
                BtnReprint.Enabled = True
            Else
                BtnReprint.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("SfDataGrid1_SelectionChanged:", ex)
        End Try
    End Sub

    Private Sub BtnReprint_Click(sender As Object, e As EventArgs) Handles BtnReprint.Click
        Try
            If CurrentGiftCertificate IsNot Nothing Then

                If CurrentGiftCertificate.GC_Status = CertificateStatus.Entered Then
                    MessageBox.Show("The certificate hasnt yet been processed, cannot reprint until it has been processed", "Reprint certificate", MessageBoxButtons.OK)
                Else
                    Dim DeliverCertificateForm As New FrmProcessPrint
                    DeliverCertificateForm.Certificate = CurrentGiftCertificate
                    DeliverCertificateForm.IsReprint = True

                    DeliverCertificateForm.ShowDialog()

                    'If CurrentGiftCertificate.GC_Status = CertificateStatus.Processing Then
                    '    'We can update the status to completed
                    '    UpdateGCOrderStatus(CurrentGiftCertificate, CertificateStatus.Completed)
                    'End If

                    Me.Close()

                End If
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub FrmReprint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            'Setup Datagrid
            SetupForm()
            Button7.PerformClick()
        Catch ex As Exception

        End Try
    End Sub


    Public Function GetCertificatesToProcess(entrydate As Date, Filter As FilterStates) As List(Of ClsGiftCertificate2)
        Try
            'Get a list of certificates for a given date.

            Dim gclist1 As New List(Of ClsGiftCertificate2)
            gclist1 = RetrieveGCOrdersFromQueue(entrydate)

            Dim FilteredList As List(Of ClsGiftCertificate2)

            Select Case Filter
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

            End Select

            Return FilteredList

        Catch ex As Exception

        End Try

    End Function
    Private Sub SetupForm()
        SfDateEntry.Value = Now.Date

        'Populate fields from LIst which I want to show - Purchaser, Recipient, Shipper 
        SfDataGrid1.TableControl.VerticalScrollBarVisible = True


        SfDataGrid1.AutoGenerateColumns = False
        SfDataGrid1.AllowResizingColumns = True
        SfDataGrid1.Columns.Clear()
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "ID", .HeaderText = "Id"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Purchaser_Name", .HeaderText = "Purchaser Name"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_Status", .HeaderText = "Status"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "PointOfSale", .HeaderText = "Point Of Sale"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Online_OrderNumber", .HeaderText = "Online Order Number"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item1.Quantity", .HeaderText = "Tandem 10k"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item2.Quantity", .HeaderText = "Tandem 12k"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item3.Quantity", .HeaderText = "Tandem 10k With Vid"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item4.Quantity", .HeaderText = "Tandem 12k With Vid"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "Item5.Quantity", .HeaderText = "Video"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_TotalAmount", .HeaderText = "OrderAmount-TBD"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_TotalDiscount", .HeaderText = "DiscountAmount-TBD"})

        RdoBoth.Checked = True
        CurrentFilter = FilterStates.ProcessAndCompleteOnly
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RdoProcessed.CheckedChanged, RdoCompleted.CheckedChanged, RdoBoth.CheckedChanged
        Try

            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel1)
            Select Case s1.ToLower
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing
                Case "rdocompleted"
                    CurrentFilter = FilterStates.Completed
                Case "rdoboth"
                    CurrentFilter = FilterStates.ProcessAndCompleteOnly
            End Select


            Dim dteEntry As Date = SfDateEntry.Value
            Dim fl = GetCertificatesToProcess(dteEntry, CurrentFilter)
            SfDataGrid1.DataSource = fl

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class