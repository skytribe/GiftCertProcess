Imports System.ComponentModel
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events

Public Class FrmIncompleteItems

    Dim CurrentFilter As FilterStates
    Dim CurrentGiftCertificate As ClsGiftCertificate2

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub FrmIncompleteItems_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupForm()

        'Get Which Radio Button is clicked for status
        'Get All GCOrders for that state
        RdoEntered.Checked = True


    End Sub

    Private Sub SetupForm()
        'Populate fields from LIst which I want to show - Purchaser, Recipient, Shipper 
        SfDataGrid1.TableControl.VerticalScrollBarVisible = True

        SfDataGrid1.AutoGenerateColumns = False
        SfDataGrid1.AllowResizingColumns = True
        SfDataGrid1.Columns.Clear()
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "ID", .HeaderText = "Id"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_DateEntered", .HeaderText = "Entered Date"})
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
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "GC_TotalDiscount", .HeaderText = "DiscountAmount"})
        SfDataGrid1.Columns.Add(New GridTextColumn() With {.MappingName = "OriginalOrderDate", .HeaderText = "Original Order Date"})

        If IO.File.Exists("FrmIncomplete1.xml") Then
            Try
                Using file = System.IO.File.Open("FrmIncomplete1.xml", System.IO.FileMode.Open)
                    Me.SfDataGrid1.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub RdoEntered_CheckedChanged(sender As Object, e As EventArgs) Handles RdoEntered.CheckedChanged, RdoProcessed.CheckedChanged, RdoBoth.CheckedChanged
        Try
            '//Which was checked
            Dim s1 = WhatRadioIsSelected(Me.Panel3)
            Select Case s1.ToLower

                Case "rdoentered"
                    CurrentFilter = FilterStates.Entered
                Case "rdoprocessed"
                    CurrentFilter = FilterStates.Processing

                Case "rdoboth"
                    CurrentFilter = FilterStates.Incomplete

            End Select

            GetGCOrderForStatus()
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub GetGCOrderForStatus()
        Dim lst1 = RetrieveGCOrdersFromQueue(CurrentFilter)
        'Set the datasource property
        SfDataGrid1.DataSource = lst1
    End Sub

    Private Sub SfDataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles SfDataGrid1.SelectionChanged
        'Set Status of processing
        Try
            If CType(sender, SfDataGrid).SelectedItem IsNot Nothing Then
                CurrentGiftCertificate = CType(SfDataGrid1.SelectedItem, ClsGiftCertificate2)
                If CurrentGiftCertificate.GC_Status = CertificateStatus.Entered Then
                    Button1.Text = "Process"
                ElseIf CurrentGiftCertificate.GC_Status = CertificateStatus.Processing Then
                    Button1.Text = "Print"
                End If
                Button1.Enabled = True
            Else
                Button1.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' LogError("SfDataGrid1_SelectionChanged:", ex)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CurrentGiftCertificate.GC_Status = CertificateStatus.Entered Then
            Button1.Text = "Process"
            Dim objProcess As New FrmProcess
            objProcess.DefaultDate = CurrentGiftCertificate.GC_DateEntered
            objProcess.ShowDialog()
        ElseIf CurrentGiftCertificate.GC_Status = CertificateStatus.Processing Then
            Button1.Text = "Print"
            Dim objProcess As New FrmProcessPrint
            objProcess.Certificate = CurrentGiftCertificate
            objProcess.ShowDialog()
        End If
    End Sub

    Private Sub FrmIncompleteItems_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Using file = System.IO.File.Create("FrmIncomplete1.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using
    End Sub
End Class