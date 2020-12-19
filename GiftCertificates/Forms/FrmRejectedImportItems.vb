Imports System.ComponentModel

Public Class FrmRejectedImportItems
    Public Property TotalAttemptedImports As Integer
    Public Property ValidImports As Integer
    Public Property DuplicateImports As Integer
    Public Property RejectedImports As List(Of ClsWooLineItem)

    Sub New(total As Integer, Valid As Integer, duplicate As Integer, rejected As List(Of ClsWooLineItem))

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        TotalAttemptedImports = total
        ValidImports = Valid
        DuplicateImports = duplicate
        RejectedImports = rejected

    End Sub
    Private Sub FrmRejectedImportItems_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblValid.Text = ValidImports.ToString
        LblDuplicate.Text = DuplicateImports.ToString
        SfDataGrid1.DataSource = RejectedImports
        lblTotal.Text = TotalAttemptedImports.ToString
        lblrejected.Text = RejectedImports.Count.ToString

        If IO.File.Exists("FrmRejected.xml") Then
            Try
                Using file = System.IO.File.Open("FrmIncomplete1.xml", System.IO.FileMode.Open)
                    Me.SfDataGrid1.Deserialize(file)
                End Using
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub FrmRejectedImportItems_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Using file = System.IO.File.Create("FrmProcess2.xml")
            Me.SfDataGrid1.Serialize(file)
        End Using
    End Sub
End Class