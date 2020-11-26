Public Class FrmDevelopmentReset
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim objsql As New FrmSQL
        objsql.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DeleteImportedGiftCertificates()
        MsgBox("Completed")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ResetImportedGiftCertificatesStatus()
        MsgBox("Completed")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim cls As New ClsGiftCertificate2
            With cls
                cls.ID = 1001
                cls.JR_PurchaserID = 85315
                cls.Online_OrderNumber = "GC1"
                cls.Item1.Quantity = 2
                cls.Item1.ItemId = 1
                cls.Item2.Quantity = 1
                cls.Item2.ItemId = 2
            End With

            '//Process the line items
            GenerateIndividualCertificateRecordsFromGCOrder(cls, True)


        Catch ex As Exception

        End Try 'A test of call InvInsert to see if the record it creates are OK

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DeleteAllGCOData()
        MsgBox("Completed")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ds = GetAllGCOData()
        If ds IsNot Nothing Then
            SfDataGrid1.DataSource = ds.Tables(0)
            SfDataGrid2.DataSource = ds.Tables(1)
            SfDataGrid3.DataSource = ds.Tables(2)
            SfDataGrid4.DataSource = ds.Tables(3)
            SfDataGrid5.DataSource = ds.Tables(4)
            SfDataGrid6.DataSource = ds.Tables(5)
            SfDataGrid7.DataSource = ds.Tables(6)
            SfDataGrid8.DataSource = ds.Tables(7)
            SfDataGrid9.DataSource = ds.Tables(8)
            SfDataGrid10.DataSource = ds.Tables(9)

        End If
        MsgBox("Completed")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim x = RetrieveOperators()
        Dim sb As New System.Text.StringBuilder
        For Each i In x
            sb.AppendLine(i)
        Next
        MsgBox(sb.ToString
               )
    End Sub
End Class