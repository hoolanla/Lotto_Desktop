Public Class clsDR2DT
    Inherits System.Data.Common.DbDataAdapter

    Public Function ReaderToTable(ByVal dr As IDataReader, ByVal tableName As String) As DataTable

        Dim dt As New DataTable(tableName)

        Me.Fill(dt, dr)

        dr.Close()
        Return dt

    End Function

    Public Function ReaderToTable(ByRef dt As DataTable, ByVal dr As IDataReader) As Integer

        Return Me.Fill(dt, dr)

    End Function

    Protected Overrides Function CreateRowUpdatedEvent(ByVal dr As DataRow, ByVal idc As IDbCommand, ByVal st As StatementType, ByVal dtm As System.Data.Common.DataTableMapping) As System.Data.Common.RowUpdatedEventArgs

        Return DirectCast(New EventArgs(), System.Data.Common.RowUpdatedEventArgs)

    End Function

    Protected Overrides Function CreateRowUpdatingEvent(ByVal dr As DataRow, ByVal idc As IDbCommand, ByVal st As StatementType, ByVal dtm As System.Data.Common.DataTableMapping) As System.Data.Common.RowUpdatingEventArgs
        Return DirectCast(New EventArgs(), System.Data.Common.RowUpdatingEventArgs)

    End Function

    Protected Overrides Sub OnRowUpdated(ByVal e As System.Data.Common.RowUpdatedEventArgs)

    End Sub

    Protected Overrides Sub OnRowUpdating(ByVal e As System.Data.Common.RowUpdatingEventArgs)

    End Sub
End Class
