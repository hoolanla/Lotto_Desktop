Imports System.data.SqlClient

Public Class clsManageSQL

    Private conn As SqlConnection
    Private comm As SqlCommand
    Private tr As SqlTransaction
    '  Private strcon As String = "Data Source=localhost;User ID=sa;PWD=12345679;initial catalog=asset"


    ' Private strcon As String = "Server=.\SQLExpress;AttachDbFilename=E:\work\asset\bak_source\data\asset.mdf;Database=asset; Trusted_Connection=Yes;"
    Private strcon As String = "Data Source=" & My.Settings.DBHOST & ";User ID=" & My.Settings.USERNAME & ";PWD=" & My.Settings.password & ";initial catalog=lotto"

    Public Function Open() As Boolean

        Try
            conn = New SqlConnection
            comm = New SqlCommand
            With conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = strcon
                .Open()
                Return True
            End With
        Catch ex As Exception
            Throw New Exception("{clsManageSQL.open}" & ex.Message & " " & Now)
            Return False
        End Try

    End Function

    Public Sub Dispose()

        conn.Close()
        conn.Dispose()
        comm.Dispose()

    End Sub

    Public Function ExecuteDataHaving(ByVal sql As String) As Boolean
        Try
            Dim rd As SqlDataReader
            comm.CommandType = CommandType.Text
            comm.CommandText = sql
            comm.Connection = conn
            rd = comm.ExecuteReader()

            If rd.Read Then
                rd.Close()
                Return True
            Else
                rd.Close()
                Return False
            End If

        Catch ex As Exception
            Throw New Exception("ExecuteDataHaving" & ex.Message)
        Finally
        End Try
    End Function

    Public Function ExecuteNonQuery(ByVal sql As String) As Boolean

        Try
            Dim i As Integer
            comm.Connection = conn
            tr = conn.BeginTransaction
            comm.Transaction = tr
            comm.CommandType = CommandType.Text
            comm.CommandText = sql
            i = comm.ExecuteNonQuery()
            tr.Commit()
            If i > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw New Exception("clsManageSql Error: {ExecuteNonQuery}-" & ex.Message)
            tr.Rollback()
            Return False
        End Try

    End Function

    Public Function ExecuteDataTable(ByVal sql As String) As DataTable

        Dim da As SqlDataAdapter
        Dim ds As New DataSet
        conn = New SqlConnection
        comm = New SqlCommand
        With conn
            If .State = ConnectionState.Open Then .Close()
            .ConnectionString = strcon
            .Open()
        End With
        da = New SqlDataAdapter(sql, strcon)
        da.Fill(ds, "table")
        Return ds.Tables(0).Copy()

    End Function

    Public Function getMaxID(ByVal tableName As String, ByVal fdIdxName As String) As Integer

        Try
            Dim sql As String
            Open()

            Dim rd As SqlDataReader
            sql = "select max(" & fdIdxName & ")as maxID from " & tableName
            comm.CommandType = CommandType.Text
            comm.CommandText = sql
            comm.Connection = conn
            rd = comm.ExecuteReader()

            If rd.Read Then
                Return CInt(rd("maxID")) + 1
            Else
                Return 1
            End If


        Catch ex As Exception
            Return 1
        Finally
            Dispose()
        End Try

    End Function

End Class

