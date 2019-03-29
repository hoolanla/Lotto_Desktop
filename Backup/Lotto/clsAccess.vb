Imports System.Data.SqlClient
Imports System.Configuration.ConfigurationSettings
Imports System.Data.OleDb

Public Class clsAccess


    ' Public Conn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\rep50653\preall53_kan.mdb"
    Dim db_Name As String
    Public Conn As String


    Public Sub New(ByVal dbName As String)

        Dim path As String
        path = Application.StartupPath
        db_Name = dbName

        Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & "\" & db_Name & ";Persist Security Info=False"

    End Sub

    Public Function ReturnDataTable(ByVal strSQL As String) As DataTable

        Dim dt As DataTable
        Dim cnSQL As OleDbConnection
        Dim cmSQL As OleDbCommand
        Dim drSQL As OleDbDataReader

        Try
            cnSQL = New OleDbConnection(Conn)
            cnSQL.Open()
            cmSQL = New OleDbCommand(strSQL, cnSQL)
            drSQL = cmSQL.ExecuteReader
            Dim obj As New clsDR2DT()
            dt = obj.ReaderToTable(drSQL, "tb")
            drSQL.Close()
            cnSQL.Close()
            cmSQL.Dispose()
            cnSQL.Dispose()
            Return dt
        Catch e As OleDbException
            MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
            Return New DataTable
        Catch e As Exception
            MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
            Return New DataTable
        End Try

    End Function




    Public Sub ExecuteNonQuery(ByVal strSQL As String)

        Dim cnSQL As OleDbConnection
        Dim cmSQL As OleDbCommand

        '   Try
        cnSQL = New OleDbConnection(Conn)
        cnSQL.Open()
        cmSQL = New OleDbCommand(strSQL, cnSQL)
        cmSQL.ExecuteNonQuery()

        cnSQL.Close()
        cmSQL.Dispose()
        cnSQL.Dispose()

        'Catch e As OleDbException


        '    '   Throw New Exception(e.Message)

        '    'MsgBox(e.Message, MsgBoxStyle.Critical, "oledb Error")
        'Catch e As Exception

        '    '  Throw New Exception(e.Message)
        '    ' MsgBox(e.Message, MsgBoxStyle.Critical, "General Error")
        'End Try

    End Sub


End Class
