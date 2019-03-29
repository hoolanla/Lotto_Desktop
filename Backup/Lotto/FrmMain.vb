
Imports System.io
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.data

Public Class FrmMain
    Dim connGetBarcode As New clsManageSQL
    Dim dtMain, dtGetBarcode As DataTable
    Dim MyConnection As System.Data.OleDb.OleDbConnection
    Dim DtSet As System.Data.DataSet
    Dim dt As DataTable


    Dim max_, max_2under, max_teng_, max_tod_ As String

#Region "Common Function EXCEL"
    Private Sub Load_Excel_Details()



        Dim flagGet As Boolean = True
        flagGet = Me.chkGetorNot()







        'Extracting from database
        Dim str, filename As String
        Dim col, row As Integer
        Dim conn As New clsAccess("lotto.mdb")


        Dim dt As DataTable
        Try
            Dim sql As String
            'sql = "select no_ as เลข,cutting as ราคา  from temp where cutting > 0")




            If flagGet = True Then
                sql = "select no_ as เลข,cutting as บน  from temp "
            Else
                sql = "select no_ as เลข,price as บน  from temp "
            End If

            dt = conn.ReturnDataTable(sql)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MsgBox("It appears that Excel is not installed on this machine. This operation requires MS Excel to be installed on this machine.", MsgBoxStyle.Critical)
            Return
        End If


        'Export to Excel process
        Try
            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()


                Dim i As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    .cells(1, i).value = dt.Columns(col).ColumnName
                    .cells(1, i).EntireRow.Font.Bold = True
                    .cells(1, i).style.NumberFormat = "00"
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    i = 2
                    For row = 0 To dt.Rows.Count - 1
                        .Cells(i, k).Value = dt.Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next

                If flagGet = True Then
                    filename = "c:\เก็บได้_2ตัวบน_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                Else
                    filename = "c:\เก็บไม่ได้_2ตัวบน_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                End If
                .ActiveCell.Worksheet.SaveAs(filename)
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox("Data's are exported to Excel Succesfully in '" & filename & "'", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' The excel is created and opened for insert value. We most close this excel using this system
        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub Load_Excel_Details2under()

        Dim flagGet As Boolean
        flagGet = chkGetorNot2Under()



        'Extracting from database
        Dim str, filename As String
        Dim col, row As Integer
        Dim conn As New clsAccess("lotto.mdb")


        Dim dt As DataTable
        Try
            Dim sql As String
            'sql = "select no_,cutting from temp2under where cutting > 0"
            If flagGet = True Then
                sql = "select no_ as เลข,cutting  as ล่าง from temp2under "
            Else
                sql = "select no_ as เลข,price  as ล่าง from temp2under "
            End If

            dt = conn.ReturnDataTable(sql)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MsgBox("It appears that Excel is not installed on this machine. This operation requires MS Excel to be installed on this machine.", MsgBoxStyle.Critical)
            Return
        End If


        'Export to Excel process
        Try
            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    .cells(1, i).value = dt.Columns(col).ColumnName
                    .cells(1, i).EntireRow.Font.Bold = True
                    .cells(1, i).style.NumberFormat = "00"
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    i = 2
                    For row = 0 To dt.Rows.Count - 1
                        .Cells(i, k).Value = dt.Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next
                If flagGet = True Then
                    filename = "c:\เก็บได้_2ตัวล่าง_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"

                Else
                    filename = "c:\เก็บไม่ได้_2ตัวล่าง_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"

                End If
                .ActiveCell.Worksheet.SaveAs(filename)
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox("Data's are exported to Excel Succesfully in '" & filename & "'", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' The excel is created and opened for insert value. We most close this excel using this system
        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub



    Private Sub Load_Excel_Details3()
        'Extracting from database
        Dim str, filename As String
        Dim col, row As Integer
        Dim conn As New clsAccess("lotto.mdb")


        Dim dt As DataTable
        Try
            Dim sql As String
            ' sql = "select no_,cutting from temp_x where cutting > 0"
            sql = "select no_ as เลข,cutting as ราคา  from temp_x "


            dt = conn.ReturnDataTable(sql)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MsgBox("It appears that Excel is not installed on this machine. This operation requires MS Excel to be installed on this machine.", MsgBoxStyle.Critical)
            Return
        End If


        'Export to Excel process
        Try
            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    .cells(1, i).value = dt.Columns(col).ColumnName
                    .cells(1, i).EntireRow.Font.Bold = True
                    .cells(1, i).style.NumberFormat = "00"
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    i = 2
                    For row = 0 To dt.Rows.Count - 1
                        .Cells(i, k).Value = dt.Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next
                filename = "c:\3ตัวเต็ง_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                .ActiveCell.Worksheet.SaveAs(filename)
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox("Data's are exported to Excel Succesfully in '" & filename & "'", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' The excel is created and opened for insert value. We most close this excel using this system
        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub Load_Excel_Details3_Tod()
        'Extracting from database
        Dim str, filename As String
        Dim col, row As Integer
        Dim conn As New clsAccess("lotto.mdb")


        Dim dt As DataTable
        Try
            Dim sql As String
            ' sql = "select no_,cutting from temp_x where cutting > 0"
            sql = "select no_ as เลข,cutting as ราคา  from temp_tod_120 "


            dt = conn.ReturnDataTable(sql)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MsgBox("It appears that Excel is not installed on this machine. This operation requires MS Excel to be installed on this machine.", MsgBoxStyle.Critical)
            Return
        End If


        'Export to Excel process
        Try
            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    .cells(1, i).value = dt.Columns(col).ColumnName
                    .cells(1, i).EntireRow.Font.Bold = True
                    .cells(1, i).style.NumberFormat = "00"
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To dt.Columns.Count - 1
                    i = 2
                    For row = 0 To dt.Rows.Count - 1
                        .Cells(i, k).Value = dt.Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next
                filename = "c:\3ตัวโต๊ด_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                .ActiveCell.Worksheet.SaveAs(filename)
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox("Data's are exported to Excel Succesfully in '" & filename & "'", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' The excel is created and opened for insert value. We most close this excel using this system
        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub


#End Region


    Private Function chkGetorNot() As Boolean



        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select avg(cutting) as cut_ from [temp]"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows(0)("cut_").ToString = "0" Then
            Return False
        Else
            Return True
        End If



    End Function

    Private Function chkGetorNot2Under() As Boolean



        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select avg(cutting) as cut_ from [temp2under]"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows(0)("cut_").ToString = "0" Then
            Return False
        Else
            Return True
        End If



    End Function



    Private Sub insertFilename(ByVal filename)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "insert into tb_file_name([file_name]) values('" & filename & "')"

        conn.ExecuteNonQuery(sql)

    End Sub

    Private Sub bindNotPrice()

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select no_,price  from temp where price = 0"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        Me.dgNotPrice.DataSource = dt


    End Sub


    Private Sub bindNotPrice2Under()

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select no_,price  from temp2Under where price = 0"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        Me.dgNotPrice2Under.DataSource = dt


    End Sub



    Private Sub insertFilename2under(ByVal filename)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "insert into tb_file_name_2under([file_name]) values('" & filename & "')"

        conn.ExecuteNonQuery(sql)

    End Sub



    Private Sub insertFilename3(ByVal filename)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "insert into tb_filename3([file_name]) values('" & filename & "')"

        conn.ExecuteNonQuery(sql)

    End Sub


    Private Sub insertFilename3_tod(ByVal filename)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "insert into tb_filename3_tod([file_name]) values('" & filename & "')"

        conn.ExecuteNonQuery(sql)

    End Sub





    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If Me.tbPath_asset.Text = "" Then
            MsgBox("โปรดเลือกไฟล์ EXCEL")
            Exit Sub
        End If


        Dim arr() As String
        arr = Me.tbPath_asset.Text.Split("\")
        Dim tmp As String
        tmp = arr(UBound(arr))


        If chkFilename(tmp) Then
            MsgBox("มีไฟล์นี้ในระบบแล้วโปรดตรวจสอบ")


            bindFilename()

            Exit Sub
        Else
            insertFilename(tmp)

            bindFilename()
        End If





        selectSheet(Me.tbPath_asset.Text)



        Me.ImportFile(Me.tbPath_asset.Text, "Sheet1$")



        bnt_import_Click(Me.btn_import, e)



        If Not dt Is Nothing Then
            dt.Dispose()
        End If

        bindNotPrice()
        bindNotPrice2Under()
        lblCountON.Text = lblCountON_()
        lblCountUnder.Text = lblCountUnder_()


    End Sub


    Private Sub bindFilename()



        Dim sql2 As String
        Dim dt1 As DataTable
        sql2 = "select * from tb_file_name"
        Dim conn As New clsAccess("lotto.mdb")
        dt1 = conn.ReturnDataTable(sql2)
        dg_filename.DataSource = dt1


    End Sub


 


    Private Sub bindFilename3()



        Dim sql2 As String
        Dim dt1 As DataTable
        sql2 = "select * from tb_filename3"
        Dim conn As New clsAccess("lotto.mdb")
        dt1 = conn.ReturnDataTable(sql2)
        Me.DGFilename3.DataSource = dt1


    End Sub



    Private Sub bindFilename3_tod()



        Dim sql2 As String
        Dim dt1 As DataTable
        sql2 = "select * from tb_filename3_tod"
        Dim conn As New clsAccess("lotto.mdb")
        dt1 = conn.ReturnDataTable(sql2)
        Me.DGFilename3_tod.DataSource = dt1


    End Sub




    Private Function chkFilename(ByVal file_name As String) As Boolean

        Dim sql As String
        sql = "select * from tb_file_name where file_name= '" & file_name & "'"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("file_name").ToString
            If tmp = file_name Then
                Return True
            Else
                Return False
            End If


        End If
        Return False
    End Function


    Private Function chkFilename2under(ByVal file_name As String) As Boolean

        Dim sql As String
        sql = "select * from tb_file_name_2under where file_name= '" & file_name & "'"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("file_name").ToString
            If tmp = file_name Then
                Return True
            Else
                Return False
            End If


        End If
        Return False
    End Function


    Private Function chkFilename3(ByVal file_name As String) As Boolean

        Dim sql As String
        sql = "select * from tb_filename3 where file_name= '" & file_name & "'"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("file_name").ToString
            If tmp = file_name Then
                Return True
            Else
                Return False
            End If


        End If
        Return False
    End Function



    Private Function chkFilename3_tod(ByVal file_name As String) As Boolean

        Dim sql As String
        sql = "select * from tb_filename3_tod where file_name= '" & file_name & "'"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("file_name").ToString
            If tmp = file_name Then
                Return True
            Else
                Return False
            End If


        End If
        Return False
    End Function





    Private Sub selectSheet(ByVal path As String)
        Try


            Dim dt As DataTable = Nothing
            Dim con As System.Data.OleDb.OleDbConnection
            con = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & path & " '; " & "Extended Properties=Excel 8.0;")
            con.Open()
            dt = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)


            cbSheet.Items.Clear()

            Dim i As Integer
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count

                    Me.cbSheet.Items.Add(dt.Rows(0)("TABLE_NAME"))

                Next
            End If

            '   Me.ddlSelectSheet.Items.Insert(0, "-- Select Sheet --")
        Catch ex As Exception
            MsgBox("เลือกไฟล์  EXCEL เท่านั้น")
        End Try
    End Sub


    Private Sub selectSheet2(ByVal path As String)
        Try


            Dim dt As DataTable = Nothing
            Dim con As System.Data.OleDb.OleDbConnection
            con = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & path & " '; " & "Extended Properties=Excel 8.0;")
            con.Open()
            dt = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)


            cbSheet2.Items.Clear()

            Dim i As Integer
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count

                    Me.cbSheet2.Items.Add(dt.Rows(0)("TABLE_NAME"))

                Next
            End If

            '   Me.ddlSelectSheet.Items.Insert(0, "-- Select Sheet --")
        Catch ex As Exception
            MsgBox("เลือกไฟล์  EXCEL เท่านั้น")
        End Try
    End Sub


    Private Sub ImportFile(ByVal Filename As String, ByVal SheetName As String)
        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & Filename & " '; " & "Extended Properties=Excel 8.0;")

            'Select the data from Sheet1 of the workbook.


            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & SheetName & "]", MyConnection)




            MyCommand.TableMappings.Add("Table", "AccData")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)

            dtMain = DtSet.Tables("AccData")
            Dg1.DataSource = dtMain 'DtSet.Tables(0)
            MyConnection.Close()

            'If Me.radAsset.Checked Then
            '    If Dg1.ColumnCount <= 2 Then
            '        MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            '        Dim dt As New DataTable
            '        Dg1.DataSource = dt
            '    End If
            'End If

        Catch ex As Exception
            MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            Me.Dispose()
            Me.Close()
        End Try

    End Sub


    Private Sub ImportFile2under(ByVal Filename As String, ByVal SheetName As String)
        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & Filename & " '; " & "Extended Properties=Excel 8.0;")

            'Select the data from Sheet1 of the workbook.


            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & SheetName & "]", MyConnection)




            MyCommand.TableMappings.Add("Table", "AccData")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)

            dtMain = DtSet.Tables("AccData")
            DGDT2under.DataSource = dtMain 'DtSet.Tables(0)
            MyConnection.Close()

            'If Me.radAsset.Checked Then
            '    If Dg1.ColumnCount <= 2 Then
            '        MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            '        Dim dt As New DataTable
            '        Dg1.DataSource = dt
            '    End If
            'End If

        Catch ex As Exception
            MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            Me.Dispose()
            Me.Close()
        End Try

    End Sub


    Private Sub ImportFile3(ByVal Filename As String, ByVal SheetName As String)
        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & Filename & " '; " & "Extended Properties=Excel 8.0;")

            'Select the data from Sheet1 of the workbook.


            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & SheetName & "]", MyConnection)




            MyCommand.TableMappings.Add("Table", "AccData")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)

            dtMain = DtSet.Tables("AccData")
            DGDT3.DataSource = dtMain 'DtSet.Tables(0)
            MyConnection.Close()

            'If Me.radAsset.Checked Then
            '    If Dg1.ColumnCount <= 2 Then
            '        MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            '        Dim dt As New DataTable
            '        Dg1.DataSource = dt
            '    End If
            'End If

        Catch ex As Exception
            MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            Me.Dispose()
            Me.Close()
        End Try
    End Sub



    Private Sub ImportFile3_tod(ByVal Filename As String, ByVal SheetName As String)
        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                           "data source='" & Filename & " '; " & "Extended Properties=Excel 8.0;")

            'Select the data from Sheet1 of the workbook.


            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & SheetName & "]", MyConnection)




            MyCommand.TableMappings.Add("Table", "AccData")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)

            dtMain = DtSet.Tables("AccData")
            DGDT3_tod.DataSource = dtMain 'DtSet.Tables(0)
            MyConnection.Close()


        Catch ex As Exception
            MsgBox("ไฟล์ EXCEL อาจผิดฟอร์แมท หรือ อาจใส่ชื่อ Sheet ผิด")
            Me.Dispose()
            Me.Close()
        End Try
    End Sub





    Private Sub cbSheet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSheet.SelectedIndexChanged

        'Me.ImportFile(Me.tbPath_asset.Text, cbSheet.SelectedItem)



        'bnt_import_Click(Me.btn_import, e)



        'If Not dt Is Nothing Then
        '    dt.Dispose()
        'End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        Dim reg_ As New ClsReg

        If reg_.GetSetting("ApplicationCount") = "" Then
            reg_.SaveSetting("ApplicationCount", "0")
        End If


        If CInt(reg_.GetSetting("ApplicationCount")) >= 10 Then
            MsgBox("หมดอายุแล้วครับ!")
            Application.Exit()
        Else

            Dim iCount As Byte
            iCount = CInt(reg_.GetSetting("ApplicationCount")) + 1
            reg_.SaveSetting("ApplicationCount", CStr(iCount))
        End If





        lblFloorOn.Text = "ต้องแทงอย่างน้อย :" & My.Settings.two & " ตัว"
        lblFloorUnder.Text = "ต้องแทงอย่างน้อย :" & My.Settings.twoUnder & " ตัว"


        Me.bindDGMain()
        Me.bindFilename()



        bindNotPrice()
        bindNotPrice2Under()


    End Sub

    Private Sub btPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPath.Click
        OpenFileDialog1.InitialDirectory = "c:\"
        ' OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.Filter = "xls files (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 2
        Me.OpenFileDialog1.ShowDialog()

        Me.tbPath_asset.Text = OpenFileDialog1.FileName
    End Sub


    Private Sub insertDB_TMP()

        Try

            Dim i As Integer
            Dim No_, price_on, price_under As String


            If Dg1.RowCount = 0 Then
                MsgBox("เกิดข้อผิดพลาด")
                Exit Sub
            End If

            Dim sqlAll As String = ""

            'Dim conn As New clsManageSQL
            'conn.Open()

            Dim conn As New clsAccess("lotto.mdb")

        

            For i = 0 To Dg1.RowCount - 2
                Dim sql, sql2 As String

                If IsDBNull(Dg1.Item(1, i).Value) = True Then
                    Exit Sub
                End If

                If IsDBNull(Dg1.Item(0, i).Value) = False Then
                    No_ = Dg1.Item(0, i).Value
                End If

                '  price on 2 ตัวบน
                If IsDBNull(Dg1.Item(1, i).Value) = False Then
                    price_on = Dg1.Item(1, i).Value
                End If



                ' under 2 ตัวล่าง
                If IsDBNull(Dg1.Item(2, i).Value) = False Then
                    price_under = Dg1.Item(2, i).Value
                End If




                ' sql = "insert into  temp (price,no_) values(" & price & ",'" & No_ & "')"
                sql = " update  temp set price = price + " & price_on & " where no_ = '" & No_ & "'" & vbCrLf
                sql2 = " update  temp2under set price = price + " & price_under & " where no_ = '" & No_ & "'"


                conn.ExecuteNonQuery(sql)
                conn.ExecuteNonQuery(sql2)
            Next
            '    MsgBox("Upload Complete")
            '' conn.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub insertDB_TMP2under()

        Try


            Dim i As Integer
            Dim No_, price, price_teng, price_tod As String



            If DGDT2under.RowCount = 0 Then
                MsgBox("เกิดข้อผิดพลาด")
                Exit Sub
            End If

            Dim sqlAll As String = ""
            'Dim conn As New clsManageSQL
            'conn.Open()
            Dim conn As New clsAccess("lotto.mdb")

            Dim flag_teng As Boolean = False
            If DGDT2under.RowCount > 800 Then
                flag_teng = True
            End If

            For i = 0 To DGDT2under.RowCount - 2

                Dim sql As String = ""

                If IsDBNull(DGDT2under.Item(1, i).Value) = True Then
                    Exit Sub
                End If

                If IsDBNull(DGDT2under.Item(0, i).Value) = False Then
                    No_ = DGDT2under.Item(0, i).Value
                End If

                '  price or price teng
                If IsDBNull(DGDT2under.Item(2, i).Value) = False Then
                    price = DGDT2under.Item(2, i).Value
                End If


                If flag_teng = True Then
                    '  price or price tod
                    If IsDBNull(DGDT2under.Item(2, i).Value) = False Then
                        price_tod = DGDT2under.Item(2, i).Value
                    End If
                    '   sql = "insert into  temp_x (no_,price_teng,price_tod) values('" & No_ & "'," & price & "," & price_tod & ")"

                    sql = " update  temp set price_teng=price + " & price & ","
                    sql += " price_tod=price + " & price_tod
                    sql += " where no_='" & No_ & "'"
                Else
                    ' sql = "insert into  temp (price,no_) values(" & price & ",'" & No_ & "')"
                    sql = "update  temp2under set price=price + " & price & " where no_='" & No_ & "'"

                End If



                conn.ExecuteNonQuery(sql)
            Next
            '    MsgBox("Upload Complete")
            '' conn.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Function factor6(ByVal value As String)

        Dim p1, p2, p3 As String
        Dim r1, r2, r3, r4, r5, r6 As String

        p1 = Mid(value, 1, 1)
        p2 = Mid(value, 2, 1)
        p3 = Mid(value, 3, 1)

        r1 = p1 & p2 & p3
        r2 = p1 & p3 & p2
        r3 = p2 & p1 & p3
        r4 = p2 & p3 & p1
        r5 = p3 & p1 & p2
        r6 = p3 & p2 & p1

        Return "('" & r1 & "','" & r2 & "','" & r3 & "','" & r4 & "','" & r5 & "','" & r6 & "')"


    End Function

    Private Sub insertDB_TMP3_tod()

        Try

            Dim i As Integer
            Dim No_, price As String

            If DGDT3_tod.RowCount = 0 Then
                MsgBox("เกิดข้อผิดพลาด")
                Exit Sub
            End If

            Dim sqlAll As String = ""
            Dim conn As New clsAccess("lotto.mdb")

            Dim flag_teng As Boolean = False
            If DGDT3_tod.RowCount > 800 Then
                flag_teng = True
            End If

            For i = 0 To DGDT3_tod.RowCount - 2

                Dim sql As String = ""

                If IsDBNull(DGDT3_tod.Item(1, i).Value) = True Then
                    Exit Sub
                End If

                If IsDBNull(DGDT3_tod.Item(0, i).Value) = False Then
                    No_ = DGDT3_tod.Item(0, i).Value
                End If

                '  price or price teng
                If IsDBNull(DGDT3_tod.Item(2, i).Value) = False Then
                    price = DGDT3_tod.Item(2, i).Value
                End If

                Dim strWhere As String
                strWhere = factor6(No_)

                '   sql = "insert into  temp_x (no_,price_teng,price_tod) values('" & No_ & "'," & price & "," & price_tod & ")"
                sql = " update  temp_tod_120 set price=price + " & price
                sql += " where no_ in " & strWhere

                conn.ExecuteNonQuery(sql)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub insertDB_TMP3()

        Try
            Dim i As Integer
            Dim No_, price As String

            If DGDT3.RowCount = 0 Then
                MsgBox("เกิดข้อผิดพลาด")
                Exit Sub
            End If

            Dim sqlAll As String = ""
            Dim conn As New clsAccess("lotto.mdb")

            Dim flag_teng As Boolean = False
            If DGDT3.RowCount > 800 Then
                flag_teng = True
            End If

            For i = 0 To DGDT3.RowCount - 2

                Dim sql As String = ""

                If IsDBNull(DGDT3.Item(1, i).Value) = True Then
                    Exit Sub
                End If

                If IsDBNull(DGDT3.Item(0, i).Value) = False Then
                    No_ = DGDT3.Item(0, i).Value
                End If

                '  price or price teng
                If IsDBNull(DGDT3.Item(1, i).Value) = False Then
                    price = DGDT3.Item(1, i).Value
                End If

                'sql = "insert into  temp_tod_120 (no_,price) values('" & No_ & "'," & price & ")"

                sql = " update  temp_x set price=price + " & price
                sql += " where no_ = '" & No_ & "'"

                conn.ExecuteNonQuery(sql)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub bnt_import_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_import.Click
        insertDB_TMP()
        Me.bindDGMain()
        MsgBox("Complete")

    End Sub


    Private Function getMax() As String
        Dim m As String
        m = max()

        Dim sql1 As String
        sql1 = "select count(*) as c from temp where price = " & m
        Dim conn1 As New clsAccess("lotto.mdb")
        Dim dt1 As DataTable
        dt1 = conn1.ReturnDataTable(sql1)
        If dt1.Rows.Count > 0 Then
            Dim t As String
            t = dt1.Rows(0)("c").ToString
            If t = "100" Then
                selectAll(m)
                Return "end"
            End If
        End If

        Dim sql As String
        sql = "select price from temp where price <= " & m & " order by price desc"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("price").ToString()
            Dim chkCount As String
            chkCount = chkCountRatePrice(tmp)

            If Int(chkCount) >= Int(My.Settings.two) Then
                Return tmp
            Else
                Return subMax(tmp)
                '     Return "False"
            End If
        End If

    End Function


    Private Function getMax2under() As String
        Dim m As String
        m = max2under()

        Dim sql1 As String
        sql1 = "select count(*) as c from temp2under where price = " & m
        Dim conn1 As New clsAccess("lotto.mdb")
        Dim dt1 As DataTable
        dt1 = conn1.ReturnDataTable(sql1)
        If dt1.Rows.Count > 0 Then
            Dim t As String
            t = dt1.Rows(0)("c").ToString
            If t = "100" Then

                selectAll2under(m)
                Return "end"
            End If
        End If

        Dim sql As String
        sql = "select price from temp2under where price <= " & m & " order by price desc"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("price").ToString()
            Dim chkCount As String
            chkCount = chkCountRatePrice2under(tmp)

            If Int(chkCount) >= Int(My.Settings.twoUnder) Then
                Return tmp
            Else

                Return subMax2under(tmp)
                '     Return "False"
            End If
        End If

    End Function

    Private Function getMax3() As String
        Dim m As String
        m = max3()

        Dim sql1 As String
        sql1 = "select count(*) as c from temp_x where price = " & m
        Dim conn1 As New clsAccess("lotto.mdb")
        Dim dt1 As DataTable
        dt1 = conn1.ReturnDataTable(sql1)
        If dt1.Rows.Count > 0 Then
            Dim t As String
            t = dt1.Rows(0)("c").ToString
            If t = "1000" Then


                selectAll3(m)
                Return "end"
            End If
        End If

        Dim sql As String
        sql = "select price from temp_x where price <= " & m & " order by price desc"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("price").ToString()
            Dim chkCount As String
            chkCount = chkCountRatePrice3(tmp)

            If Int(chkCount) >= Int(My.Settings.tree500) Then
                Return tmp
            Else

                Return subMax3(tmp)
                '     Return "False"
            End If
        End If

    End Function


    Private Function getMax3_tod() As String
        Dim m As String
        m = max3_tod()

        Dim sql1 As String
        sql1 = "select count(*) as c from temp_tod_120 where price = " & m
        Dim conn1 As New clsAccess("lotto.mdb")
        Dim dt1 As DataTable
        dt1 = conn1.ReturnDataTable(sql1)
        If dt1.Rows.Count > 0 Then
            Dim t As String
            t = dt1.Rows(0)("c").ToString
            If t = "120" Then


                selectAll3_tod(m)
                Return "end"
            End If
        End If

        Dim sql As String
        sql = "select price from temp_tod_120 where price <= " & m & " order by price desc"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim tmp As String
            tmp = dt.Rows(0)("price").ToString()
            Dim chkCount As String
            chkCount = chkCountRatePrice3_tod(tmp)

            If Int(chkCount) >= Int(My.Settings.tree100) Then
                Return tmp
            Else

                Return subMax3_tod(tmp)
                '     Return "False"
            End If
        End If

    End Function


    Private Function subMax(ByVal tmp As String) As String
        Dim sql As String
        Dim conn As New clsAccess("lotto.mdb")
        sql = "select price from temp where price < " & tmp & " order by price desc"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("price").ToString
        End If
    End Function




    Private Function subMax2under(ByVal tmp As String) As String
        Dim sql As String
        Dim conn As New clsAccess("lotto.mdb")
        sql = "select price from temp2under where price < " & tmp & " order by price desc"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("price").ToString
        End If
    End Function



    Private Function subMax3(ByVal tmp As String) As String
        Dim sql As String
        Dim conn As New clsAccess("lotto.mdb")
        sql = "select price from temp_x where price < " & tmp & " order by price desc"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("price").ToString
        End If
    End Function


    Private Function subMax3_tod(ByVal tmp As String) As String
        Dim sql As String
        Dim conn As New clsAccess("lotto.mdb")
        sql = "select price from temp_tod_120 where price < " & tmp & " order by price desc"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("price").ToString
        End If
    End Function


    Private Function chkCountRatePrice(ByVal price As String) As String

        Dim sql As String
        sql = "select count(*) as tot from temp where price >= " & price
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)("tot").ToString()
        End If


    End Function

    Private Function chkCountRatePrice2under(ByVal price As String) As String

        Dim sql As String
        sql = "select count(*) as tot from temp2under where price >= " & price
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)("tot").ToString()
        End If


    End Function


    Private Function chkCountRatePrice3(ByVal price As String) As String

        Dim sql As String
        sql = "select count(*) as tot from temp_x where price >= " & price
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)("tot").ToString()
        End If


    End Function


    Private Function chkCountRatePrice3_tod(ByVal price As String) As String

        Dim sql As String
        sql = "select count(*) as tot from temp_tod_120 where price >= " & price
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)("tot").ToString()
        End If


    End Function


    Private Function chkMin(ByVal max_ As String) As Boolean

        Dim sql As String
        sql = "select min(price) as m from temp "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Dim tmp As String
            tmp = dt.Rows(0)("m").ToString
            If max_ = tmp Then
                Return True
            Else
                Return False
            End If

        End If


    End Function


    Private Function chkMin2under(ByVal max_ As String) As Boolean

        Dim sql As String
        sql = "select min(price) as m from temp2under "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Dim tmp As String
            tmp = dt.Rows(0)("m").ToString
            If max_2under = tmp Then
                Return True
            Else
                Return False
            End If

        End If


    End Function

    Private Function chkMin3(ByVal max_ As String) As Boolean

        Dim sql As String
        sql = "select min(price) as m from temp_x "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Dim tmp As String
            tmp = dt.Rows(0)("m").ToString
            If max_teng_ = tmp Then
                Return True
            Else
                Return False
            End If

        End If


    End Function


    Private Function chkMin3_tod(ByVal max_ As String) As Boolean

        Dim sql As String
        sql = "select min(price) as m from temp_tod_120 "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Dim tmp As String
            tmp = dt.Rows(0)("m").ToString
            If max_tod_ = tmp Then
                Return True
            Else
                Return False
            End If

        End If


    End Function



    Private Function chkFirst() As Boolean

        Dim sql As String
        sql = "select count(*) as c from temp where price > 0"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As Integer
            tmp = Int(dt.Rows(0)("c").ToString)
            If tmp < Int(My.Settings.two) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function


    Private Function chkFirst2Under() As Boolean

        Dim sql As String
        sql = "select count(*) as c from temp2under where price > 0"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As Integer
            tmp = Int(dt.Rows(0)("c").ToString)
            If tmp < Int(My.Settings.two) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function


    Private Function chkFirst3() As Boolean

        Dim sql As String
        sql = "select count(*) as c from temp_x where price > 0"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As Integer
            tmp = Int(dt.Rows(0)("c").ToString)
            If tmp < Int(My.Settings.tree500) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function


    Private Function chkFirst3_tod() As Boolean

        Dim sql As String
        sql = "select count(*) as c from temp_tod_120 where price > 0"
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)

        If dt.Rows.Count > 0 Then
            Dim tmp As Integer
            tmp = Int(dt.Rows(0)("c").ToString)

            tmp = Int(tmp) * 6

            If tmp < Int(My.Settings.tree100) Then
                Return True
            Else
                Return False
            End If
        End If

    End Function





    Private Function upperFirstZero() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select price from temp where price > 0 order by price asc"
        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        Return dt.Rows(0)("price").ToString





    End Function




    Private Sub process_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles process.Click

        lblCountON.Text = lblCountON_()


            If chkFirst() Then
                MsgBox("ไม่สามารถเก็บได้")
                Exit Sub
            End If


        max_ = getMax()



        ' Edit Percentage ##############

        max_ = upperFirstZero()


        '###############################






            If chkMin(max_) Then

                Dim sumLessPay As String
                sumLessPay = sumdefult(max_)
            Dim countThan As String
                countThan = countThanDefult(max_)
            bee(max_, countThan, sumLessPay)
                Exit Sub
            End If

        If max_ = "end" Then
            Exit Sub
        End If

            If max_ = "False" Then
                MsgBox("NO ")
                Exit Sub
            End If


            Dim show As String
            Dim i As Integer
            Dim arrList1 As New ArrayList()

            For i = 0 To 10000
                show = q(max_)


            If show <= 100 - Int(My.Settings.two) Then

                Dim sumLessPay As String
                sumLessPay = sumdefult(max_)

                Dim countThan As String
                countThan = countThanDefult(max_)

                bee(max_, countThan, sumLessPay)


                Exit For
            Else
                max_ = max_ - 10
            End If
            Next

    End Sub









    Private Function bee(ByVal max_ As String, ByVal count_ As String, ByVal sumLessPay As String) As String

        Dim arrList As New ArrayList


        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select * from temp where price >= " & max_ & " order by price "
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim i As Integer
            Dim dupValue As String = ""

            For i = 0 To dt.Rows.Count - 1
                Dim tmp As Integer

                Dim tmpRight As Integer


                Try

                    tmp = (sumLessPay + retSumZigma(i + 1, dt)) + (Int(dt.Rows(i + 1)("price") * (count_ - 1)))
                    tmp = tmp - (tmp * (Int(My.Settings.percent2on) / 100))

                Catch ex As Exception



                    'Edit percentage ##############################

                    '  MsgBox("Not get")

                    selectAll(Int(dt.Rows(i)("price")))
                    '##############################################

                    Exit Function
                End Try

                ' Edit percentage ###################################

                tmpRight = Int(My.Settings.two) * Int(dt.Rows(i + 1)("price"))
                'tmpRight = Int(My.Settings.two) * Int(dt.Rows(i + 1)("price"))
                'tmpRight = tmpRight - (tmpRight * (Int(My.Settings.percent2on) / 100))
                '##################################################

                arrList.Add(dt.Rows(i + 1)("price"))



                'โกยกำไรให้เหลือน้อยสุด

                If tmp <= tmpRight Then
                    'Pumb Profit 

                    Dim show As String

                    If i = 0 Then
                        show = arrList(arrList.Count - 1)

                    Else
                        show = arrList(arrList.Count - 2)
                    End If


                    ' Edit percentage #######

                    If i > 0 Then
                        selectAll(show)
                    End If
                    '#######################3

                    Dim newTmp As Integer
                    newTmp = tmp - (Int(dt.Rows(i + 1)("price") * (count_ - 1)))


                    Dim no_max As String
                    sql = "select no_ from temp where [get] =" & show
                    Dim conn3 As New clsAccess("lotto.mdb")
                    Dim dt2 As DataTable
                    dt2 = conn.ReturnDataTable(sql)
                    Dim ddd As String
                    ddd = newTmp '+ profit_arm(dt2.Rows(0)("no_").ToString())
                    '#################    CORE   ###########################
                    'Edit percenatge ##############################3
                    If ddd < 0 Then
                        MsgBox("Not get")
                        Exit For
                    End If
                    '###############################################

                    If core(ddd, count_ - 1) Then
                        MsgBox("OK")

                    End If

                    '########################################################


                    Return show
                End If
                count_ = count_ - 1

            Next

        End If


    End Function


    Private Function bee2under(ByVal max_ As String, ByVal count_ As String, ByVal sumLessPay As String) As String

        Dim arrList As New ArrayList


        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select * from temp2under where price >= " & max_ & " order by price "
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim i As Integer
            Dim dupValue As String = ""

            For i = 0 To dt.Rows.Count - 1
                Dim tmp As Integer

                Dim tmpRight As Integer


                Try

                    tmp = (sumLessPay + retSumZigma2under(i + 1, dt)) + (Int(dt.Rows(i + 1)("price") * (count_ - 1)))
                    tmp = tmp - (tmp * (Int(My.Settings.percent2on) / 100))

                Catch ex As Exception

                    selectAll2under(Int(dt.Rows(i)("price")))


                    Exit Function
                End Try





                tmpRight = Int(My.Settings.twoUnder) * Int(dt.Rows(i + 1)("price"))

                arrList.Add(dt.Rows(i + 1)("price"))





                If tmp <= tmpRight Then
                    'Pumb Profit 

                    Dim show As String

                    If i = 0 Then
                        show = arrList(arrList.Count - 1)

                    Else
                        show = arrList(arrList.Count - 2)
                    End If

                    selectAll2under(show)


                    Dim newTmp As Integer
                    newTmp = tmp - (Int(dt.Rows(i + 1)("price") * (count_ - 1)))


                    Dim no_max As String
                    sql = "select no_ from temp2under where [get] =" & show
                    Dim conn3 As New clsAccess("lotto.mdb")
                    Dim dt2 As DataTable
                    dt2 = conn.ReturnDataTable(sql)
                    Dim ddd As String
                    ddd = newTmp '+ profit_arm(dt2.Rows(0)("no_").ToString())
                    '#################    CORE   ###########################

                    If core2under(ddd, count_ - 1) Then
                        MsgBox("OK")

                    End If

                    '########################################################


                    Return show
                End If
                count_ = count_ - 1

            Next

        End If


    End Function


    Private Function bee3(ByVal max_ As String, ByVal count_ As String, ByVal sumLessPay As String) As String

        Dim arrList As New ArrayList


        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select * from temp_x where price >= " & max_ & " order by price "
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim i As Integer
            Dim dupValue As String = ""

            For i = 0 To dt.Rows.Count - 1
                Dim tmp As Integer

                Dim tmpRight As Integer


                Try

                    tmp = (sumLessPay + retSumZigma3(i + 1, dt)) + (Int(dt.Rows(i + 1)("price") * (count_ - 1)))
                    tmp = tmp - (tmp * (Int(My.Settings.percent2on) / 100))

                Catch ex As Exception

                    selectAll3(Int(dt.Rows(i)("price")))


                    Exit Function
                End Try
                tmpRight = Int(My.Settings.tree500) * Int(dt.Rows(i + 1)("price"))

                arrList.Add(dt.Rows(i + 1)("price"))





                If tmp <= tmpRight Then
                    'Pumb Profit 

                    Dim show As String

                    If i = 0 Then
                        show = arrList(arrList.Count - 1)

                    Else
                        show = arrList(arrList.Count - 2)
                    End If

                    selectAll3(show)


                    Dim newTmp As Integer
                    newTmp = tmp - (Int(dt.Rows(i + 1)("price") * (count_ - 1)))


                    Dim no_max As String
                    sql = "select no_ from temp_x where [get] =" & show
                    Dim conn3 As New clsAccess("lotto.mdb")
                    Dim dt2 As DataTable
                    dt2 = conn.ReturnDataTable(sql)
                    Dim ddd As String
                    ddd = newTmp '+ profit_arm(dt2.Rows(0)("no_").ToString())
                    '#################    CORE   ###########################

                    If core3(ddd, count_ - 1) Then
                        MsgBox("OK")

                    End If

                    '########################################################


                    Return show
                End If
                count_ = count_ - 1

            Next

        End If


    End Function



    Private Function bee3_tod(ByVal max_ As String, ByVal count_ As String, ByVal sumLessPay As String) As String

        Dim arrList As New ArrayList


        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select * from temp_tod_120 where price >= " & max_ & " order by price "
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim i As Integer
            Dim dupValue As String = ""

            For i = 0 To dt.Rows.Count - 1
                Dim tmp As Integer

                Dim tmpRight As Integer


                Try

                    tmp = (sumLessPay + retSumZigma3_tod(i + 1, dt)) + (Int(dt.Rows(i + 1)("price") * (count_ - 1)))
                    tmp = tmp - (tmp * (Int(My.Settings.percent2on) / 100))

                Catch ex As Exception

                    selectAll3_tod(Int(dt.Rows(i)("price")))


                    Exit Function
                End Try
                tmpRight = Int(My.Settings.tree100) * Int(dt.Rows(i + 1)("price"))

                arrList.Add(dt.Rows(i + 1)("price"))





                If tmp <= tmpRight Then
                    'Pumb Profit 

                    Dim show As String

                    If i = 0 Then
                        show = arrList(arrList.Count - 1)

                    Else
                        show = arrList(arrList.Count - 2)
                    End If

                    selectAll3_tod(show)


                    Dim newTmp As Integer
                    newTmp = tmp - (Int(dt.Rows(i + 1)("price") * (count_ - 1)))


                    Dim no_max As String
                    sql = "select no_ from temp_tod_120 where [get] =" & show
                    Dim conn3 As New clsAccess("lotto.mdb")
                    Dim dt2 As DataTable
                    dt2 = conn.ReturnDataTable(sql)
                    Dim ddd As String
                    ddd = newTmp '+ profit_arm(dt2.Rows(0)("no_").ToString())
                    '#################    CORE   ###########################

                    If core2under(ddd, count_ - 1) Then
                        MsgBox("OK")

                    End If

                    '########################################################


                    Return show
                End If
                count_ = count_ - 1

            Next

        End If


    End Function




    Private Function core(ByVal ddd As Integer, ByVal count_ As Integer) As Boolean


        Dim get_max As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        Dim sql As String
        sql = "select max([get]) as m from temp "
        dt = conn.ReturnDataTable(sql)
        get_max = dt.Rows(0)("m").ToString()


        Dim profit As String
        sql = "select no_ from temp where [get] =" & get_max
        Dim conn3 As New clsAccess("lotto.mdb")
        Dim dt2 As DataTable
        dt2 = conn.ReturnDataTable(sql)


        ' Edit percentage ##################
        profit = profit_arm(dt2.Rows(0)("no_").ToString())

        ' Edit2

        If profit < 0 Then
            '  selectAll(0)
            Return True
        End If


        '###################################

        ' เอากำไร 90  กัน loop
        If Int(profit) < 100 Then
            '   MsgBox(Int(profit))

            Return True
        End If

        Dim tmp1, tmp2 As Integer
        tmp1 = ddd + ((get_max + (profit / (100 - count_))) * count_)
        tmp1 = tmp1 - (tmp1 * (Int(My.Settings.percent2on) / 100))

        ' EDIT Percentage    ###############
        tmp2 = Int(My.Settings.two) * (get_max + (profit / (100 - count_)))
        'tmp2 = Int(My.Settings.two) * (get_max + (profit / (100 - count_)))
        'tmp2 = tmp2 - (tmp2 * (Int(My.Settings.percent2on) / 100))



        '##################################

        If tmp1 < tmp2 Then

            Return True
        Else

            selectAll(Int(get_max + (profit / (100 - count_))))

            core(ddd, count_)
        End If


    End Function


    Private Function core2under(ByVal ddd As Integer, ByVal count_ As Integer) As Boolean

        Dim get_max As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        Dim sql As String
        sql = "select max([get]) as m from temp2under "
        dt = conn.ReturnDataTable(sql)
        get_max = dt.Rows(0)("m").ToString()


        Dim profit As String
        sql = "select no_ from temp2under where [get] =" & get_max
        Dim conn3 As New clsAccess("lotto.mdb")
        Dim dt2 As DataTable
        dt2 = conn.ReturnDataTable(sql)



        ' Edit percentage ##################
        profit = profit_arm2under(dt2.Rows(0)("no_").ToString())

        ' Edit2

        If profit < 0 Then
            '  selectAll(0)
            Return True
        End If


        '###################################

        ' profit = profit_arm2under(dt2.Rows(0)("no_").ToString())





        If Int(profit) < 100 Then
            Return True
        End If

        Dim tmp1, tmp2 As Integer
        tmp1 = ddd + ((get_max + (profit / (100 - count_))) * count_)
        tmp1 = tmp1 - (tmp1 * (Int(My.Settings.percent2on) / 100))
        ' tmp2 = Int(My.Settings.twoUnder) * (get_max + (profit / (100 - count_)))



        ' EDIT Percentage    ###############
        tmp2 = Int(My.Settings.twoUnder) * (get_max + (profit / (100 - count_)))
        'tmp2 = Int(My.Settings.two) * (get_max + (profit / (100 - count_)))
        'tmp2 = tmp2 - (tmp2 * (Int(My.Settings.percent2on) / 100))



        '##################################



        If tmp1 < tmp2 Then

            Return True
        Else

            selectAll2under(Int(get_max + (profit / (100 - count_))))

            core2under(ddd, count_)
        End If


    End Function

    Private Function core3(ByVal ddd As Integer, ByVal count_ As Integer) As Boolean

        Dim get_max As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        Dim sql As String
        sql = "select max([get]) as m from temp_x "
        dt = conn.ReturnDataTable(sql)
        get_max = dt.Rows(0)("m").ToString()


        Dim profit As String
        sql = "select no_ from temp_x where [get] =" & get_max
        Dim conn3 As New clsAccess("lotto.mdb")
        Dim dt2 As DataTable
        dt2 = conn.ReturnDataTable(sql)
        profit = profit_arm3(dt2.Rows(0)("no_").ToString())

        If Int(profit) < 100 Then
            Return True
        End If

        Dim tmp1, tmp2 As Integer
        tmp1 = ddd + ((get_max + (profit / (100 - count_))) * count_)
        tmp2 = Int(My.Settings.tree500) * (get_max + (profit / (100 - count_)))

        If tmp1 < tmp2 Then

            Return True
        Else

            selectAll3(Int(get_max + (profit / (100 - count_))))

            core3(ddd, count_)
        End If


    End Function


    Private Function core3_tod(ByVal ddd As Integer, ByVal count_ As Integer) As Boolean

        Dim get_max As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable
        Dim sql As String
        sql = "select max([get]) as m from temp_tod_120 "
        dt = conn.ReturnDataTable(sql)
        get_max = dt.Rows(0)("m").ToString()


        Dim profit As String
        sql = "select no_ from temp_tod_120 where [get] =" & get_max
        Dim conn3 As New clsAccess("lotto.mdb")
        Dim dt2 As DataTable
        dt2 = conn.ReturnDataTable(sql)
        profit = profit_arm3(dt2.Rows(0)("no_").ToString())

        If Int(profit) < 100 Then
            Return True
        End If

        Dim tmp1, tmp2 As Integer
        tmp1 = ddd + ((get_max + (profit / (100 - count_))) * count_)
        tmp2 = Int(My.Settings.tree500) * (get_max + (profit / (100 - count_)))

        If tmp1 < tmp2 Then

            Return True
        Else

            selectAll3(Int(get_max + (profit / (100 - count_))))

            core3(ddd, count_)
        End If


    End Function



    Private Function beeTeng(ByVal max_ As String, ByVal count_ As String, ByVal sumLessPay As String) As String

        Dim arrList As New ArrayList


        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "select * from temp_x where price_teng > " & max_ & " order by price_teng "
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                Dim tmp As Integer

                Dim tmpRight As Integer

                tmp = (sumLessPay + retSumZigma(i + 1, dt)) + (Int(dt.Rows(i + 1)("price_teng") * (count_ - 1)))


                tmpRight = Int(My.Settings.tree500) * Int(dt.Rows(i + 1)("price_teng"))

                arrList.Add(dt.Rows(i + 1)("price_teng"))

                If tmp <= tmpRight Then


                    Dim show As String
                    show = arrList(arrList.Count - 2)
                    Return show
                End If
                count_ = count_ - 1

            Next

        End If


    End Function



    Private Function retSumZigma(ByVal loop_ As Integer, ByVal dt As DataTable) As Integer

        Dim tmp As Integer = 0
        For i As Integer = 1 To loop_
            tmp = tmp + dt.Rows(i - 1)("price")

        Next

        Return tmp


    End Function



    Private Function retSumZigma2under(ByVal loop_ As Integer, ByVal dt As DataTable) As Integer

        Dim tmp As Integer = 0
        For i As Integer = 1 To loop_
            tmp = tmp + dt.Rows(i - 1)("price")

        Next


        Return tmp


    End Function

    Private Function retSumZigma3(ByVal loop_ As Integer, ByVal dt As DataTable) As Integer

        Dim tmp As Integer = 0
        For i As Integer = 1 To loop_
            tmp = tmp + dt.Rows(i - 1)("price")

        Next


        Return tmp


    End Function


    Private Function retSumZigma3_tod(ByVal loop_ As Integer, ByVal dt As DataTable) As Integer
        Dim tmp As Integer = 0
        For i As Integer = 1 To loop_
            tmp = tmp + dt.Rows(i - 1)("price")
        Next
        Return tmp
    End Function


    Private Function sumdefult(ByVal max_ As Integer) As String

        Dim sql As String
        sql = "select sum(price) from temp where price <= " & max_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function sumdefult2under(ByVal max_ As Integer) As String

        Dim sql As String
        sql = "select sum(price) from temp2under where price <= " & max_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function sumdefult3(ByVal max_ As Integer) As String

        Dim sql As String
        sql = "select sum(price) from temp_x where price <= " & max_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function






    Private Function sumdefult3_tod(ByVal max_ As Integer) As String

        Dim sql As String
        sql = "select sum(price) from temp_tod_120 where price <= " & max_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function





    Private Function sumdefultTeng(ByVal max_ As Integer) As String

        Dim sql As String
        sql = "select sum(price_teng) from temp_x where price_teng < " & max_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function countThanDefult(ByVal max_1) As String

        Dim sql As String
        sql = "select count(price) as tot from temp where price >= " & max_1
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function countThanDefult2under(ByVal max_1) As String

        Dim sql As String
        sql = "select count(price) as tot from temp2under where price >= " & max_1
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function countThanDefult3(ByVal max_1) As String

        Dim sql As String
        sql = "select count(price) as tot from temp_x where price >= " & max_1
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Function countThanDefult3_tod(ByVal max_1) As String

        Dim sql As String
        sql = "select count(price) as tot from temp_tod_120 where price >= " & max_1
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function



    Private Function countThanDefultTeng() As String

        Dim sql As String
        sql = "select count(price_teng) as tot from temp_x where price_teng >= " & max_teng_
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Return dt.Rows(0)(0).ToString
        End If


    End Function


    Private Sub selectAll(ByVal max_arm As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting = IIf([price]>" & max_arm & ",[price] - " & max_arm & ",0)," & vbCrLf
        sql += " [get] = IIF([price] < " & max_arm & ",[price]," & max_arm & ")"

        conn.ExecuteNonQuery(sql)


        Dim dt As DataTable



        sql = "select * from temp"
        dt = conn.ReturnDataTable(sql)
        dgTest.DataSource = dt

        takeLotto()
        cutLotto()
        price2onAll()


    End Sub


    Private Sub selectAll2under(ByVal max_arm As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp2under"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting = IIf([price]>" & max_arm & ",[price] - " & max_arm & ",0)," & vbCrLf
        sql += " [get] = IIF([price] < " & max_arm & ",[price]," & max_arm & ")"

        conn.ExecuteNonQuery(sql)


        Dim dt As DataTable



        sql = "select * from temp2under"
        dt = conn.ReturnDataTable(sql)
        Me.DGMain_2under.DataSource = dt

        takeLotto2under()
        cutLotto2under()
        price2UnderAll()


    End Sub


    Private Sub selectAll3(ByVal max_arm As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp_x"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting = IIf([price]>" & max_arm & ",[price] - " & max_arm & ",0)," & vbCrLf
        sql += " [get] = IIF([price] < " & max_arm & ",[price]," & max_arm & ")"

        conn.ExecuteNonQuery(sql)


        Dim dt As DataTable



        sql = "select * from temp_x"
        dt = conn.ReturnDataTable(sql)
        Me.dgMain3.DataSource = dt

        takeLotto3()
        cutLotto3()
        price3All()


    End Sub

    Private Sub selectAll3_tod(ByVal max_arm As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp_tod_120"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting = IIf([price]>" & max_arm & ",[price] - " & max_arm & ",0)," & vbCrLf
        sql += " [get] = IIF([price] < " & max_arm & ",[price]," & max_arm & ")"

        conn.ExecuteNonQuery(sql)


        Dim dt As DataTable



        sql = "select * from temp_tod_120"
        dt = conn.ReturnDataTable(sql)
        Me.dgMain3_tod.DataSource = dt

        takeLotto3_tod()
        cutLotto3_tod()
        price3All_tod()


    End Sub



    Private Sub selectAll2(ByVal max_teng_ As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp_x"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting_teng = IIf([price_teng]>" & max_teng_ & ",[price_teng] - " & max_teng_ & ",0)," & vbCrLf
        sql += " [get_teng] = IIF([price_teng] < " & max_teng_ & ",[price_teng]," & max_teng_ & ")"

        conn.ExecuteNonQuery(sql)

        Dim dt As DataTable

        sql = "select * from temp_x"
        dt = conn.ReturnDataTable(sql)
        dgMain3.DataSource = dt

        takeLotto3()
        cutLotto3()


    End Sub

    Private Sub selectAllTeng(ByVal max_teng_ As Integer)

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        sql = "UPDATE temp_x"
        '  sql += " SET cutting = CASE WHEN (price > " & max_ & ") THEN price - " & max_ & " END,"
        ' sql += " [get] = CASE WHEN (price < " & max_ & ") THEN price WHEN (price >= " & max_ & ") THEN " & max_ & " END"

        sql += " SET cutting_teng = IIf([price_teng]>= " & max_teng_ & ",[price_teng] - " & max_teng_ & ",[price_teng])," & vbCrLf
        sql += " [get_teng] = IIF([price_teng] <= " & max_teng_ & ",[price_teng]," & max_teng_ & ")"
        conn.ExecuteNonQuery(sql)

        Dim dt As DataTable



        sql = "select no_,price_teng,get_teng,cutting_teng from temp_x"
        dt = conn.ReturnDataTable(sql)
        dgMain3.DataSource = dt

        takeLotto3()
        cutLotto3()


    End Sub


    Private Sub takeLotto()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([get]) from temp"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            txtTake.Text = dt.Rows(0)(0).ToString

        End If

    End Sub

    Private Sub takeLotto2under()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([get]) from temp2under"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtTake2under.Text = dt.Rows(0)(0).ToString

        End If

    End Sub

    Private Sub price2onAll()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([price]) from temp"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtPrice2onAll.Text = dt.Rows(0)(0).ToString

        End If

    End Sub

    Private Sub price2UnderAll()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([price]) from temp2under"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtPrice2underAll.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub price3All()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([price]) from temp_x"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtPrice3All.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub price3All_tod()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([price]) from temp_tod_120"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtPrice3todAll.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub takeLotto3()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([get]) from temp_x"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            txtTake3.Text = dt.Rows(0)(0).ToString

        End If

    End Sub

    Private Sub takeLotto3_tod()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([get]) from temp_tod_120"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            txtTake3_tod.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub cutLotto()
        Dim sql As String
        Dim dt As DataTable
        sql = "select sum(cutting) from temp"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            txtCut.Text = dt.Rows(0)(0).ToString

        End If


    End Sub

    Private Sub cutLotto2under()
        Dim sql As String
        Dim dt As DataTable
        sql = "select sum(cutting) from temp2under"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtCutting2under.Text = dt.Rows(0)(0).ToString

        End If


    End Sub

    Private Sub cutLotto2()
        Dim sql As String
        Dim dt As DataTable
        sql = "select sum(cutting) from temp_x"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtCutting3.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub cutLotto3()
        Dim sql As String
        Dim dt As DataTable
        sql = "select sum(cutting) from temp_x"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtCutting3.Text = dt.Rows(0)(0).ToString
        End If
    End Sub

    Private Sub cutLotto3_tod()
        Dim sql As String
        Dim dt As DataTable
        sql = "select sum(cutting) from temp_tod_120"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtCutting3_tod.Text = dt.Rows(0)(0).ToString
        End If
    End Sub

    Private Sub takeLottoTeng()

        Dim sql As String
        Dim dt As DataTable
        sql = "select sum([get]) from temp_x"
        Dim conn As New clsAccess("lotto.mdb")

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Me.txtTake3.Text = dt.Rows(0)(0).ToString

        End If

    End Sub


    Private Sub test()

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select no_,count(*) as count_ from temp where price < " & max() & " group by no_"

        dt = conn.ReturnDataTable(sql)
        dgTest.DataSource = dt

    End Sub


    Private Function q(ByVal max As Integer) As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select count(*) as count_ from temp where price < " & max

        dt = conn.ReturnDataTable(sql)
        Dim result As String
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)("count_")
            Return result
        End If
    End Function


    Private Function q2under(ByVal max As Integer) As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select count(*) as count_ from temp2under where price < " & max

        dt = conn.ReturnDataTable(sql)
        Dim result As String
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)("count_")
            Return result
        End If
    End Function


    Private Function q3(ByVal max As Integer) As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select count(*) as count_ from temp_x where price < " & max

        dt = conn.ReturnDataTable(sql)
        Dim result As String
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)("count_")
            Return result
        End If
    End Function

    Private Function q3_tod(ByVal max As Integer) As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select count(*) as count_ from temp_tod_120 where price < " & max

        dt = conn.ReturnDataTable(sql)
        Dim result As String
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)("count_")
            Return result
        End If
    End Function


    Private Function q2(ByVal max As Integer) As String
        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        sql = "select count(*) as count_ from temp_x where price_teng < " & max

        dt = conn.ReturnDataTable(sql)
        Dim result As String
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)("count_")
            Return result
        End If
    End Function

    Private Function max() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        Dim result As String
        sql = "select avg(price) from temp"

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)(0)
            Return result
        End If
        Return "0"

    End Function

    Private Function max2under() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        Dim result As String
        sql = "select avg(price) from temp2under"

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)(0)
            Return result
        End If
        Return "0"

    End Function


    Private Function max3() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        Dim result As String
        sql = "select avg(price) from temp_x"

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)(0)
            Return result
        End If
        Return "0"

    End Function


    Private Function max3_tod() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        Dim result As String
        sql = "select avg(price) from temp_tod_120"

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)(0)
            Return result
        End If
        Return "0"

    End Function



    Private Function max_teng() As String

        Dim conn As New clsAccess("lotto.mdb")
        Dim sql As String
        Dim dt As DataTable
        Dim result As String
        sql = "select max(price_teng) from temp_x"

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            result = dt.Rows(0)(0)
            Return result
        End If
        Return "0"

    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Load_Excel_Details()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        two_()

    End Sub

    Private Function profit_arm(ByVal no_ As String)
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.two) & " as total, "
        sql += " [get] * " & Int(My.Settings.two) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.two) & " as cutting "
        sql += " from temp where no_ = '" & no_ & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Profit2(dt.Rows(0)("we").ToString())
        End If
    End Function




    Private Function profit_arm2under(ByVal no_ As String)
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.twoUnder) & " as total, "
        sql += " [get] * " & Int(My.Settings.twoUnder) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.twoUnder) & " as cutting "
        sql += " from temp2under where no_ = '" & no_ & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Profit2_2under(dt.Rows(0)("we").ToString())
        End If
    End Function

    Private Function profit_arm3(ByVal no_ As String)
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.tree500) & " as total, "
        sql += " [get] * " & Int(My.Settings.tree500) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.tree500) & " as cutting "
        sql += " from temp_x where no_ = '" & no_ & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Profit2_3(dt.Rows(0)("we").ToString())
        End If
    End Function

    Private Function profit_arm3_tod(ByVal no_ As String)
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.tree100) & " as total, "
        sql += " [get] * " & Int(My.Settings.tree100) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.tree100) & " as cutting "
        sql += " from temp_tod_120 where no_ = '" & no_ & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Profit2_3_tod(dt.Rows(0)("we").ToString())
        End If
    End Function


    Private Sub two_()
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.two) & " as total, "
        sql += " [get] * " & Int(My.Settings.two) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.two) & " as cutting "
        sql += " from temp where no_ = '" & two.Text & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Me.txt2total.Text = dt.Rows(0)("total").ToString()
            Me.txt2We.Text = dt.Rows(0)("we").ToString()

            Me.txt2Dealer.Text = dt.Rows(0)("cutting").ToString()

            Me.txtProfit.Text = Profit2(dt.Rows(0)("we").ToString())

        End If

    End Sub

    Private Sub two_under()
        Dim sql As String
        sql = "select [price]  *  " & Int(My.Settings.two) & " as total, "
        sql += " [get] * " & Int(My.Settings.two) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.two) & " as cutting "
        sql += " from temp2under where no_ = '" & twoUnder.Text & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Me.txt2Undertotal.Text = dt.Rows(0)("total").ToString()
            Me.txt2UnderWe.Text = dt.Rows(0)("we").ToString()

            Me.txt2UnderDealer.Text = dt.Rows(0)("cutting").ToString()

            Me.txtProfit2Under.Text = Profit2_2under(dt.Rows(0)("we").ToString())

        End If

    End Sub


    Private Function Profit2(ByVal pay As String) As String

        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        Dim sql As String
        sql = "select sum([get]) from temp "
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Int(dt.Rows(0)(0).ToString) - Int(pay)

        End If

    End Function


    Private Function Profit2_2under(ByVal pay As String) As String

        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        Dim sql As String
        sql = "select sum([get]) from temp2under "
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Int(dt.Rows(0)(0).ToString) - Int(pay)

        End If

    End Function

    Private Function Profit2_3(ByVal pay As String) As String

        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        Dim sql As String
        sql = "select sum([get]) from temp_x "
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Int(dt.Rows(0)(0).ToString) - Int(pay)

        End If

    End Function


    Private Function Profit2_3_tod(ByVal pay As String) As String

        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        Dim sql As String
        sql = "select sum([get]) from temp_tod_120 "
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return Int(dt.Rows(0)(0).ToString) - Int(pay)

        End If

    End Function


    Private Sub Pay_Three()
        Dim sql As String
        sql = "select [price] * " & Int(My.Settings.tree500) & " as total, "
        sql += " [get] * " & Int(My.Settings.tree500) & " as we, "
        sql += " [cutting] * " & Int(My.Settings.tree500) & " as cutting "
        sql += " from temp_x where no_ = '" & three.Text & "'"
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Me.txtPayTengTotal.Text = dt.Rows(0)("total").ToString()
            Me.txtPayTengWe.Text = dt.Rows(0)("we").ToString()
            Me.txtPayTengDealer.Text = dt.Rows(0)("cutting").ToString()



            Me.txtProfit3.Text = Profit2_3(dt.Rows(0)("we").ToString())
        End If
    End Sub


    Private Sub Pay_Tod(ByVal strWhere As String)
        Dim sql As String
        sql = "select [price] * " & (Int(My.Settings.tree100) / 6) & " as total, "
        sql += " [get] * " & (Int(My.Settings.tree100) / 6) & " as we, "
        sql += " [cutting] * " & (Int(My.Settings.tree100) / 6) & " as cutting "
        sql += " from temp_tod_120 where no_ in " & strWhere
        Dim conn As New clsAccess("lotto.mdb")

        Dim dt As DataTable
        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then

            Me.txtPayTodTotal.Text = dt.Rows(0)("total").ToString()
            Me.txtPayTodWe.Text = dt.Rows(0)("we").ToString()
            Me.txtPayTodDealer.Text = dt.Rows(0)("cutting").ToString()



            Me.txtProfit3_tod.Text = Profit2_3_tod(dt.Rows(0)("we").ToString())
        End If
    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub btnProcessTeng_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess3.Click

        If chkFirst3() Then
            MsgBox("ไม่สามารถเก็บได้")
            Exit Sub
        End If


        max_teng_ = getMax3()


        If chkMin3(max_teng_) Then

            Dim sumLessPay As String
            sumLessPay = sumdefult3(max_teng_)
            Dim countThan As String
            countThan = countThanDefult3(max_teng_)
            bee3(max_teng_, countThan, sumLessPay)
            Exit Sub
        End If

        If max_teng_ = "end" Then
            Exit Sub
        End If

        If max_teng_ = "False" Then
            MsgBox("NO ")
            Exit Sub
        End If


        Dim show As String
        Dim i As Integer
        Dim arrList1 As New ArrayList()

        For i = 0 To 10000
            show = q3(max_teng_)


            If show <= 1000 - Int(My.Settings.tree500) Then

                Dim sumLessPay As String
                sumLessPay = sumdefult3(max_teng_)

                Dim countThan As String
                countThan = countThanDefult3(max_teng_)

                bee3(max_teng_, countThan, sumLessPay)


                Exit For
            Else
                max_teng_ = max_teng_ - 10
            End If
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim conn As New clsAccess("lotto.mdb")

        Dim sql As String
        sql = "update temp_x set price_teng = 0"
        conn.ExecuteNonQuery(sql)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Pay_Three()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        If MessageBox.Show("ต้องการล้างข้อมูล 2 ตัวบนและล่าง เพื่อคำนวณใหม่ทั้งหมด", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then


            Dim conn As New clsAccess("lotto.mdb")

            Dim sql, sql2 As String
            sql = "update temp set price = 0,cutting=0,[get] = 0"
            sql2 = "update temp2under set price = 0,cutting=0,[get] = 0"

            conn.ExecuteNonQuery(sql)
            conn.ExecuteNonQuery(sql2)

            Dim conn1 As New clsAccess("lotto.mdb")

            Dim sql1 As String
            sql1 = "delete from tb_file_name "
            conn1.ExecuteNonQuery(sql1)


            resetForm()
            resetForm2under()

        End If

        bindNotPrice()
        bindNotPrice2Under()

    End Sub


    Private Sub resetForm()

        bindDGMain()
        bindFilename()
        Me.txtPrice2onAll.Text = ""
        txtCut.Text = ""
        txtTake.Text = ""

    End Sub


    Private Sub resetForm2under()

        bindDGMain2under()
        Me.txtPrice2underAll.Text = ""
        Me.txtCutting2under.Text = ""
        Me.txtTake2under.Text = ""

    End Sub

    Private Sub resetForm3()

        bindDGMain3()
        bindFilename3()
        Me.txtPrice3All.Text = ""
        Me.txtCutting3.Text = ""
        Me.txtTake3.Text = ""
    End Sub


    Private Sub resetForm3_tod()
        bindDGMain3_tod()
        bindFilename3_tod()
        Me.txtPrice3todAll.Text = ""
        Me.txtCutting3_tod.Text = ""
        Me.txtTake3_tod.Text = ""
    End Sub

    Private Sub bindDGMain()

        Dim sql As String
        sql = "select * from temp"
        Dim dt As DataTable
        Dim conn As New clsAccess("lotto.mdb")
        dt = conn.ReturnDataTable(sql)
        Me.dgTest.DataSource = dt

    End Sub


    Private Sub bindDGMain2under()

        Dim sql As String
        sql = "select * from temp2under"
        Dim dt As DataTable
        Dim conn As New clsAccess("lotto.mdb")
        dt = conn.ReturnDataTable(sql)
        Me.DGMain_2under.DataSource = dt

    End Sub

    Private Sub bindDGMain3()

        Dim sql As String
        sql = "select * from temp_tod_120"
        Dim dt As DataTable
        Dim conn As New clsAccess("lotto.mdb")
        dt = conn.ReturnDataTable(sql)
        Me.dgMain3_tod.DataSource = dt

    End Sub


    Private Sub bindDGMain3_tod()

        Dim sql As String
        sql = "select * from temp_tod_120"
        Dim dt As DataTable
        Dim conn As New clsAccess("lotto.mdb")
        dt = conn.ReturnDataTable(sql)
        Me.dgMain3_tod.DataSource = dt

    End Sub


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        OpenFileDialog2.InitialDirectory = "c:\"
        ' OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog2.RestoreDirectory = True
        OpenFileDialog2.Filter = "xls files (*.xls)|*.xls"
        OpenFileDialog2.FilterIndex = 2
        Me.OpenFileDialog2.ShowDialog()
        Me.tbPath3.Text = OpenFileDialog2.FileName

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnimportfile3.Click

        If Me.tbPath3.Text = "" Then
            MsgBox("โปรดเลือกไฟล์ EXCEL")
            Exit Sub
        End If

        Dim arr() As String
        arr = Me.tbPath3.Text.Split("\")
        Dim tmp As String
        tmp = arr(UBound(arr))

        If chkFilename3(tmp) Then
            MsgBox("มีไฟล์นี้ในระบบแล้วโปรดตรวจสอบ")

            bindFilename3()

            Exit Sub
        Else
            insertFilename3(tmp)

            bindFilename3()
        End If

        Me.ImportFile3(Me.tbPath3.Text, "Sheet1$")
        btn_imp3_Click(Me.btn_imp3, e)

        If Not dt Is Nothing Then
            dt.Dispose()
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSheet2.SelectedIndexChanged

        Me.ImportFile3(Me.tbPath3.Text, cbSheet2.SelectedItem)

        If Not dt Is Nothing Then
            dt.Dispose()
        End If
    End Sub

    Private Sub btn_imp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_imp3.Click

        insertDB_TMP3()
        Me.bindDGMain3()
        MsgBox("Complete")

    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetForm3.Click

        If MessageBox.Show("ต้องการล้างข้อมูล 3 ตัวเต็ง เพื่อคำนวณใหม่ทั้งหมด", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then


            Dim conn As New clsAccess("lotto.mdb")

            Dim sql As String
            sql = "update temp_x set price = 0,cutting=0,[get] = 0" & vbCrLf

            conn.ExecuteNonQuery(sql)


            Dim conn1 As New clsAccess("lotto.mdb")

            Dim sql1 As String
            sql1 = "delete from tb_filename3 "
            conn1.ExecuteNonQuery(sql1)

            resetForm3()
        End If
    End Sub


    Private Sub txtReset2under_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReset2under.Click

        If MessageBox.Show("ต้องการล้างข้อมูล 2 ตัวล่างเพื่อคำนวณใหม่ทั้งหมด", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

            Dim conn As New clsAccess("lotto.mdb")

            Dim sql, sql2 As String

            sql = "update temp2under set price = 0,cutting=0,[get] = 0"
            sql2 = "update temp set price = 0,cutting=0,[get] = 0"



            conn.ExecuteNonQuery(sql)
            conn.ExecuteNonQuery(sql2)


            resetForm()
            resetForm2under()

        End If
        bindNotPrice()
        bindNotPrice2Under()
    End Sub

    Private Sub btnCal2under_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCal2under.Click


        lblCountUnder.Text = lblCountUnder_()


        If chkFirst2Under() Then
            MsgBox("ไม่สามารถเก็บได้")
            Exit Sub
        End If

        max_2under = getMax2under()

        If chkMin2under(max_2under) Then

            Dim sumLessPay As String
            sumLessPay = sumdefult2under(max_2under)
            Dim countThan As String
            countThan = countThanDefult2under(max_2under)
            bee2under(max_2under, countThan, sumLessPay)
            Exit Sub
        End If

        If max_2under = "end" Then
            Exit Sub
        End If

        If max_2under = "False" Then
            MsgBox("NO ")
            Exit Sub
        End If


        Dim show As String
        Dim i As Integer
        Dim arrList1 As New ArrayList()

        For i = 0 To 10000
            show = q2under(max_2under)

            If show <= 100 - Int(My.Settings.twoUnder) Then

                Dim sumLessPay As String
                sumLessPay = sumdefult2under(max_2under)
                Dim countThan As String
                countThan = countThanDefult2under(max_2under)
                bee2under(max_2under, countThan, sumLessPay)

                Exit For
            Else
                max_2under = max_2under - 10
            End If
        Next


    End Sub

    Private Sub btn_importfile_2under_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_importfile_2under.Click

        insertDB_TMP2under()

        Me.bindDGMain2under()
        MsgBox("Complete")
    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Load_Excel_Details3()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        Dim i As Integer
        i = TabControl1.SelectedIndex + 1

        If i = 2 Then
            Me.bindDGMain2under()
        ElseIf i = 3 Then
            Me.bindDGMain3()
            Me.bindFilename3()
        End If
        '   MessageBox.Show("you selected the fifth tab: Tab No. " & i.ToString)
    End Sub

    Private Sub Button5_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        two_under()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Load_Excel_Details2under()
    End Sub

    Private Sub btnimportfile3_tod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnimportfile3_tod.Click

        If Me.tbPath3_tod.Text = "" Then
            MsgBox("โปรดเลือกไฟล์ EXCEL")
            Exit Sub
        End If

        Dim arr() As String
        arr = Me.tbPath3_tod.Text.Split("\")
        Dim tmp As String
        tmp = arr(UBound(arr))


        If chkFilename3_tod(tmp) Then
            MsgBox("มีไฟล์นี้ในระบบแล้วโปรดตรวจสอบ")
            bindFilename3_tod()
            Exit Sub
        Else
            insertFilename3_tod(tmp)

            bindFilename3_tod()
        End If

        Me.ImportFile3_tod(Me.tbPath3_tod.Text, "Sheet1$")
        btn_imp3_tod_Click(Me.btn_imp3_tod, e)
        If Not dt Is Nothing Then
            dt.Dispose()
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click

        OpenFileDialog2.InitialDirectory = "c:\"
        ' OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog2.RestoreDirectory = True
        OpenFileDialog2.Filter = "xls files (*.xls)|*.xls"
        OpenFileDialog2.FilterIndex = 2
        Me.OpenFileDialog2.ShowDialog()
        Me.tbPath3_tod.Text = OpenFileDialog2.FileName

    End Sub

    Private Sub btn_imp3_tod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_imp3_tod.Click

        insertDB_TMP3_tod()
        Me.bindDGMain3_tod()
        MsgBox("Complete")

    End Sub

    Private Sub btnResetForm3_tod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetForm3_tod.Click

        If MessageBox.Show("ต้องการล้างข้อมูล 3 ตัวโต๊ด เพื่อคำนวณใหม่ทั้งหมด", "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then


            Dim conn As New clsAccess("lotto.mdb")

            Dim sql As String
            sql = "update temp_tod_120 set price = 0,cutting=0,[get] = 0" & vbCrLf
            conn.ExecuteNonQuery(sql)
            Dim conn1 As New clsAccess("lotto.mdb")
            Dim sql1 As String
            sql1 = "delete from tb_filename3_tod "
            conn1.ExecuteNonQuery(sql1)
            resetForm3_tod()
        End If


    End Sub

    Private Sub btnProcess3_tod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess3_tod.Click

        If chkFirst3_tod() Then
            MsgBox("ไม่สามารถเก็บได้")
            Exit Sub
        End If

        max_tod_ = getMax3_tod()

        If chkMin3_tod(max_tod_) Then

            Dim sumLessPay As String
            sumLessPay = sumdefult3_tod(max_tod_)
            Dim countThan As String
            countThan = countThanDefult3_tod(max_tod_)
            bee3_tod(max_tod_, countThan, sumLessPay)
            Exit Sub
        End If

        If max_tod_ = "end" Then
            Exit Sub
        End If

        If max_tod_ = "False" Then
            MsgBox("NO ")
            Exit Sub
        End If
        Dim show As String
        Dim i As Integer
        Dim arrList1 As New ArrayList()
        For i = 0 To 10000
            show = q3_tod(max_tod_)
            If show <= 1000 - Int(My.Settings.tree100) Then
                Dim sumLessPay As String
                sumLessPay = sumdefult3_tod(max_tod_)
                Dim countThan As String
                countThan = countThanDefult3_tod(max_tod_)
                bee3_tod(max_tod_, countThan, sumLessPay)
                Exit For
            Else
                max_tod_ = max_tod_ - 10
            End If
        Next

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

        Dim strWhere As String
        strWhere = factor6(Me.tb3tod.Text)
        Pay_Tod(strWhere)
    End Sub





    Private Function lblCountON_() As String

        Dim sql As String = "select count(*) as c from temp where price > 0 "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return "แทงไปแล้ว : " & dt.Rows(0)("c").ToString() & " ตัว"
        End If

    End Function




    Private Function lblCountUnder_() As String

        Dim sql As String = "select count(*) as c from temp2under where price > 0 "
        Dim conn As New clsAccess("lotto.mdb")
        Dim dt As DataTable

        dt = conn.ReturnDataTable(sql)
        If dt.Rows.Count > 0 Then
            Return "แทงไปแล้ว : " & dt.Rows(0)("c").ToString() & " ตัว"
        End If

    End Function




    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Load_Excel_Details3_Tod()
    End Sub

    Private Sub dgTest_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgTest.CellContentClick

    End Sub
End Class
