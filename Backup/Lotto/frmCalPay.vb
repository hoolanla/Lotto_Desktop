Public Class frmCalPay

    Private Sub frmCalPay_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Dim m_two As String
    Public Property two()
        Get
            Return m_two
        End Get
        Set(ByVal value)
            m_two = value
        End Set
    End Property


    Dim m_two_under As String
    Public Property two_under()
        Get
            Return m_two_under
        End Get
        Set(ByVal value)
            m_two_under = value
        End Set
    End Property


    Dim m_three As String
    Public Property three()
        Get
            Return m_three
        End Get
        Set(ByVal value)
            m_three = value
        End Set
    End Property


    Dim m_tod As String
    Public Property tod()
        Get
            Return m_tod
        End Get
        Set(ByVal value)
            m_tod = value
        End Set
    End Property






    Public Overloads Sub ShowDialog(ByVal mcolumnDesc As String, ByVal mColumnWidth As String, ByVal mTableName As String, Optional ByVal parent As IWin32Window = Nothing)


        Dim i As Integer = 0
        Dim sArray() As String
        Dim nFormWidth As Long = 0
        Me.Cursor = Cursors.WaitCursor
        Try


        Catch ex As Exception
        Finally
            Me.Cursor = Cursors.Default

        End Try


        ' Show the dialog.
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.ShowDialog(parent)
    End Sub
    Public Function getString() As String
        Return "test"
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim tmp1, tmp2 As Integer
        tmp1 = Int(Me.tb2_1.Text) - Int(Me.tb2_2.Text)
        tmp2 = tmp1 - Int(Me.tb2_3.Text)

        Me.tb2_4.Text = 100 - ((tmp2 / tmp1) * 100)

        GroupBox1.BackColor = Color.AliceBlue

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click



        Dim tmp1, tmp2 As Integer
        tmp1 = (Int(Me.tb22_2.Text) * Int(Me.tb22_3.Text) / 100) + Int(Me.tb22_3.Text)


        Me.tb22_4.Text = tmp1

        GroupBox1.BackColor = Color.AliceBlue
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click


        Dim tmp1, tmp2 As Integer
        tmp1 = Int(Me.tb2under1.Text) - Int(Me.tb2under2.Text)
        tmp2 = tmp1 - Int(Me.tb2under3.Text)

        Me.tb2under4.Text = 100 - ((tmp2 / tmp1) * 100)

        GroupBox3.BackColor = Color.AliceBlue
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Dim tmp1, tmp2 As Integer
        tmp1 = (Int(Me.tb2under2_2.Text) * Int(Me.tb2under2_3.Text) / 100) + Int(Me.tb2under2_3.Text)

        Me.tb2under2_4.Text = tmp1

        GroupBox3.BackColor = Color.AliceBlue
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click



        If GroupBox1.BackColor <> Color.AliceBlue Or GroupBox2.BackColor <> Color.AliceBlue Or GroupBox3.BackColor <> Color.AliceBlue Then
            MsgBox("โปรดคำนวณทุกช่องให้ครบก่อน")
            Exit Sub
        End If








        If Int(tb2_4.Text) > Int(tb22_4.Text) Then
            m_two = tb2_4.Text
        Else
            m_two = tb22_4.Text
        End If


        If Int(tb2under4.Text) > Int(tb2under2_4.Text) Then

            m_two_under = tb2under4.Text
        Else
            m_two_under = tb2under2_4.Text
        End If




        If Int(tb3_4.Text) > Int(tb32_4.Text) Then

            m_three = tb3_4.Text
        Else
            m_three = tb32_4.Text
        End If



        If Int(tb3tod_4.Text) > Int(tb3tod2_4.Text) Then

            m_tod = tb3tod_4.Text
        Else
            m_tod = tb3tod2_4.Text
        End If



        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim tmp1, tmp2 As Integer
        tmp1 = Int(Me.tb3_1.Text) - Int(Me.tb3_2.Text)
        tmp2 = tmp1 - Int(Me.tb3_3.Text)

        Me.tb3_4.Text = 1000 - ((tmp2 / tmp1) * 1000)

        GroupBox2.BackColor = Color.AliceBlue


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click




        Dim tmp1, tmp2 As Integer
        tmp1 = (Int(Me.tb32_2.Text) * Int(Me.tb32_3.Text) / 1000) + Int(Me.tb32_3.Text)
        Me.tb32_4.Text = tmp1



        GroupBox2.BackColor = Color.AliceBlue

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        Dim tmp1, tmp2 As Integer
        tmp1 = 100 - Int(Me.tb3tod_2.Text)
        tmp2 = tmp1 - (Int(Me.tb3tod_3.Text))

        Me.tb3tod_4.Text = 100 - ((tmp2 / tmp1) * 100)

        GroupBox4.BackColor = Color.AliceBlue
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click


        Dim tmp1 As Integer
        tmp1 = (Int(Me.tb3tod2_2.Text) * Int(Me.tb3tod2_3.Text) / 100) + Int(Me.tb3tod2_3.Text)


        Me.tb3tod2_4.Text = tmp1

        GroupBox4.BackColor = Color.AliceBlue
    End Sub
End Class