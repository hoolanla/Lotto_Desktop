Public Class frmSetting

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click



        My.Settings.two = two.Text
        My.Settings.twoUnder = twoUnder.Text
        My.Settings.tree500 = three500.Text
        My.Settings.tree100 = three100.Text
        My.Settings.percent2on = txtPercent.Text



        My.Settings.Save()

        FrmMain.Show()
        Me.Hide()
    End Sub

    Private Sub frmSetting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        two.Text = My.Settings.two
        twoUnder.Text = My.Settings.twoUnder
        three500.Text = My.Settings.tree500
        three100.Text = My.Settings.tree100
        txtPercent.Text = My.Settings.percent2on

    End Sub

    Private Sub btnCalPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalPay.Click


        frmCalPay.ShowDialog()
        Dim str As String


        two.Text = Math.Ceiling(CDbl(frmCalPay.two))
        twoUnder.Text = Math.Ceiling(CDbl(frmCalPay.two_under))
        three500.Text = Math.Ceiling(CDbl(frmCalPay.three))



        three100.Text = Math.Ceiling(CDbl(frmCalPay.tod)) * 6


    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub two_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles two.TextChanged

    End Sub
End Class