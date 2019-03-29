<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.two = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.three500 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.three100 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnCalPay = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.twoUnder = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPercent = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'two
        '
        Me.two.Location = New System.Drawing.Point(97, 33)
        Me.two.MaxLength = 2
        Me.two.Name = "two"
        Me.two.Size = New System.Drawing.Size(55, 20)
        Me.two.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(47, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "2 µ—«∫π"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(48, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "3 µ—«‡µÁß"
        '
        'three500
        '
        Me.three500.Location = New System.Drawing.Point(98, 85)
        Me.three500.MaxLength = 3
        Me.three500.Name = "three500"
        Me.three500.Size = New System.Drawing.Size(55, 20)
        Me.three500.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(47, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "3 µ—«‚µÍ¥"
        '
        'three100
        '
        Me.three100.Location = New System.Drawing.Point(97, 112)
        Me.three100.MaxLength = 3
        Me.three100.Name = "three100"
        Me.three100.Size = New System.Drawing.Size(55, 20)
        Me.three100.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.Color.Red
        Me.Button1.Location = New System.Drawing.Point(97, 182)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(55, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnCalPay
        '
        Me.btnCalPay.ForeColor = System.Drawing.Color.Green
        Me.btnCalPay.Location = New System.Drawing.Point(20, 211)
        Me.btnCalPay.Name = "btnCalPay"
        Me.btnCalPay.Size = New System.Drawing.Size(210, 23)
        Me.btnCalPay.TabIndex = 7
        Me.btnCalPay.Text = "§”π«≥‡ªÕ√Ï‡´πµÏ®Ë“¬"
        Me.btnCalPay.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(47, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "2 µ—«≈Ë“ß"
        '
        'twoUnder
        '
        Me.twoUnder.Location = New System.Drawing.Point(97, 59)
        Me.twoUnder.MaxLength = 2
        Me.twoUnder.Name = "twoUnder"
        Me.twoUnder.Size = New System.Drawing.Size(55, 20)
        Me.twoUnder.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(33, 141)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "‡ªÕ√Ï‡´πµÏ"
        '
        'txtPercent
        '
        Me.txtPercent.Location = New System.Drawing.Point(96, 138)
        Me.txtPercent.MaxLength = 3
        Me.txtPercent.Name = "txtPercent"
        Me.txtPercent.Size = New System.Drawing.Size(55, 20)
        Me.txtPercent.TabIndex = 10
        '
        'frmSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(242, 246)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPercent)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.twoUnder)
        Me.Controls.Add(Me.btnCalPay)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.three100)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.three500)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.two)
        Me.Name = "frmSetting"
        Me.Text = "frmSetting"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents two As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents three500 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents three100 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnCalPay As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents twoUnder As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPercent As System.Windows.Forms.TextBox
End Class
