<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMSGTest
	Inherits System.Windows.Forms.Form

	'Form 覆寫 Dispose 以清除元件清單。
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'為 Windows Form 設計工具的必要項
	Private components As System.ComponentModel.IContainer

	'注意: 以下為 Windows Form 設計工具所需的程序
	'可以使用 Windows Form 設計工具進行修改。
	'請勿使用程式碼編輯器進行修改。
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
    Me.gb_SendMessage = New System.Windows.Forms.GroupBox()
    Me.txt_SendMessage = New System.Windows.Forms.TextBox()
    Me.gb_ResultMessage = New System.Windows.Forms.GroupBox()
    Me.txt_ResultMessage = New System.Windows.Forms.TextBox()
    Me.btn_SendMessage = New System.Windows.Forms.Button()
        Me.gb_SendMessage.SuspendLayout()
        Me.gb_ResultMessage.SuspendLayout()
        Me.SuspendLayout()
        '
        'gb_SendMessage
        '
        Me.gb_SendMessage.Controls.Add(Me.txt_SendMessage)
        Me.gb_SendMessage.Location = New System.Drawing.Point(14, 14)
        Me.gb_SendMessage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb_SendMessage.Name = "gb_SendMessage"
        Me.gb_SendMessage.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb_SendMessage.Size = New System.Drawing.Size(818, 571)
        Me.gb_SendMessage.TabIndex = 2
        Me.gb_SendMessage.TabStop = False
        Me.gb_SendMessage.Text = "SendMessage"
        '
        'txt_SendMessage
        '
        Me.txt_SendMessage.Location = New System.Drawing.Point(14, 29)
        Me.txt_SendMessage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_SendMessage.Multiline = True
        Me.txt_SendMessage.Name = "txt_SendMessage"
        Me.txt_SendMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_SendMessage.Size = New System.Drawing.Size(804, 533)
        Me.txt_SendMessage.TabIndex = 0
        '
        'gb_ResultMessage
        '
        Me.gb_ResultMessage.Controls.Add(Me.txt_ResultMessage)
        Me.gb_ResultMessage.Location = New System.Drawing.Point(27, 632)
        Me.gb_ResultMessage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb_ResultMessage.Name = "gb_ResultMessage"
        Me.gb_ResultMessage.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gb_ResultMessage.Size = New System.Drawing.Size(818, 176)
        Me.gb_ResultMessage.TabIndex = 3
        Me.gb_ResultMessage.TabStop = False
        Me.gb_ResultMessage.Text = "ResultMessage"
        '
        'txt_ResultMessage
        '
        Me.txt_ResultMessage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_ResultMessage.Location = New System.Drawing.Point(7, 29)
        Me.txt_ResultMessage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txt_ResultMessage.Multiline = True
        Me.txt_ResultMessage.Name = "txt_ResultMessage"
        Me.txt_ResultMessage.Size = New System.Drawing.Size(804, 140)
        Me.txt_ResultMessage.TabIndex = 3
        '
        'btn_SendMessage
        '
        Me.btn_SendMessage.Location = New System.Drawing.Point(885, 545)
        Me.btn_SendMessage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btn_SendMessage.Name = "btn_SendMessage"
        Me.btn_SendMessage.Size = New System.Drawing.Size(135, 100)
        Me.btn_SendMessage.TabIndex = 4
        Me.btn_SendMessage.Text = "SendMessage"
        Me.btn_SendMessage.UseVisualStyleBackColor = True
        '
        'FrmMSGTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1083, 821)
        Me.Controls.Add(Me.btn_SendMessage)
        Me.Controls.Add(Me.gb_ResultMessage)
        Me.Controls.Add(Me.gb_SendMessage)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "FrmMSGTest"
        Me.Text = "FrmMSGTest"
        Me.gb_SendMessage.ResumeLayout(False)
        Me.gb_SendMessage.PerformLayout()
        Me.gb_ResultMessage.ResumeLayout(False)
        Me.gb_ResultMessage.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents gb_SendMessage As Windows.Forms.GroupBox
	Friend WithEvents txt_SendMessage As Windows.Forms.TextBox
	Friend WithEvents gb_ResultMessage As Windows.Forms.GroupBox
	Friend WithEvents txt_ResultMessage As Windows.Forms.TextBox
	Friend WithEvents btn_SendMessage As Windows.Forms.Button
End Class
