<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormShop
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
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tb_SHOP_NO = New System.Windows.Forms.TextBox()
    Me.tb_SHOP_MEMO = New System.Windows.Forms.TextBox()
    Me.btn_SHOP_SAVE = New System.Windows.Forms.Button()
    Me.btn_SHOP_DELETE = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.Label1.Location = New System.Drawing.Point(72, 50)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(217, 40)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "賣場編號："
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.Label2.Location = New System.Drawing.Point(72, 125)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(217, 40)
    Me.Label2.TabIndex = 0
    Me.Label2.Text = "賣場名稱："
    '
    'tb_SHOP_NO
    '
    Me.tb_SHOP_NO.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.tb_SHOP_NO.Location = New System.Drawing.Point(278, 47)
    Me.tb_SHOP_NO.Name = "tb_SHOP_NO"
    Me.tb_SHOP_NO.Size = New System.Drawing.Size(292, 55)
    Me.tb_SHOP_NO.TabIndex = 1
    '
    'tb_SHOP_MEMO
    '
    Me.tb_SHOP_MEMO.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.tb_SHOP_MEMO.Location = New System.Drawing.Point(278, 122)
    Me.tb_SHOP_MEMO.Name = "tb_SHOP_MEMO"
    Me.tb_SHOP_MEMO.Size = New System.Drawing.Size(454, 55)
    Me.tb_SHOP_MEMO.TabIndex = 1
    '
    'btn_SHOP_SAVE
    '
    Me.btn_SHOP_SAVE.AutoSize = True
    Me.btn_SHOP_SAVE.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.btn_SHOP_SAVE.Location = New System.Drawing.Point(190, 225)
    Me.btn_SHOP_SAVE.Name = "btn_SHOP_SAVE"
    Me.btn_SHOP_SAVE.Size = New System.Drawing.Size(137, 51)
    Me.btn_SHOP_SAVE.TabIndex = 2
    Me.btn_SHOP_SAVE.Text = "儲存"
    Me.btn_SHOP_SAVE.UseVisualStyleBackColor = True
    '
    'btn_SHOP_DELETE
    '
    Me.btn_SHOP_DELETE.AutoSize = True
    Me.btn_SHOP_DELETE.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.btn_SHOP_DELETE.Location = New System.Drawing.Point(412, 225)
    Me.btn_SHOP_DELETE.Name = "btn_SHOP_DELETE"
    Me.btn_SHOP_DELETE.Size = New System.Drawing.Size(137, 51)
    Me.btn_SHOP_DELETE.TabIndex = 2
    Me.btn_SHOP_DELETE.Text = "刪除"
    Me.btn_SHOP_DELETE.UseVisualStyleBackColor = True
    '
    'FormShop
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(800, 310)
    Me.Controls.Add(Me.btn_SHOP_DELETE)
    Me.Controls.Add(Me.btn_SHOP_SAVE)
    Me.Controls.Add(Me.tb_SHOP_MEMO)
    Me.Controls.Add(Me.tb_SHOP_NO)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Name = "FormShop"
    Me.Text = "新增賣場"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents Label1 As Windows.Forms.Label
  Friend WithEvents Label2 As Windows.Forms.Label
  Friend WithEvents tb_SHOP_NO As Windows.Forms.TextBox
  Friend WithEvents tb_SHOP_MEMO As Windows.Forms.TextBox
  Friend WithEvents btn_SHOP_SAVE As Windows.Forms.Button
  Friend WithEvents btn_SHOP_DELETE As Windows.Forms.Button
End Class
