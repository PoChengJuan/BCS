<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormMsg
  Inherits System.Windows.Forms.Form

  'Form 覆寫 Dispose 以清除元件清單。
  <System.Diagnostics.DebuggerNonUserCode()>
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
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.lb_ItemCnt_Str = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'lb_ItemCnt_Str
    '
    Me.lb_ItemCnt_Str.Anchor = System.Windows.Forms.AnchorStyles.Left
    Me.lb_ItemCnt_Str.AutoSize = True
    Me.lb_ItemCnt_Str.Font = New System.Drawing.Font("新細明體", 20.0!)
    Me.lb_ItemCnt_Str.Location = New System.Drawing.Point(53, 27)
    Me.lb_ItemCnt_Str.Name = "lb_ItemCnt_Str"
    Me.lb_ItemCnt_Str.Size = New System.Drawing.Size(122, 40)
    Me.lb_ItemCnt_Str.TabIndex = 0
    Me.lb_ItemCnt_Str.Text = "Label1"
    Me.lb_ItemCnt_Str.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'FormMsg
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.AutoSize = True
    Me.ClientSize = New System.Drawing.Size(528, 298)
    Me.Controls.Add(Me.lb_ItemCnt_Str)
    Me.Name = "FormMsg"
    Me.ShowIcon = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "訊息"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents lb_ItemCnt_Str As Windows.Forms.Label
End Class
