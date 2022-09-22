<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
    Me.WMSMenuStrip = New System.Windows.Forms.MenuStrip()
    Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
    Me.ChangeLogLevelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.TSCBViewLogLevel = New System.Windows.Forms.ToolStripComboBox()
    Me.ForTestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.T10F1S1PrintCarrierLabelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.T10F1S2PrintItemLabelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.T10F1S21PrintShippingDTLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
    Me.TabControl1 = New System.Windows.Forms.TabControl()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
    Me.lb_BarCode3 = New System.Windows.Forms.Label()
    Me.lb_BarCode2 = New System.Windows.Forms.Label()
    Me.lb_BarCode1 = New System.Windows.Forms.Label()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.ComboBox1 = New System.Windows.Forms.ComboBox()
    Me.tb_LotNo = New System.Windows.Forms.TextBox()
    Me.TabPage2 = New System.Windows.Forms.TabPage()
    Me.TabPage3 = New System.Windows.Forms.TabPage()
    Me.Button1 = New System.Windows.Forms.Button()
    Me.寄件編號 = New System.Windows.Forms.Label()
    Me.TextBox2 = New System.Windows.Forms.TextBox()
    Me.lb_BarCode4 = New System.Windows.Forms.Label()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.btn_START = New System.Windows.Forms.Button()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.WMSMenuStrip.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.Panel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'WMSMenuStrip
    '
    Me.WMSMenuStrip.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
    Me.WMSMenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
    Me.WMSMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem2, Me.ForTestToolStripMenuItem})
    Me.WMSMenuStrip.Location = New System.Drawing.Point(0, 0)
    Me.WMSMenuStrip.Name = "WMSMenuStrip"
    Me.WMSMenuStrip.Padding = New System.Windows.Forms.Padding(9, 2, 0, 2)
    Me.WMSMenuStrip.Size = New System.Drawing.Size(1256, 31)
    Me.WMSMenuStrip.TabIndex = 12
    Me.WMSMenuStrip.Text = "MenuStrip1"
    '
    'ToolStripMenuItem2
    '
    Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ChangeLogLevelToolStripMenuItem})
    Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
    Me.ToolStripMenuItem2.Size = New System.Drawing.Size(94, 27)
    Me.ToolStripMenuItem2.Text = "LogTool"
    '
    'ChangeLogLevelToolStripMenuItem
    '
    Me.ChangeLogLevelToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSCBViewLogLevel})
    Me.ChangeLogLevelToolStripMenuItem.Name = "ChangeLogLevelToolStripMenuItem"
    Me.ChangeLogLevelToolStripMenuItem.Size = New System.Drawing.Size(305, 34)
    Me.ChangeLogLevelToolStripMenuItem.Text = "Change View Log Level"
    '
    'TSCBViewLogLevel
    '
    Me.TSCBViewLogLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.TSCBViewLogLevel.Items.AddRange(New Object() {"Error", "WARN", "TRACE", "DEBUG", "ALL"})
    Me.TSCBViewLogLevel.MaxDropDownItems = 5
    Me.TSCBViewLogLevel.Name = "TSCBViewLogLevel"
    Me.TSCBViewLogLevel.Size = New System.Drawing.Size(121, 31)
    '
    'ForTestToolStripMenuItem
    '
    Me.ForTestToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.T10F1S1PrintCarrierLabelToolStripMenuItem, Me.T10F1S2PrintItemLabelToolStripMenuItem, Me.T10F1S21PrintShippingDTLToolStripMenuItem})
    Me.ForTestToolStripMenuItem.Name = "ForTestToolStripMenuItem"
    Me.ForTestToolStripMenuItem.Size = New System.Drawing.Size(88, 27)
    Me.ForTestToolStripMenuItem.Text = "ForTest"
    '
    'T10F1S1PrintCarrierLabelToolStripMenuItem
    '
    Me.T10F1S1PrintCarrierLabelToolStripMenuItem.Name = "T10F1S1PrintCarrierLabelToolStripMenuItem"
    Me.T10F1S1PrintCarrierLabelToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
    Me.T10F1S1PrintCarrierLabelToolStripMenuItem.Text = "T10F1S1_PrintCarrierLabel"
    '
    'T10F1S2PrintItemLabelToolStripMenuItem
    '
    Me.T10F1S2PrintItemLabelToolStripMenuItem.Name = "T10F1S2PrintItemLabelToolStripMenuItem"
    Me.T10F1S2PrintItemLabelToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
    Me.T10F1S2PrintItemLabelToolStripMenuItem.Text = "T10F1S2_PrintItemLabel"
    '
    'T10F1S21PrintShippingDTLToolStripMenuItem
    '
    Me.T10F1S21PrintShippingDTLToolStripMenuItem.Name = "T10F1S21PrintShippingDTLToolStripMenuItem"
    Me.T10F1S21PrintShippingDTLToolStripMenuItem.Size = New System.Drawing.Size(346, 34)
    Me.T10F1S21PrintShippingDTLToolStripMenuItem.Text = "T10F1S21_PrintShippingDTL"
    '
    'Timer1
    '
    Me.Timer1.Interval = 1000
    '
    'NotifyIcon1
    '
    Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
    Me.NotifyIcon1.Text = "NotifyIcon1"
    '
    'TabControl1
    '
    Me.TabControl1.Controls.Add(Me.TabPage1)
    Me.TabControl1.Controls.Add(Me.TabPage2)
    Me.TabControl1.Controls.Add(Me.TabPage3)
    Me.TabControl1.Location = New System.Drawing.Point(12, 45)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(1232, 777)
    Me.TabControl1.TabIndex = 13
    '
    'TabPage1
    '
    Me.TabPage1.BackColor = System.Drawing.Color.Transparent
    Me.TabPage1.Controls.Add(Me.btn_START)
    Me.TabPage1.Controls.Add(Me.Panel1)
    Me.TabPage1.Controls.Add(Me.Label1)
    Me.TabPage1.Controls.Add(Me.ComboBox1)
    Me.TabPage1.Controls.Add(Me.tb_LotNo)
    Me.TabPage1.Location = New System.Drawing.Point(4, 28)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(1224, 745)
    Me.TabPage1.TabIndex = 0
    Me.TabPage1.Text = "條碼收集"
    '
    'lb_BarCode3
    '
    Me.lb_BarCode3.AutoSize = True
    Me.lb_BarCode3.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.lb_BarCode3.Location = New System.Drawing.Point(36, 185)
    Me.lb_BarCode3.Name = "lb_BarCode3"
    Me.lb_BarCode3.Size = New System.Drawing.Size(215, 40)
    Me.lb_BarCode3.TabIndex = 2
    Me.lb_BarCode3.Text = "賣場(Lot)："
    '
    'lb_BarCode2
    '
    Me.lb_BarCode2.AutoSize = True
    Me.lb_BarCode2.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.lb_BarCode2.Location = New System.Drawing.Point(36, 114)
    Me.lb_BarCode2.Name = "lb_BarCode2"
    Me.lb_BarCode2.Size = New System.Drawing.Size(215, 40)
    Me.lb_BarCode2.TabIndex = 2
    Me.lb_BarCode2.Text = "賣場(Lot)："
    '
    'lb_BarCode1
    '
    Me.lb_BarCode1.AutoSize = True
    Me.lb_BarCode1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.lb_BarCode1.Location = New System.Drawing.Point(36, 37)
    Me.lb_BarCode1.Name = "lb_BarCode1"
    Me.lb_BarCode1.Size = New System.Drawing.Size(215, 40)
    Me.lb_BarCode1.TabIndex = 2
    Me.lb_BarCode1.Text = "賣場(Lot)："
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.Label1.Location = New System.Drawing.Point(351, 40)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(215, 40)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "賣場(Lot)："
    '
    'ComboBox1
    '
    Me.ComboBox1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.ComboBox1.FormattingEnabled = True
    Me.ComboBox1.Items.AddRange(New Object() {"7-11", "OK Mart", "Family"})
    Me.ComboBox1.Location = New System.Drawing.Point(49, 37)
    Me.ComboBox1.Name = "ComboBox1"
    Me.ComboBox1.Size = New System.Drawing.Size(235, 48)
    Me.ComboBox1.TabIndex = 1
    '
    'tb_LotNo
    '
    Me.tb_LotNo.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.tb_LotNo.Location = New System.Drawing.Point(572, 37)
    Me.tb_LotNo.Name = "tb_LotNo"
    Me.tb_LotNo.Size = New System.Drawing.Size(229, 55)
    Me.tb_LotNo.TabIndex = 0
    '
    'TabPage2
    '
    Me.TabPage2.BackColor = System.Drawing.Color.Transparent
    Me.TabPage2.Location = New System.Drawing.Point(4, 28)
    Me.TabPage2.Name = "TabPage2"
    Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage2.Size = New System.Drawing.Size(1224, 745)
    Me.TabPage2.TabIndex = 1
    Me.TabPage2.Text = "報表"
    '
    'TabPage3
    '
    Me.TabPage3.BackColor = System.Drawing.Color.Transparent
    Me.TabPage3.Controls.Add(Me.Button1)
    Me.TabPage3.Controls.Add(Me.寄件編號)
    Me.TabPage3.Controls.Add(Me.TextBox2)
    Me.TabPage3.Location = New System.Drawing.Point(4, 28)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Size = New System.Drawing.Size(1224, 745)
    Me.TabPage3.TabIndex = 2
    Me.TabPage3.Text = "查詢"
    '
    'Button1
    '
    Me.Button1.AutoSize = True
    Me.Button1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.Button1.Location = New System.Drawing.Point(475, 392)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(148, 50)
    Me.Button1.TabIndex = 2
    Me.Button1.Text = "查詢"
    Me.Button1.UseVisualStyleBackColor = True
    '
    '寄件編號
    '
    Me.寄件編號.AutoSize = True
    Me.寄件編號.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.寄件編號.Location = New System.Drawing.Point(35, 52)
    Me.寄件編號.Name = "寄件編號"
    Me.寄件編號.Size = New System.Drawing.Size(217, 40)
    Me.寄件編號.TabIndex = 1
    Me.寄件編號.Text = "寄件編號："
    '
    'TextBox2
    '
    Me.TextBox2.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.TextBox2.Location = New System.Drawing.Point(258, 49)
    Me.TextBox2.Name = "TextBox2"
    Me.TextBox2.Size = New System.Drawing.Size(641, 55)
    Me.TextBox2.TabIndex = 0
    '
    'lb_BarCode4
    '
    Me.lb_BarCode4.AutoSize = True
    Me.lb_BarCode4.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.lb_BarCode4.Location = New System.Drawing.Point(36, 253)
    Me.lb_BarCode4.Name = "lb_BarCode4"
    Me.lb_BarCode4.Size = New System.Drawing.Size(215, 40)
    Me.lb_BarCode4.TabIndex = 2
    Me.lb_BarCode4.Text = "賣場(Lot)："
    '
    'Panel1
    '
    Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.Panel1.Controls.Add(Me.ListBox1)
    Me.Panel1.Controls.Add(Me.lb_BarCode1)
    Me.Panel1.Controls.Add(Me.lb_BarCode4)
    Me.Panel1.Controls.Add(Me.lb_BarCode2)
    Me.Panel1.Controls.Add(Me.lb_BarCode3)
    Me.Panel1.Location = New System.Drawing.Point(6, 116)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(1212, 605)
    Me.Panel1.TabIndex = 3
    '
    'btn_START
    '
    Me.btn_START.AutoSize = True
    Me.btn_START.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
    Me.btn_START.Location = New System.Drawing.Point(893, 35)
    Me.btn_START.Name = "btn_START"
    Me.btn_START.Size = New System.Drawing.Size(187, 50)
    Me.btn_START.TabIndex = 4
    Me.btn_START.Text = "開始刷取"
    Me.btn_START.UseVisualStyleBackColor = True
    '
    'ListBox1
    '
    Me.ListBox1.FormattingEnabled = True
    Me.ListBox1.ItemHeight = 18
    Me.ListBox1.Location = New System.Drawing.Point(397, 37)
    Me.ListBox1.Name = "ListBox1"
    Me.ListBox1.Size = New System.Drawing.Size(120, 94)
    Me.ListBox1.TabIndex = 3
    '
    'FrmMain
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1256, 834)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.WMSMenuStrip)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
    Me.Name = "FrmMain"
    Me.Text = "BCS"
    Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
    Me.WMSMenuStrip.ResumeLayout(False)
    Me.WMSMenuStrip.PerformLayout()
    Me.TabControl1.ResumeLayout(False)
    Me.TabPage1.ResumeLayout(False)
    Me.TabPage1.PerformLayout()
    Me.TabPage3.ResumeLayout(False)
    Me.TabPage3.PerformLayout()
    Me.Panel1.ResumeLayout(False)
    Me.Panel1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents WMSMenuStrip As Windows.Forms.MenuStrip
  Friend WithEvents ToolStripMenuItem2 As Windows.Forms.ToolStripMenuItem
  Friend WithEvents ChangeLogLevelToolStripMenuItem As Windows.Forms.ToolStripMenuItem
  Friend WithEvents TSCBViewLogLevel As Windows.Forms.ToolStripComboBox
  Friend WithEvents ForTestToolStripMenuItem As Windows.Forms.ToolStripMenuItem
  Friend WithEvents Timer1 As Windows.Forms.Timer
  Friend WithEvents T10F1S1PrintCarrierLabelToolStripMenuItem As Windows.Forms.ToolStripMenuItem
  Friend WithEvents NotifyIcon1 As Windows.Forms.NotifyIcon
  Friend WithEvents T10F1S2PrintItemLabelToolStripMenuItem As Windows.Forms.ToolStripMenuItem
  Friend WithEvents T10F1S21PrintShippingDTLToolStripMenuItem As Windows.Forms.ToolStripMenuItem
  Friend WithEvents TabControl1 As Windows.Forms.TabControl
  Friend WithEvents TabPage1 As Windows.Forms.TabPage
  Friend WithEvents tb_LotNo As Windows.Forms.TextBox
  Friend WithEvents TabPage2 As Windows.Forms.TabPage
  Friend WithEvents TabPage3 As Windows.Forms.TabPage
  Friend WithEvents 寄件編號 As Windows.Forms.Label
  Friend WithEvents TextBox2 As Windows.Forms.TextBox
  Friend WithEvents Button1 As Windows.Forms.Button
  Friend WithEvents Label1 As Windows.Forms.Label
  Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
  Friend WithEvents lb_BarCode3 As Windows.Forms.Label
  Friend WithEvents lb_BarCode2 As Windows.Forms.Label
  Friend WithEvents lb_BarCode1 As Windows.Forms.Label
  Friend WithEvents Panel1 As Windows.Forms.Panel
  Friend WithEvents lb_BarCode4 As Windows.Forms.Label
  Friend WithEvents btn_START As Windows.Forms.Button
  Friend WithEvents ListBox1 As Windows.Forms.ListBox
End Class
