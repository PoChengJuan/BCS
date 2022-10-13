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
    Me.新增賣場ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
    Me.TabControl1 = New System.Windows.Forms.TabControl()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lb_Memo_Str = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CB_BarCode_SHOP = New System.Windows.Forms.ComboBox()
        Me.tb_BarCodeInput = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lb_BarCode1 = New System.Windows.Forms.Label()
        Me.lb_BarCode4 = New System.Windows.Forms.Label()
        Me.lb_BarCode2 = New System.Windows.Forms.Label()
        Me.lb_BarCode3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lb_Report_Memo_Str = New System.Windows.Forms.Label()
        Me.CB_Report_SHOP = New System.Windows.Forms.ComboBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.TimePicker_End = New System.Windows.Forms.DateTimePicker()
        Me.TimePicker_Start = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lbl_LotNo_Report = New System.Windows.Forms.Label()
        Me.btn_CreateReport = New System.Windows.Forms.Button()
        Me.DatePicker_End = New System.Windows.Forms.DateTimePicker()
        Me.DatePicker_Start = New System.Windows.Forms.DateTimePicker()
        Me.lbl_EndTime = New System.Windows.Forms.Label()
        Me.lbl_StartTime = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.寄件編號 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.WMSMenuStrip.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'WMSMenuStrip
        '
        Me.WMSMenuStrip.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.WMSMenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.WMSMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.新增賣場ToolStripMenuItem})
        Me.WMSMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.WMSMenuStrip.Name = "WMSMenuStrip"
        Me.WMSMenuStrip.Padding = New System.Windows.Forms.Padding(9, 2, 0, 2)
        Me.WMSMenuStrip.Size = New System.Drawing.Size(1256, 31)
        Me.WMSMenuStrip.TabIndex = 12
        Me.WMSMenuStrip.Text = "MenuStrip1"
        '
        '新增賣場ToolStripMenuItem
        '
        Me.新增賣場ToolStripMenuItem.Name = "新增賣場ToolStripMenuItem"
        Me.新增賣場ToolStripMenuItem.RightToLeftAutoMirrorImage = True
        Me.新增賣場ToolStripMenuItem.Size = New System.Drawing.Size(98, 27)
        Me.新增賣場ToolStripMenuItem.Text = "新增賣場"
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
        Me.TabControl1.Size = New System.Drawing.Size(1232, 526)
        Me.TabControl1.TabIndex = 13
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Transparent
        Me.TabPage1.Controls.Add(Me.lb_Memo_Str)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.CB_BarCode_SHOP)
        Me.TabPage1.Controls.Add(Me.tb_BarCodeInput)
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.ComboBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 28)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1224, 494)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "條碼收集"
        '
        'lb_Memo_Str
        '
        Me.lb_Memo_Str.AutoSize = True
        Me.lb_Memo_Str.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.lb_Memo_Str.Location = New System.Drawing.Point(466, 103)
        Me.lb_Memo_Str.Name = "lb_Memo_Str"
        Me.lb_Memo_Str.Size = New System.Drawing.Size(0, 40)
        Me.lb_Memo_Str.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.Label3.Location = New System.Drawing.Point(351, 103)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(137, 40)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "備註："
        '
        'CB_BarCode_SHOP
        '
        Me.CB_BarCode_SHOP.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.CB_BarCode_SHOP.FormattingEnabled = True
        Me.CB_BarCode_SHOP.Location = New System.Drawing.Point(473, 37)
        Me.CB_BarCode_SHOP.Name = "CB_BarCode_SHOP"
        Me.CB_BarCode_SHOP.Size = New System.Drawing.Size(283, 48)
        Me.CB_BarCode_SHOP.TabIndex = 6
        '
        'tb_BarCodeInput
        '
        Me.tb_BarCodeInput.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.tb_BarCodeInput.Location = New System.Drawing.Point(862, 36)
        Me.tb_BarCodeInput.Name = "tb_BarCodeInput"
        Me.tb_BarCodeInput.Size = New System.Drawing.Size(264, 55)
        Me.tb_BarCodeInput.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lb_BarCode1)
        Me.Panel1.Controls.Add(Me.lb_BarCode4)
        Me.Panel1.Controls.Add(Me.lb_BarCode2)
        Me.Panel1.Controls.Add(Me.lb_BarCode3)
        Me.Panel1.Location = New System.Drawing.Point(6, 170)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1212, 317)
        Me.Panel1.TabIndex = 3
        '
        'lb_BarCode1
        '
        Me.lb_BarCode1.AutoSize = True
        Me.lb_BarCode1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lb_BarCode1.Location = New System.Drawing.Point(36, 38)
        Me.lb_BarCode1.Name = "lb_BarCode1"
        Me.lb_BarCode1.Size = New System.Drawing.Size(215, 40)
        Me.lb_BarCode1.TabIndex = 2
        Me.lb_BarCode1.Text = "賣場(Lot)："
        '
        'lb_BarCode4
        '
        Me.lb_BarCode4.AutoSize = True
        Me.lb_BarCode4.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lb_BarCode4.Location = New System.Drawing.Point(36, 254)
        Me.lb_BarCode4.Name = "lb_BarCode4"
        Me.lb_BarCode4.Size = New System.Drawing.Size(215, 40)
        Me.lb_BarCode4.TabIndex = 2
        Me.lb_BarCode4.Text = "賣場(Lot)："
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
        'lb_BarCode3
        '
        Me.lb_BarCode3.AutoSize = True
        Me.lb_BarCode3.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lb_BarCode3.Location = New System.Drawing.Point(36, 184)
        Me.lb_BarCode3.Name = "lb_BarCode3"
        Me.lb_BarCode3.Size = New System.Drawing.Size(215, 40)
        Me.lb_BarCode3.TabIndex = 2
        Me.lb_BarCode3.Text = "賣場(Lot)："
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(351, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(137, 40)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "賣場："
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"7-11BarCode", "7-11QRCode", "OK Mart", "Family"})
        Me.ComboBox1.Location = New System.Drawing.Point(50, 38)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(235, 48)
        Me.ComboBox1.TabIndex = 1
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Transparent
        Me.TabPage2.Controls.Add(Me.lb_Report_Memo_Str)
        Me.TabPage2.Controls.Add(Me.CB_Report_SHOP)
        Me.TabPage2.Controls.Add(Me.Button2)
        Me.TabPage2.Controls.Add(Me.ComboBox2)
        Me.TabPage2.Controls.Add(Me.TimePicker_End)
        Me.TabPage2.Controls.Add(Me.TimePicker_Start)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.lbl_LotNo_Report)
        Me.TabPage2.Controls.Add(Me.btn_CreateReport)
        Me.TabPage2.Controls.Add(Me.DatePicker_End)
        Me.TabPage2.Controls.Add(Me.DatePicker_Start)
        Me.TabPage2.Controls.Add(Me.lbl_EndTime)
        Me.TabPage2.Controls.Add(Me.lbl_StartTime)
        Me.TabPage2.Location = New System.Drawing.Point(4, 28)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1224, 494)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "報表"
        '
        'lb_Report_Memo_Str
        '
        Me.lb_Report_Memo_Str.AutoSize = True
        Me.lb_Report_Memo_Str.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.lb_Report_Memo_Str.Location = New System.Drawing.Point(457, 106)
        Me.lb_Report_Memo_Str.Name = "lb_Report_Memo_Str"
        Me.lb_Report_Memo_Str.Size = New System.Drawing.Size(0, 40)
        Me.lb_Report_Memo_Str.TabIndex = 11
        '
        'CB_Report_SHOP
        '
        Me.CB_Report_SHOP.Font = New System.Drawing.Font("新細明體", 20.0!)
        Me.CB_Report_SHOP.FormattingEnabled = True
        Me.CB_Report_SHOP.Location = New System.Drawing.Point(198, 103)
        Me.CB_Report_SHOP.Name = "CB_Report_SHOP"
        Me.CB_Report_SHOP.Size = New System.Drawing.Size(235, 48)
        Me.CB_Report_SHOP.TabIndex = 10
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Button2.Location = New System.Drawing.Point(872, 40)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(244, 94)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "件數查詢"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"7-11BarCode", "7-11QRCode", "OK Mart", "Family"})
        Me.ComboBox2.Location = New System.Drawing.Point(198, 40)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(235, 48)
        Me.ComboBox2.TabIndex = 9
        '
        'TimePicker_End
        '
        Me.TimePicker_End.CalendarFont = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TimePicker_End.CustomFormat = "HH:mm:ss"
        Me.TimePicker_End.Font = New System.Drawing.Font("新細明體", 14.0!)
        Me.TimePicker_End.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TimePicker_End.Location = New System.Drawing.Point(474, 234)
        Me.TimePicker_End.Name = "TimePicker_End"
        Me.TimePicker_End.ShowUpDown = True
        Me.TimePicker_End.Size = New System.Drawing.Size(200, 41)
        Me.TimePicker_End.TabIndex = 7
        '
        'TimePicker_Start
        '
        Me.TimePicker_Start.CalendarFont = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TimePicker_Start.CustomFormat = "HH:mm:ss"
        Me.TimePicker_Start.Font = New System.Drawing.Font("新細明體", 14.0!)
        Me.TimePicker_Start.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TimePicker_Start.Location = New System.Drawing.Point(474, 171)
        Me.TimePicker_Start.Name = "TimePicker_Start"
        Me.TimePicker_Start.ShowUpDown = True
        Me.TimePicker_Start.Size = New System.Drawing.Size(200, 41)
        Me.TimePicker_Start.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(54, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 28)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "平台："
        '
        'lbl_LotNo_Report
        '
        Me.lbl_LotNo_Report.AutoSize = True
        Me.lbl_LotNo_Report.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lbl_LotNo_Report.Location = New System.Drawing.Point(54, 116)
        Me.lbl_LotNo_Report.Name = "lbl_LotNo_Report"
        Me.lbl_LotNo_Report.Size = New System.Drawing.Size(150, 28)
        Me.lbl_LotNo_Report.TabIndex = 5
        Me.lbl_LotNo_Report.Text = "賣場(Lot)："
        '
        'btn_CreateReport
        '
        Me.btn_CreateReport.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btn_CreateReport.Location = New System.Drawing.Point(56, 352)
        Me.btn_CreateReport.Name = "btn_CreateReport"
        Me.btn_CreateReport.Size = New System.Drawing.Size(244, 94)
        Me.btn_CreateReport.TabIndex = 4
        Me.btn_CreateReport.Text = "報表生成"
        Me.btn_CreateReport.UseVisualStyleBackColor = True
        '
        'DatePicker_End
        '
        Me.DatePicker_End.CustomFormat = ""
        Me.DatePicker_End.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.DatePicker_End.Location = New System.Drawing.Point(198, 234)
        Me.DatePicker_End.Name = "DatePicker_End"
        Me.DatePicker_End.Size = New System.Drawing.Size(246, 41)
        Me.DatePicker_End.TabIndex = 3
        '
        'DatePicker_Start
        '
        Me.DatePicker_Start.CustomFormat = ""
        Me.DatePicker_Start.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.DatePicker_Start.Location = New System.Drawing.Point(198, 171)
        Me.DatePicker_Start.Name = "DatePicker_Start"
        Me.DatePicker_Start.Size = New System.Drawing.Size(246, 41)
        Me.DatePicker_Start.TabIndex = 2
        '
        'lbl_EndTime
        '
        Me.lbl_EndTime.AutoSize = True
        Me.lbl_EndTime.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lbl_EndTime.Location = New System.Drawing.Point(52, 243)
        Me.lbl_EndTime.Name = "lbl_EndTime"
        Me.lbl_EndTime.Size = New System.Drawing.Size(152, 28)
        Me.lbl_EndTime.TabIndex = 1
        Me.lbl_EndTime.Text = "結束時間："
        '
        'lbl_StartTime
        '
        Me.lbl_StartTime.AutoSize = True
        Me.lbl_StartTime.Font = New System.Drawing.Font("新細明體", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lbl_StartTime.Location = New System.Drawing.Point(52, 180)
        Me.lbl_StartTime.Name = "lbl_StartTime"
        Me.lbl_StartTime.Size = New System.Drawing.Size(152, 28)
        Me.lbl_StartTime.TabIndex = 0
        Me.lbl_StartTime.Text = "開始時間："
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.Transparent
        Me.TabPage3.Controls.Add(Me.Button1)
        Me.TabPage3.Controls.Add(Me.寄件編號)
        Me.TabPage3.Controls.Add(Me.TextBox2)
        Me.TabPage3.Location = New System.Drawing.Point(4, 28)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1224, 494)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "查詢"
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Button1.Location = New System.Drawing.Point(59, 363)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(160, 75)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "查詢"
        Me.Button1.UseVisualStyleBackColor = True
        '
        '寄件編號
        '
        Me.寄件編號.AutoSize = True
        Me.寄件編號.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.寄件編號.Location = New System.Drawing.Point(34, 52)
        Me.寄件編號.Name = "寄件編號"
        Me.寄件編號.Size = New System.Drawing.Size(217, 40)
        Me.寄件編號.TabIndex = 1
        Me.寄件編號.Text = "寄件編號："
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("新細明體", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(258, 50)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(642, 55)
        Me.TextBox2.TabIndex = 0
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1256, 580)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.WMSMenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "FrmMain"
        Me.Text = "BCS"
        Me.WMSMenuStrip.ResumeLayout(False)
        Me.WMSMenuStrip.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents WMSMenuStrip As Windows.Forms.MenuStrip
    Friend WithEvents Timer1 As Windows.Forms.Timer
    Friend WithEvents NotifyIcon1 As Windows.Forms.NotifyIcon
    Friend WithEvents TabControl1 As Windows.Forms.TabControl
    Friend WithEvents TabPage1 As Windows.Forms.TabPage
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
  Friend WithEvents btn_CreateReport As Windows.Forms.Button
  Friend WithEvents DatePicker_End As Windows.Forms.DateTimePicker
  Friend WithEvents DatePicker_Start As Windows.Forms.DateTimePicker
  Friend WithEvents lbl_EndTime As Windows.Forms.Label
  Friend WithEvents lbl_StartTime As Windows.Forms.Label
    Friend WithEvents lbl_LotNo_Report As Windows.Forms.Label
    Friend WithEvents TimePicker_End As Windows.Forms.DateTimePicker
    Friend WithEvents TimePicker_Start As Windows.Forms.DateTimePicker
    Friend WithEvents tb_BarCodeInput As Windows.Forms.TextBox
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents 新增賣場ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents CB_Report_SHOP As Windows.Forms.ComboBox
    Friend WithEvents CB_BarCode_SHOP As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents lb_Memo_Str As Windows.Forms.Label
    Friend WithEvents lb_Report_Memo_Str As Windows.Forms.Label
End Class
