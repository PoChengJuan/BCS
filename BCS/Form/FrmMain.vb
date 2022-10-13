Imports System.IO
Imports System.IO.Ports
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Serialization



Public Class FrmMain
  Public interfaceDB As ClsInterfaceDB
  Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon
  Public RefreshAliveTime As Integer = 3000

  Public DBTool As eCA_DBTool.clsDBTool
  Public gdicSHOP As New Dictionary(Of String, clsBCS_M_SHOP)
  Dim BarCode1 = ""
  Dim BarCode2 = ""
  Dim BarCode3 = ""
  Sub New()

    ' 設計工具需要此呼叫。
    InitializeComponent()
    ' 在 InitializeComponent() 呼叫之後加入所有初始設定。		 
    LogTool = New eCALogTool.CLogTool
  End Sub


  Private Function LogToolSetting() As Boolean
    Try
      LogTool.APName = My.Application.Info.AssemblyName
      'LogTool.trcListView = lswtrc
      LogTool.ViewColor = True
      LogTool.InitialEnd = True

      Return True
    Catch ex As Exception
      MsgBox(ex.ToString)
      Return False
    End Try
  End Function
  Private Function I_Load_Config() As Boolean
    Try
      SendMessageToLog("Load Config...", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      '檢查程式根目錄內的Config.xml檔案是否存在
      If My.Computer.FileSystem.FileExists(ConfigPath) Then
        '取得Config.xml檔案
        Dim ConfigXML As XDocument = XDocument.Load(ConfigPath)

        '抓出Public的資料
        Dim Common As IEnumerable = From c In ConfigXML.<Config>.<Public>
        For Each Pub As XElement In Common
          'LogPath
          LogPath = Pub.<LogPath>.Value

          InterfaceType = Pub.<InterfaceType>.Value
          CLIENT_ID = Pub.<CLIENT_ID>.Value
          IP = Pub.<IP>.Value
          'If CLIENT_ID.Length = 0 And IP.Length = 0 Then
          '	'MsgBox("ClentID and IP config is empty")
          '	'Return False
          'ElseIf CLIENT_ID.Length = 0 And IP.Length = 0 Then
          '	MsgBox("ClentID and IP config Error")
          '	Return False
          'End If
          HTTPPath = Pub.<HTTPPath>.Value
          UpLoadKey = Pub.<HTTPUpLoadKey>.Value
          gFileRootPath = Pub.<FileRootPath>.Value
          gReportFunction = Pub.<ReportFunction>.Value
          If Strings.Right(LogPath, 1) <> "\" Then
            LogPath = LogPath & "\"
          End If
          '檢查路徑是否包含到磁碟
          Dim hasDrive As Integer = 0
          Dim linq = From driver In IO.Directory.GetLogicalDrives Where Strings.Left(driver, 1).ToLower() = Strings.Left(LogPath, 1).ToLower()
          For Each driveletter As String In linq
            hasDrive = 1
            Exit For
          Next
          If hasDrive = 0 Then
            MsgBox("Please check the LogPath in Config.xml.")
            Return False
          End If

          'ExportPath
          ExportPath = Pub.<ExportPath>.Value
          If Strings.Right(ExportPath, 1) <> "\" Then
            ExportPath = ExportPath & "\"
          End If
          If ExportPath = "" Then
            ExportPath = LogPath & "Export\"
          End If
          '檢查路徑是否包含到磁碟
          hasDrive = 0
          linq = From driver In IO.Directory.GetLogicalDrives Where Strings.Left(driver, 1).ToLower() = Strings.Left(ExportPath, 1).ToLower()
          For Each driveletter As String In linq
            hasDrive = 1
            Exit For
          Next
          If hasDrive = 0 Then
            MsgBox("Please check the ExportPath in Config.xml.")
            Return False
          End If

          '匯出時用到的Sample檔案的目錄
          SamplePath = Application.StartupPath & Pub.<SampleFolder>.Value
          If Strings.Right(SamplePath, 1) <> "\" Then
            SamplePath = SamplePath & "\"
          End If

          '匯出時用到的PrintFormatFile的檔案位置
          PrintFormatFile = Application.StartupPath & Pub.<PrintFormatFile>.Value

          'TraceLevel
          LogTool.LogPath = LogPath
          LogTool.TraceLevel = Pub.<TraceLv>.Value
          'MaxViewLine
          LogTool.MaxViewLine = Pub.<MaxViewLine>.Value
          'ViewLV
          LogTool.ViewLV = Pub.<ViewLV>.Value
          'RefreshAliveTime
          RefreshAliveTime = IIf(Pub.<RefreshAliveTime>.Value Is Nothing OrElse Pub.<RefreshAliveTime>.Value = 0, 5000, Pub.<RefreshAliveTime>.Value)
          'WCFHostIpPort
          WCFHostIpPort = Pub.<WCFHostIpPort>.Value
          'PDFPath'
          strPDFPath = Pub.<PDFPath>.Value
        Next

        '抓出DBList的資料
        Dim DB As IEnumerable = From c In ConfigXML.<Config>.<DBList>
        For Each DBList As XElement In DB
          '取得每一個DB連線設定的資料
          For Each DBInfo As XElement In DBList.Nodes
            'DBTool记录挡案资料夹
            DBTool_Name = "Print_" & DBInfo.<DBName>.Value & "_" & DBInfo.<DBType>.Value & "_" & DBInfo.<UID>.Value
            Dim objDBInfo As New eCA_DBTool.clsDBTool(DBTool_Name)
            'DBType
            objDBInfo.m_nDBType = DBInfo.<DBType>.Value
            'DBServer
            objDBInfo.m_szDBServer = DBInfo.<DBServer>.Value
            'DBName
            objDBInfo.m_szDBName = DBInfo.<DBName>.Value
            'UID
            objDBInfo.m_szDBUID = DBInfo.<UID>.Value
            'PWD
            objDBInfo.m_szDBPWD = DBInfo.<PWD>.Value
            DBTool = objDBInfo
          Next
        Next
        '抓出PRINTLIST的資料
        Dim PRINTER As IEnumerable = From c In ConfigXML.<Config>.<PRINTERLIST>
        For Each PRINTERList As XElement In PRINTER
          For Each PRINTERInfo As XElement In PRINTERList.Nodes
            PRINTER_NAME_A4 = PRINTERInfo.<PRINTER_NAME_A4>.Value
            PRINTER_NAME_10x10Label = PRINTERInfo.<PRINTER_NAME_10x10Label>.Value
            PRINTER_LIMITED_COUNT = PRINTERInfo.<PRINTER_LIMITED_COUNT>.Value
            If PRINTERInfo.<PRINTER_Label_Limited>.Value IsNot Nothing Then
              PRINTER_Label_Limited = PRINTERInfo.<PRINTER_Label_Limited>.Value
            End If
            If PRINTERInfo.<PageTop>.Value IsNot Nothing Then
              blnPageTop = CBool(PRINTERInfo.<PageTop>.Value)
            End If
            'PRINTER_NAME_A4 = PRINTERInfo.Value
          Next
        Next
      Else
        MsgBox("Cannot find config file: " & ConfigPath)
        Return False
      End If
      SendMessageToLog("Load Config Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Return True
    Catch ex As Exception
      MsgBox(ex.ToString, MsgBoxStyle.Question)
      Return False
    End Try
  End Function
  Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      If UBound(Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName)) > 0 Then
        MsgBox(Process.GetCurrentProcess.ProcessName & " is running ! Path[" & Application.StartupPath & "]", vbCritical, "Error")
        End
      End If
      '讀取config
      If Not I_Load_Config() Then MsgBox("Load Config Error") : Return

      'Setting
      If Not LogToolSetting() Then MsgBox("LogToolSetting Error") : Return

      'TextBox1.Select()
      Me.Hide()

      'TSCBViewLogLevel.SelectedIndex = LogTool.ViewLV - 1

      NotifyIcon = New System.Windows.Forms.NotifyIcon()
      NotifyIcon1.Text = My.Application.Info.AssemblyName

      SendMessageToLog("[PrintTool Start]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Handling() '-根據Config 確認這次的interface

      Timer1.Interval = RefreshAliveTime
      Timer1.Enabled = True

      gdicSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY("")
      For Each objSHOP In gdicSHOP.Values
        CB_BarCode_SHOP.Items.Add(objSHOP.SHOP)
      Next
      For Each objSHOP In gdicSHOP.Values
        CB_Report_SHOP.Items.Add(objSHOP.SHOP)
      Next
      '掃描頁面
      If gReportFunction = "0" Then
        TabPage2.Parent = Nothing
      End If
      lb_BarCode1.Text = ""
      lb_BarCode2.Text = ""
      lb_BarCode3.Text = ""
      lb_BarCode4.Text = ""
      ComboBox1.SelectedIndex = 0
      tb_BarCodeInput.Select()

      '報表頁面
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Function I_BuildToolsConnection(ByRef RetMsg As String) As Boolean
    Try
      SendMessageToLog("BuildToolsConnection...", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      If DBTool.OpenConnection(RetMsg) Then
        SendMessageToLog("BuildToolsConnection OpenConnection Success", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        Return True
      Else
        SendMessageToLog("BuildToolsConnection OpenConnection Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False

      End If

    Catch ex As Exception
      MsgBox(ex.ToString)
      Return False
    End Try
  End Function
  Private Sub Handling()
    Try
      If InterfaceType = EnuInterfaceType.DB Then
        'lblinterfacetype.Text = EnuInterfaceType.DB.ToString
        Dim RetMsg = ""
        If I_BuildToolsConnection(RetMsg) = False Then 'OPEN CONNECTION
          SendMessageToLog(RetMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
          MsgBox(RetMsg)
        Else '-DB 連線成功
          interfaceDB = New ClsInterfaceDB(LogTool, DBTool)
        End If



      ElseIf InterfaceType = EnuInterfaceType.RabbitMQ Then


      Else
        MsgBox("InterFaceType Error  check Config")

      End If

    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub

  Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    Try
      'lblThrCount.Text = interfaceDB.int_tGUIDBHandle


    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub



  Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
    If Me.WindowState = FormWindowState.Minimized Then
      Me.Hide()
      NotifyIcon1.Visible = True
    Else
      NotifyIcon1.Visible = False
    End If
  End Sub

  Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
    Try
      NotifyIcon1.Visible = False
      Me.Show()
      Me.WindowState = FormWindowState.Normal
    Catch ex As Exception

    End Try
  End Sub





  Private Sub TabControl_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
    If e.TabPage.TabIndex = TabPage1.TabIndex Then
      tb_BarCodeInput.Select()
    ElseIf e.TabPage.TabIndex = TabPage3.TabIndex Then
      TextBox2.Select()
    End If

  End Sub
  Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    Try
      Dim result_msg = ""
      Dim PlatForm = ""
      Select Case ComboBox2.SelectedIndex
        Case 0
          PlatForm = "7-11BarCode"
        Case 1
          PlatForm = "7-11QRCode"
        Case 2
          PlatForm = "OK Mart"
        Case 3
          PlatForm = "Family"
        Case Else
          PlatForm = ""
      End Select
      Dim index = CB_Report_SHOP.SelectedIndex

      Dim LotNo = "" 'tb_LotNo_Report.Text
      If index = -1 Then
        LotNo = ""
      Else
        LotNo = CB_Report_SHOP.Items(index)
      End If
      Dim Start_Date = DatePicker_Start.Value.Date.ToString("yyyy/MM/dd")

      Dim Start_Hour = TimePicker_Start.Value.Hour.ToString.PadLeft(2, "0")
      Dim Start_Min = TimePicker_Start.Value.Minute.ToString.PadLeft(2, "0")
      Dim Start_Second = TimePicker_Start.Value.Second.ToString.PadLeft(2, "0")
      Dim Start_Time = Start_Hour & ":" & Start_Min & ":" & Start_Second 'TimePicker_Start.Value.TimeOfDay.ToString("hh:mm:ss")
      Dim End_Date = DatePicker_End.Value.Date.ToString("yyyy/MM/dd")
      Dim End_Hour = TimePicker_End.Value.Hour.ToString.PadLeft(2, "0")
      Dim End_Min = TimePicker_End.Value.Minute.ToString.PadLeft(2, "0")
      Dim End_Second = TimePicker_End.Value.Second.ToString.PadLeft(2, "0")
      Dim End_Time = End_Hour & ":" & End_Min & ":" & End_Second ' TimePicker_End.Value.TimeOfDay.ToString("hh:mm:ss")

      Dim Start_DateTime = Start_Date & " " & Start_Time
      Dim End_DateTime = End_Date & " " & End_Time

      Dim dicStore_Item = STORE_ITEMManagement.GetDataDictionaryByItemReport(PlatForm, LotNo, Start_DateTime, End_DateTime)
      If dicStore_Item.Any = False Then
        result_msg = "查無條碼"
        MsgBox(result_msg)
        Return
      End If

      Dim dicStore_Item_711BarCode As New Dictionary(Of String, clsSTORE_ITEM)
      Dim dicStore_Item_711QRCode As New Dictionary(Of String, clsSTORE_ITEM)
      Dim dicStore_Item_OK As New Dictionary(Of String, clsSTORE_ITEM)
      Dim dicStore_Item_Family As New Dictionary(Of String, clsSTORE_ITEM)

      For Each obj In dicStore_Item.Values
        If obj.PlatForm = "7-11BarCode" Then
          dicStore_Item_711BarCode.Add(obj.gid, obj)
        ElseIf obj.PlatForm = "7-11QRCode" Then
          dicStore_Item_711QRCode.Add(obj.gid, obj)
        ElseIf obj.PlatForm = "OK Mart" Then
          dicStore_Item_OK.Add(obj.gid, obj)
        ElseIf obj.PlatForm = "Family" Then
          dicStore_Item_Family.Add(obj.gid, obj)
        End If
      Next
      result_msg = "7-11 BarCode：" & dicStore_Item_711BarCode.Count & "件" & vbCrLf &
                   "7-11 QRCode：" & dicStore_Item_711QRCode.Count & "件" & vbCrLf &
                   "OK：" & dicStore_Item_OK.Count & "件" & vbCrLf &
                   "全家：" & dicStore_Item_Family.Count & "件"
      'FormMsg.lb_ItemCount_Str.Text = result_msg

      Dim _form = FormMsg.CreateForm("FormMsg", result_msg)
      If _form IsNot Nothing Then
        _form.Show()
      End If

      Return
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub


  Private Sub btn_CreateReport_Click(sender As Object, e As EventArgs) Handles btn_CreateReport.Click

    Dim PlatForm = ""
    Select Case ComboBox2.SelectedIndex
      Case 0
        PlatForm = "7-11BarCode"
      Case 1
        PlatForm = "7-11QRCode"
      Case 1
        PlatForm = "OK Mart"
      Case 2
        PlatForm = "Family"
      Case Else
        MsgBox("請選擇平台")
    End Select

    Dim index = CB_Report_SHOP.SelectedIndex
    Dim SHOP_NO = CB_Report_SHOP.Items(index)
    Dim LotNo = SHOP_NO 'tb_LotNo_Report.Text
    'Dim Start_Date = DatePicker_Start.Value.Date.ToString("yyyy/MM/dd")

    'Dim Start_Time = TimePicker_Start.Value.TimeOfDay.ToString()
    'Dim End_Date = DatePicker_End.Value.Date.ToString("yyyy/MM/dd")
    'Dim End_Time = TimePicker_End.Value.TimeOfDay.ToString()
    Dim Start_Date = DatePicker_Start.Value.Date.ToString("yyyy/MM/dd")
    Dim Start_Hour = TimePicker_Start.Value.Hour.ToString.PadLeft(2, "0")
    Dim Start_Min = TimePicker_Start.Value.Minute.ToString.PadLeft(2, "0")
    Dim Start_Second = TimePicker_Start.Value.Second.ToString.PadLeft(2, "0")
    Dim Start_Time = Start_Hour & ":" & Start_Min & ":" & Start_Second 'TimePicker_Start.Value.TimeOfDay.ToString("hh:mm:ss")
    Dim End_Date = DatePicker_End.Value.Date.ToString("yyyy/MM/dd")
    Dim End_Hour = TimePicker_End.Value.Hour.ToString.PadLeft(2, "0")
    Dim End_Min = TimePicker_End.Value.Minute.ToString.PadLeft(2, "0")
    Dim End_Second = TimePicker_End.Value.Second.ToString.PadLeft(2, "0")
    Dim End_Time = End_Hour & ":" & End_Min & ":" & End_Second ' TimePicker_End.Value.TimeOfDay.ToString("hh:mm:ss")


    Dim Start_DateTime = Start_Date & " " & Start_Time
    Dim End_DateTime = End_Date & " " & End_Time

    Dim ret_Msg = ""
    Dim FileName_PDF = ""
    If Module_CreateBarCodeReport.O_Process_Message(PlatForm, LotNo, Start_DateTime, End_DateTime, FileName_PDF, ret_Msg) = False Then
      MsgBox(ret_Msg)
    Else
      MsgBox("檔案路徑" & vbCrLf & FileName_PDF)
    End If
  End Sub

  Private Sub TextBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tb_BarCodeInput.KeyPress
    If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then

      Dim index = CB_BarCode_SHOP.SelectedIndex
      If index = -1 Then
        MsgBox("請選擇賣場")
        Return
      End If
      Dim SHOP_NO = CB_BarCode_SHOP.Items(index)

      Select Case int_PlatForm
        Case 0  '7-11BarCode
        Case 1  '7-11QRCode
        Case 2  'OK
        Case 3  '全家
      End Select
      'Dim BarCode1 = ""
      'Dim BarCode2 = ""
      'Dim BarCode3 = ""
      'MsgBox("輸入的值是" & tb_BarCodeInput.Text)
      If Input_Cnt = 0 Then
        lb_BarCode1.Text = tb_BarCodeInput.Text.ToUpper
        BarCode1 = tb_BarCodeInput.Text.ToUpper
        If CheckBarcodeLength(BarCode1) = False Then
          MsgBox("條碼1長度過短")
          tb_BarCodeInput.Text = ""
          Return
        End If
        tb_BarCodeInput.Text = ""
        If int_PlatForm = 3 Or int_PlatForm = 1 Then
          Input_Cnt = 0
          Dim ret_MSG = ""
          If int_PlatForm = 3 Then
            If Module_ScanOKBarCode.O_Process_Message("Family", SHOP_NO, lb_BarCode1.Text, lb_BarCode2.Text, ret_MSG) = False Then
              MsgBox(ret_MSG)
              lb_BarCode1.Text = ""
              lb_BarCode2.Text = ""
              lb_BarCode3.Text = ""
              lb_BarCode4.Text = ""
            Else
              lb_BarCode1.Text = ""
              lb_BarCode2.Text = ""
              lb_BarCode3.Text = ""
              lb_BarCode4.Text = ""

            End If
          End If
          If int_PlatForm = 1 Then
            If Module_ScanOKBarCode.O_Process_Message("7-11QRCode", SHOP_NO, lb_BarCode1.Text, lb_BarCode2.Text, ret_MSG) = False Then
              MsgBox(ret_MSG)
              lb_BarCode1.Text = ""
              lb_BarCode2.Text = ""
              lb_BarCode3.Text = ""
              lb_BarCode4.Text = ""
            Else
              lb_BarCode1.Text = ""
              lb_BarCode2.Text = ""
              lb_BarCode3.Text = ""
              lb_BarCode4.Text = ""

            End If
          End If

        Else
            Input_Cnt = Input_Cnt + 1
        End If


      ElseIf Input_Cnt = 1 Then
        lb_BarCode2.Text = tb_BarCodeInput.Text.ToUpper
        BarCode2 = tb_BarCodeInput.Text.ToUpper

        If BarCode1 = BarCode2 Then
          MsgBox("條碼1與條碼2重復")
          tb_BarCodeInput.Text = ""
          Return
        ElseIf BarCode1.length > BarCode2.length Then
          MsgBox("條碼1與條碼2順序錯誤")
          tb_BarCodeInput.Text = ""
          Return
        End If
        If CheckBarcodeLength(BarCode2) = False Then
          MsgBox("條碼2長度過短")
          tb_BarCodeInput.Text = ""
          Return
        End If
        tb_BarCodeInput.Text = ""
        If int_PlatForm = 2 Then
          Input_Cnt = 0
          Dim ret_MSG = ""
          If Module_ScanOKBarCode.O_Process_Message("OK Mart", SHOP_NO, lb_BarCode1.Text, lb_BarCode2.Text, ret_MSG) = False Then
            MsgBox(ret_MSG)
            lb_BarCode1.Text = ""
            lb_BarCode2.Text = ""
            lb_BarCode3.Text = ""
            lb_BarCode4.Text = ""
          Else
            lb_BarCode1.Text = ""
            lb_BarCode2.Text = ""
            lb_BarCode3.Text = ""
            lb_BarCode4.Text = ""

          End If
        Else
          Input_Cnt = Input_Cnt + 1
        End If

      ElseIf Input_Cnt = 2 Then

        lb_BarCode3.Text = tb_BarCodeInput.Text.ToUpper
        BarCode3 = tb_BarCodeInput.Text.ToUpper
        If BarCode2 = BarCode3 Then
          MsgBox("條碼2與條碼3重復")
          tb_BarCodeInput.Text = ""
          Return
        ElseIf BarCode1 = BarCode3 Then
          MsgBox("條碼1與條碼3重復")
          tb_BarCodeInput.Text = ""
          Return
        ElseIf BarCode1.length > BarCode3.length Then
          MsgBox("條碼1與條碼3順序錯誤")
          tb_BarCodeInput.Text = ""
          Return
        End If
        If CheckBarcodeLength(BarCode3) = False Then
          MsgBox("條碼3長度過短")
          tb_BarCodeInput.Text = ""
          Return
        End If
        tb_BarCodeInput.Text = ""
        'Input_Cnt = 3
        Input_Cnt = 0

        Dim ret_MSG = ""
        If Module_Scan711BarCode.O_Process_Message("7-11BarCode", SHOP_NO, lb_BarCode1.Text, lb_BarCode2.Text, lb_BarCode3.Text, lb_BarCode4.Text, ret_MSG) = False Then
          MsgBox(ret_MSG)
          lb_BarCode1.Text = ""
          lb_BarCode2.Text = ""
          lb_BarCode3.Text = ""
          lb_BarCode4.Text = ""
        Else
          lb_BarCode1.Text = ""
          lb_BarCode2.Text = ""
          lb_BarCode3.Text = ""
          lb_BarCode4.Text = ""
        End If
      ElseIf Input_Cnt = 3 Then
        lb_BarCode4.Text = tb_BarCodeInput.Text.ToUpper
        tb_BarCodeInput.Text = ""
        Input_Cnt = 0

        Dim ret_MSG = ""
        If Module_Scan711BarCode.O_Process_Message("7-11", SHOP_NO, lb_BarCode1.Text, lb_BarCode2.Text, lb_BarCode3.Text, lb_BarCode4.Text, ret_MSG) = False Then
          MsgBox(ret_MSG)
          lb_BarCode1.Text = ""
          lb_BarCode2.Text = ""
          lb_BarCode3.Text = ""
          lb_BarCode4.Text = ""
        Else
          lb_BarCode1.Text = ""
          lb_BarCode2.Text = ""
          lb_BarCode3.Text = ""
          lb_BarCode4.Text = ""
        End If

      End If


    End If
  End Sub
  Private Function CheckBarcodeLength(ByVal Barcode As String) As Boolean

    If Barcode.Length < 5 Then
      Return False
    End If
    Return True
  End Function
  Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    Try
      int_PlatForm = ComboBox1.SelectedIndex


    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub CB_BarCode_SHOP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_BarCode_SHOP.SelectedIndexChanged
    Try
      Dim index = CB_BarCode_SHOP.SelectedIndex
      Dim SHOP_NO = CB_BarCode_SHOP.Items(index)

      Dim dicSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY(SHOP_NO)
      lb_Memo_Str.Text = dicSHOP.First.Value.MEMO


    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub CB_Report_SHOP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Report_SHOP.SelectedIndexChanged
    Try
      Dim index = CB_Report_SHOP.SelectedIndex
      Dim SHOP_NO = CB_Report_SHOP.Items(index)

      Dim dicSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY(SHOP_NO)
      lb_Report_Memo_Str.Text = dicSHOP.First.Value.MEMO


    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub lb_BarCode1_Click(sender As Object, e As EventArgs) Handles lb_BarCode1.Click
    Input_Cnt = 0
    lb_BarCode1.Text = ""
    lb_BarCode2.Text = ""
    lb_BarCode3.Text = ""
    lb_BarCode4.Text = ""

  End Sub

  Private Sub lb_BarCode2_Click(sender As Object, e As EventArgs) Handles lb_BarCode2.Click
    Input_Cnt = 1
    lb_BarCode2.Text = ""
    lb_BarCode3.Text = ""
    lb_BarCode4.Text = ""
  End Sub

  Private Sub lb_BarCode3_Click(sender As Object, e As EventArgs) Handles lb_BarCode3.Click
    Input_Cnt = 2
    lb_BarCode3.Text = ""
    lb_BarCode4.Text = ""
  End Sub

  Private Sub lb_BarCode4_Click(sender As Object, e As EventArgs) Handles lb_BarCode4.Click
    Input_Cnt = 3
    lb_BarCode4.Text = ""
  End Sub

  Private Sub ForTestToolStripMenuItem_Click(sender As Object, e As EventArgs)

  End Sub

  Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs)

  End Sub

  Private Sub TSCBViewLogLevel_Click(sender As Object, e As EventArgs)

  End Sub

  Private Sub 新增賣場ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 新增賣場ToolStripMenuItem.Click
    Try
      Dim _form = FormShop.CreateForm(FormShop.Name)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

  End Sub


End Class