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

      TSCBViewLogLevel.SelectedIndex = LogTool.ViewLV - 1

      NotifyIcon = New System.Windows.Forms.NotifyIcon()
      NotifyIcon1.Text = My.Application.Info.AssemblyName

      SendMessageToLog("[PrintTool Start]", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)

      Handling() '-根據Config 確認這次的interface

      Timer1.Interval = RefreshAliveTime
      Timer1.Enabled = True



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

  Private Sub T10F1S1PrintCarrierLabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T10F1S1PrintCarrierLabelToolStripMenuItem.Click
    Try
      Dim _form = FrmMSGTest.CreateForm(EnuMSGFunctionID.T10F1S1_PrintCarrierLabel)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
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

  Private Sub T10F1S2PrintItemLabelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T10F1S2PrintItemLabelToolStripMenuItem.Click
    Try
      Dim _form = FrmMSGTest.CreateForm(EnuMSGFunctionID.T10F1S2_PrintItemLabel)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  Private Sub T10F1S21PrintShippingDTLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles T10F1S21PrintShippingDTLToolStripMenuItem.Click
    Try

      Dim _form = FrmMSGTest.CreateForm(EnuMSGFunctionID.T10F1S21_PrintShippingDTL)
      If _form IsNot Nothing Then
        _form.Show()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub


  Private Sub TabControl_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
    If e.TabPage.TabIndex = TabPage1.TabIndex Then
      tb_LotNo.Select()
    ElseIf e.TabPage.TabIndex = TabPage3.TabIndex Then
      TextBox2.Select()
    End If

  End Sub

  Private Sub btn_START_Click(sender As Object, e As EventArgs) Handles btn_START.Click
    Try
      If flg_Start = False Then
        flg_Start = True
      ElseIf flg_Start = True Then
        flg_Start = False
      End If

      If flg_Start = False Then
        btn_START.Text = "開始刷取"
      ElseIf flg_Start = True Then
        btn_START.Text = "停止刷取"

        '初始COM PORT


        RS232 = New SerialPort("COM10", 9600, Parity.None, 8, 1)

        If (Not RS232.IsOpen) Then

          RS232.Open()

        End If

        Dim td As Thread = New Thread(AddressOf serialPort1_DataReceived)

        td.Start()
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Sub serialPort1_DataReceived()
    While True

      If RS232.BytesToRead > 0 Then
        strIncoming = RS232.ReadExisting.ToString

        RS232.DiscardInBuffer()

        Me.Invoke(New EventHandler(AddressOf ForDisplay))         '呼叫接收資料函式
      End If
    End While
  End Sub
  Public Sub ForDisplay()

    lb_BarCode1.Text = strIncoming

    ListBox1.Items.Add(strIncoming)       '取到回車符位置數-1

    ListBox1.SelectedIndex = ListBox1.Items.Count - 1

    If ListBox1.Items.Count > 1000 Then

      Me.ListBox1.Items.Clear()

    End If

  End Sub
End Class