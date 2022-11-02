Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.IO.Ports

Public Module ModuleDeclaration
  Public ConfigPath As String = Application.StartupPath & "\PrintConfig.xml"
  Public LogTool As eCALogTool._ILogTool
  Public LogPath As String
  Public InterfaceType As String '1:DB 2:RabbitMQ
  Public CLIENT_ID As String = ""
  Public IP As String = ""
  Public ExportPath As String
  Public SamplePath As String
  Public PrintFormatFile As String
  Public HTTPPath As String = ""
  Public UpLoadKey As String = ""
  Public strPDFPath As String = ""
  Public gFileRootPath As String = ""
  Public gReportFunction As String = ""
  Public gMonthReportTime As String = ""
  'DateTime
  Public Const DBFullTimeUUIDFormat As String = "yyyyMMddHHmmssfff"
  Public Const DBFullTimeFormat As String = "yyyy/MM/dd HH:mm:ss.fff"
  Public Const DBTimeFormat As String = "yyyy/MM/dd HH:mm:ss"
  Public Const DBDateFormat As String = "yyyy/MM/dd"
  Public Const DBMonthFormat As String = "yyyy/MM"
  Public Const DBDate_IDFormat As String = "yyMMdd"
  Public Const DBOnlyTimeFormat As String = "HH:mm:ss"




  'WCF Host IP:Port-目前尚未使用到
  Public WCFHostIpPort As String = "127.0.0.1:10001"
  Public MCSRpcURL As String
  Public AmatMCSRpcURL As String
  Public DBTool_Name As String 'DBTool记录挡案资料夹

  Public PRINTER_NAME_A4 As String
  Public PRINTER_NAME_10x10Label As String
  Public PRINTER_LIMITED_COUNT As String
  Public PRINTER_Label_Limited As String = "0"
  Public blnPageTop As Boolean = False
  Public Const LinkKey As String = "*"

  Public flg_Start As Boolean = False
  Public strIncoming As String

  Public RS232 As SerialPort

  Public Input_Cnt = 0

  '掃描頁面變數
  Public int_PlatForm = 0

  'Public PRINTER_NAME_CarrierLabel As String


  Public Function SendMessageToLog(ByVal message As String, ByVal messageLevel As eCALogTool.ILogTool.enuTrcLevel, Optional frameNum As Integer = 2) As Boolean
    Try
      If LogTool IsNot Nothing Then
        LogTool.TraceLog(String.Format("Message:{0}", message), , (New System.Diagnostics.StackTrace).GetFrame(frameNum).GetMethod.Name, messageLevel)
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  Public Function ModifyStringApostrophe(ByVal name As String) As String
    Try
      name = name.Replace("'", "''")
      name = name.Replace(" & ", "'||'&'||'")
      Return name

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Function ParseXmlStringToClass(Of T)(_XmlString As String, ByRef RetMsg As String) As T
    Try
      Return New XmlSerializer(GetType(T)).Deserialize(New MemoryStream(Encoding.UTF8.GetBytes(_XmlString)))
    Catch ex As Exception
      RetMsg = ex.ToString
      Return Nothing
    End Try
  End Function
  Public Function GetNewTime_DBFullTimeUUIDFormat() As String
    Try
      Return Now.ToString(DBFullTimeUUIDFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function GetNewTime_DBFormat() As String
    Try
      Return Now.ToString(DBTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function GetNewDate_DBFormat() As String
    Try
      Return Now.ToString(DBDate_IDFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function



  '-條碼格式轉換 
  Public Function Code128b(DataToEncode As String) As String
    Dim endchar As String

    Dim total As Integer
    total = 104
    Dim tmp As Integer
    Dim i As Integer
    Dim Ito As Integer
    Ito = Len(DataToEncode)
    Dim tempstring As String
    Dim j As Integer

    For i = 1 To Ito
      tempstring = Mid(DataToEncode, i, 1)
      tmp = AscW(tempstring)

      If tmp >= 32 Then
        total = total + (tmp - 32) * i
      Else
        total = total + (tmp + 64) * i
      End If
    Next
    Dim endasc As Integer

    endasc = total Mod 103

    If endasc >= 95 Then
      Select Case endasc
        Case 95
          endchar = ChrW(195)
        Case 96
          endchar = ChrW(196)
        Case 97
          endchar = ChrW(197)
        Case 98
          endchar = ChrW(198)
        Case 99
          endchar = ChrW(199)
        Case 100
          endchar = ChrW(200)
        Case 101
          endchar = ChrW(201)
        Case 102
          endchar = ChrW(202)
      End Select
    Else
      endasc = endasc + 32
      endchar = ChrW(endasc)
    End If
    Return ChrW(204) + DataToEncode + endchar + ChrW(206)
  End Function
End Module
