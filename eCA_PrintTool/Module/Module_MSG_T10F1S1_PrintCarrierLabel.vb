Imports System.Drawing
Imports GenCode128
Imports System.Diagnostics
Module Module_MSG_T10F1S1_PrintCarrierLabel

  Public Function O_ProcessT10F1S1_PrintCarrierLabel(ByVal strMessage As String,
								ByRef ret_strResult_Message As String,
								Optional ByRef ret_strWait_UUID As String = "") As Boolean
	Try
	  SendMessageToLog("Start Prcoess T10F1S1_PrintCarrierLabel", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
	  Dim PRINTER_NO = ""
	  '解Message的Xml
	  Dim Receive_Msg As MSG_T10F1S1_PrintCarrierLabel = Nothing
	  Receive_Msg = ParseXmlStringToClass(Of MSG_T10F1S1_PrintCarrierLabel)(strMessage, ret_strResult_Message)
	  If Receive_Msg Is Nothing Then
		SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
		ret_strResult_Message = "ParseXmlString To Class Failed"
		Return False
	  End If

	  ''檢查Adobe Reader有沒有開，沒開就打開
	  Dim processList() As Process
	  Dim strPDFProcess As String = strPDFPath.Substring(strPDFPath.LastIndexOf("\") + 1, strPDFPath.Length - strPDFPath.LastIndexOf("\") - 1).Replace(".exe", "")
	  processList = Process.GetProcessesByName(strPDFProcess)
	  If processList.Length = 0 Then
		Shell(strPDFPath, vbNormalFocus)
	  End If

	  Dim myprocess As New Process
	  Dim StartInfo As New System.Diagnostics.ProcessStartInfo
	  StartInfo = New System.Diagnostics.ProcessStartInfo(strPDFPath.Substring(strPDFPath.LastIndexOf("\") + 1, strPDFPath.Length - strPDFPath.LastIndexOf("\") - 1))
	  StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized
	  myprocess.StartInfo = StartInfo
	  myprocess.Start()
	  'PRINTER_NAME
	  Dim ret_ResultMessage As String = Nothing
	  '10cm x 10cm 標籤紙
	  If PrintSetPrinter(ret_ResultMessage, PRINTER_NAME_10x10Label, "") = False Then
		ret_ResultMessage = "找不到印表機"
		SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
		Return False
	  Else
		ret_ResultMessage = "OK"
		SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
	  End If

	  '-將資料轉成列印
	  Dim PrintEnable = Receive_Msg.Body.PrintInfo.PRINT_ENABLE '是否列印 0:不列印/1:要列印
	  Dim To_PDF = Receive_Msg.Body.PrintInfo.TO_PDF '是否轉PDF 0:不轉/1:要轉
	  Dim PrintType = Receive_Msg.Body.PrintInfo.PRINT_TYPE '列印的類型 1:紙箱/2:物流箱
	  Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
	  Dim UUID = Receive_Msg.Header.UUID
	  Dim PrintFormatSettingPath = PrintFormatFile


	  Dim ret_dicAdd_EXPORT As New Dictionary(Of String, eCA_PrintTool.clsEXPORT)
	  Dim ret_dicAdd_EXPORT_DTL As New Dictionary(Of String, eCA_PrintTool.clsEXPORT_DTL)
	  Dim EXPORT_ID As String = GetNewTime_DBFullTimeUUIDFormat()
	  Dim SAMPLE_FILE_NAME As String = ""
	  For Each fname As String In System.IO.Directory.GetFileSystemEntries(SamplePath)
		If System.IO.Path.GetFileNameWithoutExtension(fname) = "CSTinfoLabel_" & PrintType.ToString Then
		  SAMPLE_FILE_NAME = System.IO.Path.GetFileNameWithoutExtension(fname)
		  Exit For
		End If
	  Next
	  If SAMPLE_FILE_NAME = "" Then
		ret_strResult_Message = String.Format("Can not find sample file name from <{0}>,by file name <{1}>", SamplePath, "CSTinfoLabel_" & PrintType.ToString)
		SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
		Return False
	  End If

	  Dim CREATE_TIME As String = GetNewTime_DBFormat()
	  Dim FINISH_TIME As String = ""
	  Dim EXPORT_TYPE As Double = PrintEnable
	  '建立列印資訊
	  Dim New_objExport As New eCA_PrintTool.clsEXPORT(EXPORT_ID, SAMPLE_FILE_NAME, CREATE_TIME, FINISH_TIME, EXPORT_TYPE, PRINTER_NO)
	  If ret_dicAdd_EXPORT.ContainsKey(New_objExport.gid) = False Then
		ret_dicAdd_EXPORT.Add(New_objExport.gid, New_objExport)
	  End If
	  Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用	

	  Dim Print_Count = 0
	  Dim Serial_No = 1
	  Dim Total_Coount = Receive_Msg.Body.LabelList.LabelInfo.Count
	  For Each objCarrierLabel In Receive_Msg.Body.LabelList.LabelInfo
		'組列印DTL
		Dim New_objExport_DTL As eCA_PrintTool.clsEXPORT_DTL = Nothing
		Dim TABLE_INDEX_1 As Double = 1 '第幾層 只目前支援三層(從1開始)
		Dim TABLE_INDEX_2 As Double = 2 '第幾層 只目前支援三層(從1開始)
		'Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用		
		Dim VALUE_INDEX As Double = 0 '該表中的第幾個位置(從1開始)
		Dim VALUE As String = ""

		For i = 1 To 25
		  VALUE = "" '初始化
		  VALUE_INDEX += 1
		  Select Case i
			Case 1
			  VALUE = objCarrierLabel.TAG1
			Case 2
			  VALUE = objCarrierLabel.TAG2
			Case 3
			  VALUE = objCarrierLabel.TAG3
			Case 4
			  VALUE = objCarrierLabel.TAG4
			Case 5
			  VALUE = objCarrierLabel.TAG5
			Case 6
			  VALUE = objCarrierLabel.TAG6
			Case 7
			  VALUE = objCarrierLabel.TAG7
			Case 8
			  VALUE = objCarrierLabel.TAG8
			Case 9
			  VALUE = objCarrierLabel.TAG9
			Case 10
			  VALUE = objCarrierLabel.TAG10
			Case 11
			  VALUE = objCarrierLabel.TAG11
			Case 12
			  VALUE = objCarrierLabel.TAG12
			Case 13
			  VALUE = objCarrierLabel.TAG13
			Case 14
			  VALUE = objCarrierLabel.TAG14
			Case 15
			  VALUE = objCarrierLabel.TAG15
			Case 16
			  VALUE = objCarrierLabel.TAG16
			Case 17
			  VALUE = objCarrierLabel.TAG17
			Case 18
			  VALUE = objCarrierLabel.TAG18
			Case 19
			  VALUE = objCarrierLabel.TAG19
			Case 20
			  VALUE = objCarrierLabel.TAG20
			Case 21
			  VALUE = objCarrierLabel.TAG21
			Case 22
			  VALUE = objCarrierLabel.TAG22
			Case 23
			  VALUE = objCarrierLabel.TAG23
			Case 24
			  VALUE = objCarrierLabel.TAG24
			Case 25
			  VALUE = objCarrierLabel.TAG25
		  End Select
		  New_objExport_DTL = New eCA_PrintTool.clsEXPORT_DTL(EXPORT_ID, TABLE_INDEX_1, SHEET_INDEX_1, VALUE_INDEX, VALUE)
		  If ret_dicAdd_EXPORT_DTL.ContainsKey(New_objExport_DTL.gid) = False Then
			ret_dicAdd_EXPORT_DTL.Add(New_objExport_DTL.gid, New_objExport_DTL)
		  End If
		Next


		Print_Count += 1
		Total_Coount -= 1
		SHEET_INDEX_1 += 1

		'設定開始列印的標籤數
		Dim Start_Print = 200
		Start_Print = CInt(PRINTER_LIMITED_COUNT)
		'資料收集量超過設定的標籤數，或已收集完所有標籤數即開始列印
		If Print_Count >= Start_Print Or Total_Coount = 0 Then
		  '-處理方式
		  Dim ret_reportfilename = ""
		  'To_PDF = 1
		  If To_PDF = "1" Then '-要轉PDF
			If PrintEnable = "1" Then '1:要列印
			  If eCA_PrintTool.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, Serial_No.ToString, ret_strResult_Message) = False Then
				ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
				Return False
			  End If

			Else '-不要列印-會只存成PDF文件
			  If eCA_PrintTool.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, Serial_No.ToString, ret_strResult_Message) = False Then
				ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
				Return False
			  End If
			  '如果處理正確 則回寫完整擋案路徑在ResultMsg
			  If ret_strResult_Message = "" Then
				ret_strResult_Message = ret_reportfilename

				'PrintPDF(ret_reportfilename, ret_strResult_Message)
			  End If
			  'Return True
			End If

			Threading.Thread.Sleep(1000)
		  Else '-不要轉PDF
			If PrintEnable = "1" Then '1:要列印
			  If eCA_PrintTool.PrintStart(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, ret_strResult_Message, False) = False Then
				ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
				Return False
			  Else
				'Return True
			  End If

			Else '-不要列印
			  ret_strResult_Message = "不列印不轉PDF 是來亂的嗎?"
			  SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
			  'Return True
			End If
			Threading.Thread.Sleep(1000)
		  End If
		  Print_Count = 0
		  SHEET_INDEX_1 = 0
		  Serial_No += 1
		  ret_dicAdd_EXPORT_DTL.Clear()
		End If

	  Next
	  SendMessageToLog("T10F1S1 End", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
	  Return True
	Catch ex As Exception
	  ret_strResult_Message = ex.Message
	  SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
	  Return False
	End Try
  End Function

End Module
