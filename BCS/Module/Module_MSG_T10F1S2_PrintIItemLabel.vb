Imports eCA_HttpTool
Module Module_MSG_T10F1S2_PrintItemLabel

  Public Function O_ProcessT10F1S2_PrintItemLabel(ByVal strMessage As String,
								ByRef ret_strResult_Message As String,
								Optional ByRef ret_strWait_UUID As String = "") As Boolean
	Try
	  SendMessageToLog("Start Prcoess T10F1S2_PrintCarrierLabel", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
	  Dim PRINTER_NO = ""
	  '解Message的Xml
	  Dim Receive_Msg As MSG_T10F1S2_PrintItemLabel = Nothing
	  Receive_Msg = ParseXmlStringToClass(Of MSG_T10F1S2_PrintItemLabel)(strMessage, ret_strResult_Message)
	  If Receive_Msg Is Nothing Then
		SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
		ret_strResult_Message = "ParseXmlString To Class Failed"
		Return False
	  End If

	  'PRINTER_NAME
	  Dim ret_ResultMessage As String = Nothing
	  '10cm x 10cm 標籤紙
	  If PrintSetPrinter(ret_ResultMessage, PRINTER_NAME_10x10Label, "") = False Then
		ret_ResultMessage = "找不到印表機"
		SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
		Return False
	  Else
		ret_ResultMessage = "OK"
		SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
	  End If

	  '-將資料轉成列印
	  Dim PrintEnable = Receive_Msg.Body.PrintInfo.PRINT_ENABLE '是否列印 0:不列印/1:要列印
	  Dim To_PDF = Receive_Msg.Body.PrintInfo.TO_PDF '是否轉PDF 0:不轉/1:要轉
	  Dim PrintType = Receive_Msg.Body.PrintInfo.PRINT_TYPE '列印的類型 1:紙箱/2:物流箱
	  Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
	  Dim UUID = Receive_Msg.Header.UUID
	  Dim PrintFormatSettingPath = PrintFormatFile

			'開始組織欄位
			'Dim EXPORT_ID As String = GetNewTime_DBFormat()
			'Dim SAMPLE_FILE_NAME As String = ""
			'If PrintType = "1" Then
			'	SAMPLE_FILE_NAME = EnuPrintName.BoxBarcode.ToString
			'ElseIf PrintType = "2" Then
			'	SAMPLE_FILE_NAME = EnuPrintName.Logistics_BoxBarcode.ToString
			'Else
			'	ret_strResult_Message = "Error SAMPLE_FILE_NAME"
			'	Return False
			'End If
			'Dim CREATE_TIME As String = GetNewTime_DBFormat()
			'Dim FINISH_TIME As String = ""
			'Dim EXPORT_TYPE As Double = PrintEnable


			'Dim ret_dicAdd_EXPORT As New Dictionary(Of String, BCS.clsEXPORT)
			'Dim ret_dicAdd_EXPORT_DTL As New Dictionary(Of String, BCS.clsEXPORT_DTL)


			''建立列印資訊
			'Dim New_objExport As New BCS.clsEXPORT(EXPORT_ID, SAMPLE_FILE_NAME, CREATE_TIME, FINISH_TIME, EXPORT_TYPE, PRINTER_NO)
			'If ret_dicAdd_EXPORT.ContainsKey(New_objExport.gid) = False Then
			'	ret_dicAdd_EXPORT.Add(New_objExport.gid, New_objExport)
			'End If



			''組列印DTL
			'Dim New_objExport_DTL As BCS.clsEXPORT_DTL = Nothing
			'Dim TABLE_INDEX_1 As Double = 1 '第幾層 只目前支援三層(從1開始)
			'Dim TABLE_INDEX_2 As Double = 2 '第幾層 只目前支援三層(從1開始)
			'Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用		
			'Dim VALUE_INDEX As Double = 0 '該表中的第幾個位置(從1開始)
			'Dim VALUE As String = ""
			Dim ret_dicAdd_EXPORT As New Dictionary(Of String, BCS.clsEXPORT)
			Dim ret_dicAdd_EXPORT_DTL As New Dictionary(Of String, BCS.clsEXPORT_DTL)
			Dim EXPORT_ID As String = GetNewTime_DBFullTimeUUIDFormat()
			Dim SAMPLE_FILE_NAME As String = ""
			For Each fname As String In System.IO.Directory.GetFileSystemEntries(SamplePath)
				If System.IO.Path.GetFileNameWithoutExtension(fname) = "PackageinfoLabel_" & PrintType.ToString Then
					SAMPLE_FILE_NAME = System.IO.Path.GetFileNameWithoutExtension(fname)
					Exit For
				End If
			Next
			If SAMPLE_FILE_NAME = "" Then
				ret_strResult_Message = String.Format("Can not find sample file name from <{0}>,by file name <{1}>", SamplePath, "PackageinfoLabel_" & PrintType.ToString)
				SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
				Return False
			End If
			Dim CREATE_TIME As String = GetNewTime_DBFormat()
			Dim FINISH_TIME As String = ""
			Dim EXPORT_TYPE As Double = PrintEnable
			'建立列印資訊
			Dim New_objExport As New BCS.clsEXPORT(EXPORT_ID, SAMPLE_FILE_NAME, CREATE_TIME, FINISH_TIME, EXPORT_TYPE, PRINTER_NO)
			If ret_dicAdd_EXPORT.ContainsKey(New_objExport.gid) = False Then
				ret_dicAdd_EXPORT.Add(New_objExport.gid, New_objExport)
			End If
			Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用	

			Dim Print_Count = 0
			Dim Serial_No = 1
			Dim Total_Coount = Receive_Msg.Body.LabelList.LabelInfo.Count
			For Each objItemLabel In Receive_Msg.Body.LabelList.LabelInfo

				'組列印DTL
				Dim New_objExport_DTL As BCS.clsEXPORT_DTL = Nothing
				Dim TABLE_INDEX_1 As Double = 1 '第幾層 只目前支援三層(從1開始)
				Dim TABLE_INDEX_2 As Double = 2 '第幾層 只目前支援三層(從1開始)
				'Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用		
				Dim VALUE_INDEX As Double = 0 '該表中的第幾個位置(從1開始)
				Dim VALUE As String = ""
				Dim strUpLoad_ID As String = ""
				Select Case UpLoadKey.ToLower.Replace("tag", "")
					Case "1"
						strUpLoad_ID = objItemLabel.TAG1
					Case "2"
						strUpLoad_ID = objItemLabel.TAG2
					Case "3"
						strUpLoad_ID = objItemLabel.TAG3
					Case "4"
						strUpLoad_ID = objItemLabel.TAG4
					Case "5"
						strUpLoad_ID = objItemLabel.TAG5
					Case "6"
						strUpLoad_ID = objItemLabel.TAG6
					Case "7"
						strUpLoad_ID = objItemLabel.TAG7
					Case "8"
						strUpLoad_ID = objItemLabel.TAG8
					Case "9"
						strUpLoad_ID = objItemLabel.TAG9
					Case "10"
						strUpLoad_ID = objItemLabel.TAG10
					Case "11"
						strUpLoad_ID = objItemLabel.TAG11
					Case "12"
						strUpLoad_ID = objItemLabel.TAG12
					Case "13"
						strUpLoad_ID = objItemLabel.TAG13
					Case "14"
						strUpLoad_ID = objItemLabel.TAG14
					Case "15"
						strUpLoad_ID = objItemLabel.TAG15
					Case "16"
						strUpLoad_ID = objItemLabel.TAG16
					Case "17"
						strUpLoad_ID = objItemLabel.TAG17
					Case "18"
						strUpLoad_ID = objItemLabel.TAG18
					Case "19"
						strUpLoad_ID = objItemLabel.TAG19
					Case "20"
						strUpLoad_ID = objItemLabel.TAG20
					Case "21"
						strUpLoad_ID = objItemLabel.TAG21
					Case "22"
						strUpLoad_ID = objItemLabel.TAG22
					Case "23"
						strUpLoad_ID = objItemLabel.TAG23
					Case "24"
						strUpLoad_ID = objItemLabel.TAG24
					Case "25"
						strUpLoad_ID = objItemLabel.TAG25
					Case "26"
						strUpLoad_ID = objItemLabel.TAG26
					Case "27"
						strUpLoad_ID = objItemLabel.TAG27
					Case "28"
						strUpLoad_ID = objItemLabel.TAG28
					Case "29"
						strUpLoad_ID = objItemLabel.TAG29
					Case "30"
						strUpLoad_ID = objItemLabel.TAG30
					Case "31"
						strUpLoad_ID = objItemLabel.TAG31
					Case "32"
						strUpLoad_ID = objItemLabel.TAG32
					Case "33"
						strUpLoad_ID = objItemLabel.TAG33
					Case "34"
						strUpLoad_ID = objItemLabel.TAG34
					Case "35"
						strUpLoad_ID = objItemLabel.TAG35
				End Select

				'物料標籤,固定撈取35個Tag做處理
				For i = 1 To 35
					VALUE = "" '初始化
					VALUE_INDEX += 1
					Select Case i
						Case 1
							VALUE = objItemLabel.TAG1
						Case 2
							VALUE = objItemLabel.TAG2
						Case 3
							VALUE = objItemLabel.TAG3
						Case 4
							VALUE = objItemLabel.TAG4
						Case 5
							VALUE = objItemLabel.TAG5
						Case 6
							VALUE = objItemLabel.TAG6
						Case 7
							VALUE = objItemLabel.TAG7
						Case 8
							VALUE = objItemLabel.TAG8
						Case 9
							VALUE = objItemLabel.TAG9
						Case 10
							VALUE = objItemLabel.TAG10
						Case 11
							VALUE = objItemLabel.TAG11
						Case 12
							VALUE = objItemLabel.TAG12
						Case 13
							VALUE = objItemLabel.TAG13
						Case 14
							VALUE = objItemLabel.TAG14
						Case 15
							VALUE = objItemLabel.TAG15
						Case 16
							VALUE = objItemLabel.TAG16
						Case 17
							VALUE = objItemLabel.TAG17
						Case 18
							VALUE = objItemLabel.TAG18
						Case 19
							VALUE = objItemLabel.TAG19
						Case 20
							VALUE = objItemLabel.TAG20
						Case 21
							VALUE = objItemLabel.TAG21
						Case 22
							VALUE = objItemLabel.TAG22
						Case 23
							VALUE = objItemLabel.TAG23
						Case 24
							VALUE = objItemLabel.TAG24
						Case 25
							VALUE = objItemLabel.TAG25
						Case 26
							VALUE = objItemLabel.TAG26
						Case 27
							VALUE = objItemLabel.TAG27
						Case 28
							VALUE = objItemLabel.TAG28
						Case 29
							VALUE = objItemLabel.TAG29
						Case 30
							VALUE = objItemLabel.TAG30
						Case 31
							VALUE = objItemLabel.TAG31
						Case 32
							VALUE = objItemLabel.TAG32
						Case 33
							VALUE = objItemLabel.TAG33
						Case 34
							VALUE = objItemLabel.TAG34
						Case 35
							VALUE = objItemLabel.TAG35
					End Select
					New_objExport_DTL = New BCS.clsEXPORT_DTL(EXPORT_ID, TABLE_INDEX_1, SHEET_INDEX_1, VALUE_INDEX, VALUE)
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
					If To_PDF = "1" Then '-要轉PDF
						If PrintEnable = "1" Then '1:要列印
							If BCS.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, Serial_No.ToString, ret_strResult_Message) = False Then
								ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
								Return False
							End If

						Else '-不要列印-會只存成PDF文件
							If BCS.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, Serial_No.ToString, ret_strResult_Message) = False Then
								ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
								Return False
							End If
							'如果處理正確 則回寫完整擋案路徑在ResultMsg
							If ret_strResult_Message = "" Then
								ret_strResult_Message = ret_reportfilename
							End If
							'Return True
						End If
						If HTTPPath.Length > 0 AndAlso ret_reportfilename.Length > 0 Then

							Dim dicQueryParam As New Dictionary(Of String, String)
							Dim lstfilepath As List(Of String) = ret_reportfilename.Split("\").ToList
							Dim strFileName As String = lstfilepath(lstfilepath.Count - 1)
							Dim intSubIndex As Integer = lstfilepath(lstfilepath.Count - 1).Length
							Dim strfilepath As String = ret_reportfilename.Substring(0, ret_reportfilename.Length - intSubIndex - 1)
							lstfilepath.Clear()
							strUpLoad_ID = IIf(strUpLoad_ID = "", Now.ToString("yyyyMMddHHmmss") & "_" & Serial_No, strUpLoad_ID)
							dicQueryParam.Add("FileDir", strUpLoad_ID)
							dicQueryParam.Add("FILE_ID", strUpLoad_ID)
							dicQueryParam.Add("FILE_TYPE", "LABEL")
							dicQueryParam.Add("FILE_PATH", Now.ToString("yyyyMMdd") & "/" & strUpLoad_ID & "/" & strFileName)
							dicQueryParam.Add("FILE_NAME", strFileName)
							dicQueryParam.Add("FILE_DESC", strFileName)
							dicQueryParam.Add("FILE_COMMON1", strUpLoad_ID)
							dicQueryParam.Add("FILE_COMMON2", "")
							dicQueryParam.Add("FILE_COMMON3", "")
							dicQueryParam.Add("FILE_COMMON4", "")
							dicQueryParam.Add("FILE_COMMON5", "")
							dicQueryParam.Add("RELATED_PO_ID", "")
							dicQueryParam.Add("USER_ID", "PrintTool")
							ret_strResult_Message = O_HttpUploadFile(HTTPPath, dicQueryParam, strFileName, ret_reportfilename, StructureContentType.PDF)
							dicQueryParam.Clear()
						End If

					Else '-不要轉PDF
						If PrintEnable = "1" Then '1:要列印
							If BCS.PrintStart(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, ret_strResult_Message, False) = False Then
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
		  End If
		  ret_dicAdd_EXPORT_DTL.Clear()
		  Serial_No += 1
		  Print_Count = 0
		End If
		Threading.Thread.Sleep(1000)
	  Next

	  Return True
	Catch ex As Exception
	  ret_strResult_Message = ex.Message
	  SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
	  Return False
	End Try
  End Function

End Module
