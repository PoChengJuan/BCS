Imports System.IO
Imports System.Text
Imports System.Threading

Public Class ClsInterfaceDB
	Public StopGetCmd As Boolean = False '-True 會停止抓取MSG
	Public int_tGUIDBHandle As Integer = 0 '記錄執行緒的Count
	Private date_tGUIDBHandle As Date = Now '-計算時間用

	Private ThreadReceiveGUIMessage As Threading.Thread

	Sub New(ByRef logtool As eCALogTool._ILogTool, ByRef dbtool As eCA_DBTool.clsDBTool)
		WMS_T_GUI_CommandManagement.DBTool = dbtool
		Common_DBManagement.DBTool = dbtool

		ThreadReceiveGUIMessage = New Thread(New ThreadStart(AddressOf O_thr_GUIDBHandling))
		ThreadReceiveGUIMessage.IsBackground = True
		ThreadReceiveGUIMessage.Start()
	End Sub


	Public Sub O_thr_GUIDBHandling()
		Const SleepTime As Integer = 200
		Dim Count As Integer = 0
		While True
			Try
				date_tGUIDBHandle = Now
				If StopGetCmd Then
					SendMessageToLog("GetGUICmd Stop...", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
					While StopGetCmd
						Threading.Thread.Sleep(2000)
						If StopGetCmd = False Then Exit While
					End While
				End If
				If Count < 10 Then
					Count = Count + 1
				Else
					Count = 0
					If int_tGUIDBHandle > 99 Then
						int_tGUIDBHandle = 0
					Else
						int_tGUIDBHandle = int_tGUIDBHandle + 1
					End If
				End If
				I_ProcessFromGUIMessage()
			Catch ex As Exception
				SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Finally
				Threading.Thread.Sleep(SleepTime)
			End Try
		End While
		SendMessageToLog("thrGUIDBHandling End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
	End Sub
	Private Function I_ProcessFromGUIMessage() As Boolean
		Try
			SendMessageToLog(“Get GUI Message Start", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
			Dim dicFromGUICommand As Dictionary(Of String, clsFromGUICommand) = Nothing
			If IP.Length = 0 AndAlso CLIENT_ID.Length = 0 Then
				dicFromGUICommand = WMS_T_GUI_CommandManagement.GetCommandDictionaryByResultIsNULL_WaitUUIDIsNull(EnuSystemType.PrintTool)

			ElseIf IP.Length <> 0 AndAlso CLIENT_ID.Length = 0 Then
				dicFromGUICommand = WMS_T_GUI_CommandManagement.GetCommandDictionaryByReceiveSystem_IP_ResultIsNULL_WaitUUIDIsNull(EnuSystemType.PrintTool, IP)

			ElseIf IP.Length = 0 AndAlso CLIENT_ID.Length <> 0 Then
				dicFromGUICommand = WMS_T_GUI_CommandManagement.GetCommandDictionaryByReceiveSystem_ClentID_ResultIsNULL_WaitUUIDIsNull(EnuSystemType.PrintTool, CLIENT_ID)

			Else
				SendMessageToLog("賣亂阿拉   黑白設", eCALogTool.ILogTool.enuTrcLevel.lvError)
			End If


			'If IP.Length <> 0 Then ' IP or  CLIENT_ID 同時只會有一個有值
			'	dicFromGUICommand = WMS_T_GUI_CommandManagement.GetCommandDictionaryByReceiveSystem_IP_ResultIsNULL_WaitUUIDIsNull(EnuSystemType.PrintTool, IP)
			'ElseIf CLIENT_ID.Length <> 0 Then
			'	dicFromGUICommand = WMS_T_GUI_CommandManagement.GetCommandDictionaryByReceiveSystem_ClentID_ResultIsNULL_WaitUUIDIsNull(EnuSystemType.PrintTool, CLIENT_ID)

			'End If


			If dicFromGUICommand Is Nothing OrElse dicFromGUICommand.Count = 0 Then Return False
			SendMessageToLog(“Get GUI Message Final", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
			While dicFromGUICommand.Any = True
				Dim Function_ID As String = ""
				Dim XmlMessage As String = ""
				Dim Result As String = ""
				Dim ResultMessage As String = ""
				Dim Wait_UUID As String = ""
				Dim UUID As String = ""
				'用來暫存要處理的GUICommand
				Dim dicProcessFromGUICommand As New Dictionary(Of String, clsFromGUICommand)
				For Each objFromGUICommand As clsFromGUICommand In dicFromGUICommand.Values
					If UUID = "" Then
						UUID = objFromGUICommand.UUID
					End If
					'不相等表示是下一筆了，下一次再處理
					If UUID <> objFromGUICommand.UUID Then
						Exit For
					End If
					'儲存GUICommand的資訊
					If dicProcessFromGUICommand.ContainsKey(objFromGUICommand.gid) = False Then
						dicProcessFromGUICommand.Add(objFromGUICommand.gid, objFromGUICommand)
					Else
						SendMessageToLog(String.Format("GUICommand exist smae keys, UUID:{0}, Function_ID:{1}, SEQ ", objFromGUICommand.UUID, objFromGUICommand.Function_ID, objFromGUICommand.SEQ), eCALogTool.ILogTool.enuTrcLevel.lvWARN)
					End If
					Function_ID = objFromGUICommand.Function_ID
					XmlMessage = XmlMessage & objFromGUICommand.Message
				Next
				'把取得的GUICommand送出去進行處理
				If O_ProcessMessage(Function_ID, XmlMessage, ResultMessage, Wait_UUID) = True Then
					'執行成功
					'如果Wait_UUID不為空時才把Wait_UUID填入
					If Wait_UUID = "" Then
						Result = "0"
						'ResultMessage = ""
					End If
				Else
					'執行失敗
					Result = "1"
				End If
				'把GUICommand的執行結果寫入DB
				Dim lstSQL As New List(Of String)
				For Each objFromGUICommand As clsFromGUICommand In dicProcessFromGUICommand.Values
					objFromGUICommand.Result = Result
					objFromGUICommand.Result_Message = StrConv(ResultMessage, VbStrConv.TraditionalChinese, 2052) '简转繁

					objFromGUICommand.Wait_UUID = Wait_UUID
					SendMessageToLog("Set GUI Command Data Report, UUID=" & objFromGUICommand.UUID & ", Function_ID=" & objFromGUICommand.Function_ID & ", SEQ=" & objFromGUICommand.SEQ & ", Result=" & objFromGUICommand.Result & ", Result_Message=" & objFromGUICommand.Result_Message & ", Wait_UUID=" & objFromGUICommand.Wait_UUID, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
					If objFromGUICommand.O_Add_Update_SQLString(lstSQL) = False Then
						SendMessageToLog("Get SQL Faile, UUID=" & objFromGUICommand.UUID & ", Function_ID=" & objFromGUICommand.Function_ID & ", SEQ=" & objFromGUICommand.SEQ, eCALogTool.ILogTool.enuTrcLevel.lvError)
					End If
				Next
				If Common_DBManagement.BatchUpdate(lstSQL) = False Then
					SendMessageToLog("Update DB Failed, SQL=" & lstSQL.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
				End If
				'刪除已經處理過的GUICommand
				For Each objFromGUICommand As clsFromGUICommand In dicProcessFromGUICommand.Values
					If dicFromGUICommand.ContainsKey(objFromGUICommand.gid) = True Then
						dicFromGUICommand.Remove(objFromGUICommand.gid)
					End If
				Next
			End While
			SendMessageToLog(“Process GUI Message Finish", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
	Private Function O_ProcessMessage(ByVal strFunction_ID As String,
																	 ByVal strXmlMessage As String,
																	 ByRef ret_strResultMsg As String,
																	 Optional ByRef ret_strWait_UUID As String = "") As Boolean
		Try
			SendMessageToLog("ProcessCommand Function_ID=" & strFunction_ID & ", Message=" & strXmlMessage, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			Dim blnProcessResult As Boolean = False
			Select Case strFunction_ID
				Case EnuMSGFunctionID.T10F1S1_PrintCarrierLabel.ToString
					If O_ProcessT10F1S1_PrintCarrierLabel(strXmlMessage, ret_strResultMsg) Then
						blnProcessResult = True
					End If
        Case EnuMSGFunctionID.T10F1S2_PrintItemLabel.ToString
          If O_ProcessT10F1S2_PrintItemLabel(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case EnuMSGFunctionID.T10F1S21_PrintShippingDTL.ToString
          If O_ProcessT10F1S21_PrintShippingDTL(strXmlMessage, ret_strResultMsg) Then
            blnProcessResult = True
          End If
        Case Else
					SendMessageToLog("Case Error", eCALogTool.ILogTool.enuTrcLevel.lvError)

			End Select

			Return blnProcessResult
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
