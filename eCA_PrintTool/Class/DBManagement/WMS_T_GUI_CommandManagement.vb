Public Class WMS_T_GUI_CommandManagement
	Public Shared TableName As String = "WMS_T_GUI_Command"
	Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

	Enum IdxColumnName As Integer
		UUID
		SEND_SYSTEM
		RECEIVE_SYSTEM
		FUNCTION_ID
		SEQ
		USER_ID
		CREATE_TIME
		MESSAGE
		RESULT
		RESULT_MESSAGE
		WAIT_UUID
		CLIENT_ID
		IP
	End Enum
	Public Shared Function GetInsertSQL(ByRef Info As clsFromGUICommand) As String
		Try
			Dim strSQL As String = ""
			strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14},{16},{18},{20},{22},{24},{26}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}','{17}','{19}','{21}','{23}','{25}','{27}')",
				strSQL,
				TableName,
				IdxColumnName.UUID.ToString, Info.UUID,
				IdxColumnName.SEND_SYSTEM.ToString, CInt(Info.Send_System),
				IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.Receive_System),
				IdxColumnName.FUNCTION_ID.ToString, Info.Function_ID,
				IdxColumnName.SEQ.ToString, Info.SEQ,
				IdxColumnName.USER_ID.ToString, Info.User_ID,
				IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
				IdxColumnName.MESSAGE.ToString, Info.Message,
				IdxColumnName.RESULT.ToString, Info.Result,
				IdxColumnName.RESULT_MESSAGE.ToString, Info.Result_Message,
				IdxColumnName.WAIT_UUID.ToString, Info.Wait_UUID,
														 IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
														 IdxColumnName.IP.ToString, Info.IP
			)
			Dim NewSQL As String = ""
			If SQLCorrect(strSQL, NewSQL) Then
				Return NewSQL
			End If
			Return Nothing
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function GetUpdateSQL(ByRef Info As clsFromGUICommand) As String
		Try
			Dim strSQL As String = ""
			strSQL = String.Format("Update {1} SET {18}='{19}',{20}='{21}',{22}='{23}',{24}='{25}',{26}='{27}' WHERE {2}='{3}' AND {8}='{9}' AND {10}='{11}' ",
				strSQL,
				TableName,
				IdxColumnName.UUID.ToString, Info.UUID,
				IdxColumnName.SEND_SYSTEM.ToString, CInt(Info.Send_System),
				IdxColumnName.RECEIVE_SYSTEM.ToString, CInt(Info.Receive_System),
				IdxColumnName.FUNCTION_ID.ToString, Info.Function_ID,
				IdxColumnName.SEQ.ToString, Info.SEQ,
				IdxColumnName.USER_ID.ToString, Info.User_ID,
				IdxColumnName.CREATE_TIME.ToString, Info.Create_Time,
				IdxColumnName.MESSAGE.ToString, Info.Message,
				IdxColumnName.RESULT.ToString, Info.Result,
				IdxColumnName.RESULT_MESSAGE.ToString, Info.Result_Message,
				IdxColumnName.WAIT_UUID.ToString, Info.Wait_UUID,
														 IdxColumnName.CLIENT_ID.ToString, Info.CLIENT_ID,
														 IdxColumnName.IP.ToString, Info.IP
			)
			Return strSQL
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function GetDeleteSQL(ByRef Info As clsFromGUICommand) As String
		Try
			Dim strSQL As String = ""
			strSQL = String.Format("Delete FROM {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}' ",
				strSQL,
				TableName,
				IdxColumnName.UUID.ToString, Info.UUID,
				IdxColumnName.FUNCTION_ID.ToString, Info.Function_ID,
				IdxColumnName.SEQ.ToString, Info.SEQ
			)
			Return strSQL
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function GetCommandDictionaryByResultIsNULL_WaitUUIDIsNull(ByVal ReceiveSystem As EnuSystemType) As Dictionary(Of String, clsFromGUICommand)

		Try
			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				strSQL = String.Format("Select * FROM {0} WHERE ({1} IS NULL or {1} ='' ) AND ({2} IS NULL or {2} = '' ) AND  {3}={4}  ORDER BY {5} ASC, {6} ASC, {7} ASC ",
																	 TableName,
																	 IdxColumnName.RESULT.ToString,
																	 IdxColumnName.WAIT_UUID.ToString,
																	 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
																	 IdxColumnName.CREATE_TIME.ToString,
																	 IdxColumnName.UUID.ToString,
																	 IdxColumnName.SEQ.ToString)
				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function GetCommandDictionaryByReceiveSystem_IP_ResultIsNULL_WaitUUIDIsNull(ByVal ReceiveSystem As EnuSystemType, ByVal IP As String) As Dictionary(Of String, clsFromGUICommand)

		Try
			'IP有可能是複數的
			Dim _SplictIP = IP.Split(",")
			Dim _IPValue = ""
			If _SplictIP.Count = 1 Then
				_IPValue = IP
			Else
				For Each item In _SplictIP
					_IPValue = _IPValue & "'" & item & "'" & ","
				Next
				_IPValue = _IPValue.TrimEnd(",")
			End If
			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				If _SplictIP.Count = 1 Then
					strSQL = String.Format("Select * FROM {0} WHERE ({1} IS NULL or {1} ='' ) AND ({2} IS NULL or {2} = '' ) AND  {3}={4} and {8}='{9}' ORDER BY {5} ASC, {6} ASC, {7} ASC ",
																 TableName,
																 IdxColumnName.RESULT.ToString,
																 IdxColumnName.WAIT_UUID.ToString,
																 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
																 IdxColumnName.CREATE_TIME.ToString,
																 IdxColumnName.UUID.ToString,
																 IdxColumnName.SEQ.ToString,
														 IdxColumnName.IP, _IPValue)
				Else
					strSQL = String.Format("Select * FROM {0} WHERE ({1} IS NULL or {1} ='' ) AND ({2} IS NULL or {2} = '' ) AND  {3}={4} and {8} in ({9}) ORDER BY {5} ASC, {6} ASC, {7} ASC ",
															 TableName,
															 IdxColumnName.RESULT.ToString,
															 IdxColumnName.WAIT_UUID.ToString,
															 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
															 IdxColumnName.CREATE_TIME.ToString,
															 IdxColumnName.UUID.ToString,
															 IdxColumnName.SEQ.ToString,
													 IdxColumnName.IP, _IPValue)

				End If

				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	Public Shared Function GetCommandDictionaryByReceiveSystem_ClentID_ResultIsNULL_WaitUUIDIsNull(ByVal ReceiveSystem As EnuSystemType, ByVal clientid As String) As Dictionary(Of String, clsFromGUICommand)

		Try
			'ClientID有可能是複數的
			Dim _Splictclientid = clientid.Split(",")
			Dim _clientidValue = ""
			If _Splictclientid.Count = 1 Then
				_clientidValue = clientid
			Else
				For Each item In _Splictclientid
					_clientidValue = _clientidValue & "'" & item & "'" & ","
				Next
				_clientidValue = _clientidValue.TrimEnd(",")
			End If

			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				If _Splictclientid.Count = 1 Then
					strSQL = String.Format("Select * FROM {0} WHERE {1} IS NULL AND {2} IS NULL AND  {3}={4} and {8}='{9}' ORDER BY {5} ASC, {6} ASC, {7} ASC ",
																	 TableName,
																	 IdxColumnName.RESULT.ToString,
																	 IdxColumnName.WAIT_UUID.ToString,
																	 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
																	 IdxColumnName.CREATE_TIME.ToString,
																	 IdxColumnName.UUID.ToString,
																	 IdxColumnName.SEQ.ToString,
															 IdxColumnName.CLIENT_ID, _clientidValue)
				Else
					strSQL = String.Format("Select * FROM {0} WHERE {1} IS NULL AND {2} IS NULL AND  {3}={4} and {8} in ({9}) ORDER BY {5} ASC, {6} ASC, {7} ASC ",
																 TableName,
																 IdxColumnName.RESULT.ToString,
																 IdxColumnName.WAIT_UUID.ToString,
																 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
																 IdxColumnName.CREATE_TIME.ToString,
																 IdxColumnName.UUID.ToString,
																 IdxColumnName.SEQ.ToString,
														 IdxColumnName.CLIENT_ID, _clientidValue)
				End If

				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	'從資料庫抓取Command，沒有處理過的
	Public Shared Function GetCommandDictionaryByReceiveSystem_ResultIsNULL_WaitUUIDIsNull(ByVal ReceiveSystem As EnuSystemType) As Dictionary(Of String, clsFromGUICommand)

		Try
			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				strSQL = String.Format("Select * FROM {0} WHERE {1} IS NULL AND {2} IS NULL AND  {3}={4} ORDER BY {5} ASC, {6} ASC, {7} ASC ",
																	 TableName,
																	 IdxColumnName.RESULT.ToString,
																	 IdxColumnName.WAIT_UUID.ToString,
																	 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem),
																	 IdxColumnName.CREATE_TIME.ToString,
																	 IdxColumnName.UUID.ToString,
																	 IdxColumnName.SEQ.ToString)
				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	'從資料庫抓取Command，使用Wait_UUID當條件
	Public Shared Function GetCommandDictionaryByReceiveSystem_WaitUUID(ByVal ReceiveSystem As EnuSystemType, ByVal Wait_UUID As String) As Dictionary(Of String, clsFromGUICommand)

		Try
			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				strSQL = String.Format("Select * FROM {0} WHERE {1} IS NULL AND {2}='{3}' AND {4}='{5}' ",
																	 TableName,
																	 IdxColumnName.RESULT.ToString,
																	 IdxColumnName.WAIT_UUID.ToString, Wait_UUID,
																	 IdxColumnName.RECEIVE_SYSTEM, CInt(ReceiveSystem))
				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	'從資料庫抓取Command，已經處理過的
	Public Shared Function GetCommandDictionaryBySendSystem_ResultIsNotNULL(ByVal SendSystem As EnuSystemType) As Dictionary(Of String, clsFromGUICommand)

		Try
			Dim ret_dic As New Dictionary(Of String, clsFromGUICommand)
			If DBTool IsNot Nothing Then
				Dim strSQL As String = String.Empty
				Dim DatasetMessage As New DataSet
				strSQL = String.Format("Select * FROM {0} WHERE {1} IS NOT NULL AND {2}='{3}' ",
																	 TableName,
																	 IdxColumnName.RESULT.ToString,
																	 IdxColumnName.SEND_SYSTEM, CInt(SendSystem))
				DBTool.SQLExcute(strSQL, DatasetMessage)
				If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
					For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
						Dim Info As clsFromGUICommand = Nothing
						If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
							If Info IsNot Nothing Then
								If ret_dic.ContainsKey(Info.gid) = False Then
									ret_dic.Add(Info.gid, Info)
								End If
							Else
								SendMessageToLog("Get clsFromGUICommand Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
							End If
						Else
							SendMessageToLog("Get clsFromGUICommand Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
						End If
					Next
				End If
			End If
			Return ret_dic
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function
	'-不要動
	Private Shared Function SetInfoFromDB(ByRef Info As clsFromGUICommand, ByRef RowData As DataRow) As Boolean
		Try
			If RowData IsNot Nothing Then
				Dim UUID = "" & RowData.Item(IdxColumnName.UUID.ToString)
				Dim Send_System = 0 & RowData.Item(IdxColumnName.SEND_SYSTEM.ToString)
				Dim Receive_System = 0 & RowData.Item(IdxColumnName.RECEIVE_SYSTEM.ToString)
				Dim Function_ID = "" & RowData.Item(IdxColumnName.FUNCTION_ID.ToString)
				Dim SEQ = 0 & RowData.Item(IdxColumnName.SEQ.ToString)
				Dim User_ID = "" & RowData.Item(IdxColumnName.USER_ID.ToString)
				Dim Create_Time = "" & RowData.Item(IdxColumnName.CREATE_TIME.ToString)
				Dim Message = "" & RowData.Item(IdxColumnName.MESSAGE.ToString)
				Dim Result = "" & RowData.Item(IdxColumnName.RESULT.ToString)
				Dim Result_Message = "" & RowData.Item(IdxColumnName.RESULT_MESSAGE.ToString)
				Dim Wait_UUID = "" & RowData.Item(IdxColumnName.WAIT_UUID.ToString)
				Dim CLIENT_ID = "" & RowData.Item(IdxColumnName.CLIENT_ID.ToString)
				Dim IP = "" & RowData.Item(IdxColumnName.IP.ToString)
				Info = New clsFromGUICommand(UUID, Send_System, Receive_System, Function_ID, SEQ, User_ID, Create_Time, Message, Result,
																		 Result_Message, Wait_UUID, CLIENT_ID, IP)
			End If
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
	'如果是不同DB時使用(寫入後要等回覆)
	Public Shared Function BatchUpdate(ByRef lstSQL As List(Of String)) As Boolean
		Try
			If lstSQL Is Nothing Then Return -1
			If lstSQL.Count = 0 Then Return 0
			For i = 0 To lstSQL.Count - 1
				SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
			Next
			Dim rtnMsg As String = DBTool.BatchUpdate(lstSQL)
			If rtnMsg.StartsWith("OK") Then
				SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			Else
				SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
				Return False
			End If
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function
	'如果是不同DB時使用(寫入後不等回覆)
	Public Shared Function AddQueued(ByRef lstSQL As List(Of String)) As Boolean
		Try
			If lstSQL Is Nothing Then Return -1
			If lstSQL.Count = 0 Then Return 0
			For i = 0 To lstSQL.Count - 1
				SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
			Next
			DBTool.O_AddQueue_File((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, lstSQL)
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function


	Private Shared Function SQLCorrect(ByVal SQL As String, ByRef NewSQL As String) As Boolean
		Try
			Dim NewStr = ""
			Dim whereFlag = False
			For Each splitstr In SQL.Split(" ")
				If whereFlag Then '轉
					NewStr += splitstr.Replace("=''", " is null ") & " "
				Else '不用轉
					NewStr += splitstr & " "
				End If
				'找到where where後的=''轉成 not null
				If splitstr.ToUpper() = "WHERE" Then
					whereFlag = True
				End If
			Next

			NewSQL = NewStr
			Return True
		Catch ex As Exception
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return False
		End Try
	End Function
End Class
