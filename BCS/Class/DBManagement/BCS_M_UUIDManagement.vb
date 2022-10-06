Imports System.Collections.Concurrent

Partial Class BCS_M_UUIDManagement
  Public Shared TableName As String = "BCS_M_UUID"
  Public Shared objLock As New Object
  Private Shared fUseBatchUpdate As Integer = 0
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    UUID_NO
    UUID_SEQ
    IDLENGTH
    APPEND
    COMMENTS
    RESETABLE
    UPDATE_DATE
  End Enum
  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsUUID) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1} ({2},{4},{6},{8},{10},{12},{14}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}')",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.UUID_NO.ToString,
      IdxColumnName.UUID_SEQ.ToString, Info.UUID_SEQ,
      IdxColumnName.IDLENGTH.ToString, Info.IDLENGTH,
      IdxColumnName.APPEND.ToString, Info.APPEND,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.RESETABLE.ToString, BooleanConvertToInteger(Info.RESETABLE),
      IdxColumnName.UPDATE_DATE.ToString, Info.UPDATE_DATE
      )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetDeleteSQL(ByRef Info As clsUUID) As String
    Try

      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' ",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.UUID_NO.ToString
      )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQL(ByRef Info As clsUUID) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}' WHERE {2}='{3}'",
      strSQL,
      TableName,
      IdxColumnName.UUID_NO.ToString, Info.UUID_NO.ToString,
      IdxColumnName.UUID_SEQ.ToString, Info.UUID_SEQ,
      IdxColumnName.IDLENGTH.ToString, Info.IDLENGTH,
      IdxColumnName.APPEND.ToString, Info.APPEND,
      IdxColumnName.COMMENTS.ToString, Info.COMMENTS,
      IdxColumnName.RESETABLE.ToString, BooleanConvertToInteger(Info.RESETABLE),
      IdxColumnName.UPDATE_DATE.ToString, Info.UPDATE_DATE
      )
      Dim NewSQL As String = ""
      If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
        Return NewSQL
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetUpdateSQLForChangeValue(ByRef Info As clsUUID, ByRef dicChangeColumnValue As Dictionary(Of String, String)) As String
    Try
      Dim strSQL As String = ""
      Dim strUpdateColumnValue As String = ""
      If O_Get_UpdateColumnSQL(Of IdxColumnName)(dicChangeColumnValue, strUpdateColumnValue) = True Then
        If strUpdateColumnValue <> "" Then
          strSQL = String.Format("Update {1} SET {4}  WHERE {2}='{3}' ",
          strSQL,
          TableName,
          IdxColumnName.UUID_NO.ToString, Info.UUID_NO,
          strUpdateColumnValue)
          Dim NewSQL As String = ""
          If SQLCorrect(DBTool.m_nDBType, strSQL, NewSQL) Then
            Return NewSQL
          End If
        End If
      End If
      Return Nothing
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  '- Update
  Public Shared Function UpdateWMS_M_UUIDData(ByVal Info As clsUUID, Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If UpdatelstWMS_M_UUIDData(New List(Of clsUUID)({Info}), SendToDB) = True Then
          Return True
        End If
        Return False
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  Public Shared Function UpdatelstWMS_M_UUIDData(ByVal Info As List(Of clsUUID), Optional ByVal SendToDB As Boolean = True) As Boolean
    SyncLock objLock
      Try
        If Info Is Nothing Then Return False
        If Info.Count = 0 Then Return True
        If SendToDB Then
          If UpdateWMS_M_UUIDDataToDB(Info) Then
            Return True
          Else
            SendMessageToLog("UpdateDB WMS_M_UUIDData Fail", eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False
          End If
        Else
          SendMessageToLog("Do nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return True
        End If
        Return True
      Catch ex As Exception
        Return False
      End Try
    End SyncLock
  End Function
  '- GET
  Public Shared Function GetWMS_M_UUIDDataListByALL() As List(Of clsUUID)
    Try
      Dim _lstReturn As New List(Of clsUUID)
      If DBTool IsNot Nothing Then
        'If DBTool.isConnection(DBTool.m_CN) = True Then
        Dim strSQL As String = String.Empty
        Dim rs As DataSet = Nothing
        Dim DatasetMessage As New DataSet
        strSQL = String.Format("Select * from {0}", TableName)
        SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
        DBTool.SQLExcute(strSQL, DatasetMessage)

        'Dim OLEDBAdapter As New OleDbDataAdapter
        'OLEDBAdapter.Fill(DatasetMessage, rs, TableName)

        If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
          For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
            Dim Info As clsUUID = Nothing
            SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex))
            _lstReturn.Add(Info)
          Next
        End If
        'End If
      End If
      Return _lstReturn
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Public Shared Function GetclsUUIDListByUUID_NO(ByVal uuid_no As String) As Dictionary(Of String, clsUUID)
    SyncLock objLock
      Try
        Dim ret_dic As New Dictionary(Of String, clsUUID)

        If DBTool IsNot Nothing Then
          If DBTool.isConnection(DBTool.m_CN) = True Then
            Dim strSQL As String = String.Empty
            Dim DatasetMessage As New DataSet

            strSQL = String.Format("Select * from {1} WHERE  {2} = '{3}' ",
            strSQL,
            TableName,
            IdxColumnName.UUID_NO.ToString, uuid_no
            )
            SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            DBTool.SQLExcute(strSQL, DatasetMessage)


            If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
              For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
                Dim Info As clsUUID = Nothing
                If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) Then
                  If Info IsNot Nothing Then
                    If ret_dic.ContainsKey(Info.gid) = False Then
                      ret_dic.Add(Info.gid, Info)
                    End If
                  Else
                    SendMessageToLog("Get clsUUID Info is nothing ", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                  End If
                Else
                  SendMessageToLog("Get clsUUID Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Next
            End If
          End If
          Return ret_dic
        End If
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return Nothing
      End Try
    End SyncLock
  End Function
  '-Function
  Private Shared Function UpdateWMS_M_UUIDDataToDB(ByRef Info As List(Of clsUUID)) As Boolean
    Try
      If Info Is Nothing Then Return False
      If Info.Count = 0 Then Return True

      Dim strSQL As String = ""
      Dim rs As DataSet = Nothing
      Dim lstSql As New List(Of String)
      For Each CI In Info
        strSQL = String.Format("Update {1} SET {4}='{5}',{6}='{7}',{8}='{9}',{10}='{11}',{12}='{13}',{14}='{15}' WHERE {2}='{3}'",
        strSQL,
        TableName,
        IdxColumnName.UUID_NO.ToString, CI.UUID_NO.ToString,
        IdxColumnName.UUID_SEQ.ToString, CI.UUID_SEQ,
        IdxColumnName.IDLENGTH.ToString, CI.IDLENGTH,
        IdxColumnName.APPEND.ToString, CI.APPEND,
        IdxColumnName.COMMENTS.ToString, CI.COMMENTS,
        IdxColumnName.RESETABLE, BooleanConvertToInteger(CI.RESETABLE),
        IdxColumnName.UPDATE_DATE.ToString, CI.UPDATE_DATE
      )
        lstSql.Add(strSQL)
      Next
      Dim NewSQL As New List(Of String)
      If SQLCorrect(DBTool.m_nDBType, lstSql, NewSQL) = False Then
        Return Nothing
      End If
      If SendSQLToDB(NewSQL) = True Then
        Return True
      Else
        SendMessageToLog("Update to WMS_M_UUIDData DB Error", eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Shared Function SendSQLToDB(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim FullTimeUUID As String = GetNewTime_ByDataTimeFormat(DBFullTimeUUIDFormat)
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog(FullTimeUUID & "SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      If fUseBatchUpdate = 0 Then
        For i = 0 To lstSQL.Count - 1
          If DBTool.O_AddSQLQueue(TableName, lstSQL(i)) = True Then
            SendMessageToLog("O_AddSQLQueue Success", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          Else
            SendMessageToLog("O_AddSQLQueue Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          End If
        Next
      Else
        Dim rtnMsg As String = DBTool.BatchUpdate(lstSQL)
        If rtnMsg.StartsWith("OK") Then
          SendMessageToLog(FullTimeUUID & " " & rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
        Else
          SendMessageToLog(FullTimeUUID & " " & rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsUUID, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim UUID_NO = "" & RowData.Item(IdxColumnName.UUID_NO.ToString)
        Dim UUID_SEQ = IIf(IsNumeric(RowData.Item(IdxColumnName.UUID_SEQ.ToString)), RowData.Item(IdxColumnName.UUID_SEQ.ToString), 0)
        Dim IDLENGTH = IIf(IsNumeric(RowData.Item(IdxColumnName.IDLENGTH.ToString)), RowData.Item(IdxColumnName.IDLENGTH.ToString), 0)
        Dim APPEND = "" & RowData.Item(IdxColumnName.APPEND.ToString)
        Dim COMMENTS = "" & RowData.Item(IdxColumnName.COMMENTS.ToString)
        Dim RESETABLE = IntegerConvertToBoolean("" & RowData.Item(IdxColumnName.RESETABLE.ToString))
        Dim UPDATE_DATE = "" & RowData.Item(IdxColumnName.UPDATE_DATE.ToString)
        Info = New clsUUID(UUID_NO, UUID_SEQ, IDLENGTH, APPEND, COMMENTS, RESETABLE, UPDATE_DATE)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString(), eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class