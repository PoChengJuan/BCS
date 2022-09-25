Imports System.Collections.Concurrent
Partial Class STORE_HEADManagement
  Public Shared TableName As String = "STORE_HEAD"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PlatForm
    LotNo
    Store_ID
    Memo
    CreateTime
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsSTORE_HEAD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10}) values ('{3}','{5}','{7}','{9}','{11}')",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.Store_ID.ToString, Info.Store_ID,
     IdxColumnName.Memo.ToString, Info.Memo,
     IdxColumnName.CreateTime.ToString, Info.CreateTime
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsSTORE_HEAD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.Store_ID.ToString, Info.Store_ID,
     IdxColumnName.Memo.ToString, Info.Memo,
     IdxColumnName.CreateTime.ToString, Info.CreateTime
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsSTORE_HEAD) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.Store_ID.ToString, Info.Store_ID
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

  Public Shared Function GetDataDictionaryByKEY(ByVal PlatForm As String, ByVal LotNo As String, ByVal Store_ID As String) As Dictionary(Of String, clsSTORE_HEAD)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_HEAD)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If PlatForm <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.PlatForm.ToString, PlatForm)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.PlatForm.ToString, PlatForm)
            End If
          End If
          If LotNo <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.LotNo.ToString, LotNo)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.LotNo.ToString, LotNo)
            End If
          End If
          If Store_ID <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.Store_ID.ToString, Store_ID)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.Store_ID.ToString, Store_ID)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select * from {1} {2} ",
          strSQL,
          TableName,
          strWhere
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute(strSQL, DatasetMessage)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsSTORE_HEAD = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_HEAD Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_HEAD Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
              End If
            Next
          End If
        End If
      End If
      Return ret_dic
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '-不要動
  Private Shared Function SetInfoFromDB(ByRef Info As clsSTORE_HEAD, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PlatForm = "" & RowData.Item(IdxColumnName.PlatForm.ToString)
        Dim LotNo = "" & RowData.Item(IdxColumnName.LotNo.ToString)
        Dim Store_ID = "" & RowData.Item(IdxColumnName.Store_ID.ToString)
        Dim Memo = "" & RowData.Item(IdxColumnName.Memo.ToString)
        Dim CreateTime = "" & RowData.Item(IdxColumnName.CreateTime.ToString)
        Info = New clsSTORE_HEAD(PlatForm, LotNo, Store_ID, Memo, CreateTime)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
