Imports System.Collections.Concurrent
Partial Class PLANFORM_LOTNO_LISTManagement
  Public Shared TableName As String = "PLANFORM_LOTNO_LIST"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    PlatForm
    LotNo
    Count
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsPLANFORM_LOTNO_LIST) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6}) values ('{3}','{5}','{7}')",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlanForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.Count.ToString, CInt(Info.Count)
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsPLANFORM_LOTNO_LIST) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {6}='{7}' WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlanForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.Count.ToString, CInt(Info.Count)
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsPLANFORM_LOTNO_LIST) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}'",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlanForm,
     IdxColumnName.LotNo.ToString, Info.LotNo
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

  Public Shared Function GetDataDictionaryByKEY(ByVal PlatForm As String, ByVal LotNo As String) As Dictionary(Of String, clsPLANFORM_LOTNO_LIST)
    Try
      Dim ret_dic As New Dictionary(Of String, clsPLANFORM_LOTNO_LIST)
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
              Dim Info As clsPLANFORM_LOTNO_LIST = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsPLANFORM_LOTNO_LIST Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsPLANFORM_LOTNO_LIST Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function SetInfoFromDB(ByRef Info As clsPLANFORM_LOTNO_LIST, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim PlanForm = "" & RowData.Item(IdxColumnName.PlatForm.ToString)
        Dim LotNo = "" & RowData.Item(IdxColumnName.LotNo.ToString)
        Dim Count = IIf(IsNumeric(RowData.Item(IdxColumnName.Count.ToString)), RowData.Item(IdxColumnName.Count.ToString), 0)
        Info = New clsPLANFORM_LOTNO_LIST(PlanForm, LotNo, Count)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
