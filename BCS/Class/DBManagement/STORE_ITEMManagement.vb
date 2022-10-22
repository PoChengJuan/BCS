Imports System.Collections.Concurrent
Partial Class STORE_ITEMManagement
  Public Shared TableName As String = "STORE_ITEM"
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing

  Enum IdxColumnName As Integer
    SERIAL_NO
    PlatForm
    LotNo
    BarCode1
    BarCode2
    BarCode3
    BarCode4
    CreateTime
  End Enum

  '- GetSQL
  Public Shared Function GetInsertSQL(ByRef Info As clsSTORE_ITEM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Insert into {1}({2},{4},{6},{8},{10},{12},{14}) values ('{3}','{5}','{7}','{9}','{11}','{13}','{15}')",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.BarCode1.ToString, Info.BarCode1,
     IdxColumnName.BarCode2.ToString, Info.BarCode2,
     IdxColumnName.BarCode3.ToString, Info.BarCode3,
     IdxColumnName.BarCode4.ToString, Info.BarCode4,
     IdxColumnName.CreateTime.ToString, Info.CreateTime
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
  Public Shared Function GetUpdateSQL(ByRef Info As clsSTORE_ITEM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Update {1} SET {8}='{9}',{10}='{11}',{12}='{13}'' WHERE {2}='{3}' AND {4}='{5}' AND {6}='{7}'",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
     IdxColumnName.LotNo.ToString, Info.LotNo,
     IdxColumnName.BarCode1.ToString, Info.BarCode1,
     IdxColumnName.BarCode2.ToString, Info.BarCode2,
     IdxColumnName.BarCode3.ToString, Info.BarCode3,
     IdxColumnName.CreateTime.ToString, Info.CreateTime
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
  Public Shared Function GetDeleteSQL(ByRef Info As clsSTORE_ITEM) As String
    Try
      Dim strSQL As String = ""
      strSQL = String.Format("Delete From {1} WHERE {2}='{3}' AND {4}='{5}' ",
      strSQL,
      TableName,
     IdxColumnName.PlatForm.ToString, Info.PlatForm,
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

  Public Shared Function GetDataDictionaryByKEY(ByVal PlatForm As String, ByVal LotNo As String) As Dictionary(Of String, clsSTORE_ITEM)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_ITEM)
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
              Dim Info As clsSTORE_ITEM = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByALL_BarCode(ByVal BarCode1 As String, ByVal BarCode2 As String, ByVal BarCode3 As String, ByVal BarCode4 As String) As Dictionary(Of String, clsSTORE_ITEM)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_ITEM)
      If DBTool IsNot Nothing Then
        If DBTool.isConnection(DBTool.m_CN) = True Then
          Dim strWhere As String = ""
          If BarCode1 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode1.ToString, BarCode1)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode1.ToString, BarCode1)
            End If
          End If
          If BarCode2 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode2.ToString, BarCode2)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode2.ToString, BarCode2)
            End If
          End If
          If BarCode3 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode3.ToString, BarCode3)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode3.ToString, BarCode3)
            End If
          End If
          If BarCode4 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode4.ToString, BarCode4)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode4.ToString, BarCode4)
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
              Dim Info As clsSTORE_ITEM = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryByItemReport(ByVal PlatForm As String, ByVal LotNo As String, ByVal Start_Time As String, ByVal End_Time As String) As Dictionary(Of String, clsSTORE_ITEM)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_ITEM)
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
          If Start_Time <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} > '{1}' ", IdxColumnName.CreateTime.ToString, Start_Time)
            Else
              strWhere = String.Format("{0} AND {1} > '{2}' ", strWhere, IdxColumnName.CreateTime.ToString, Start_Time)
            End If
          End If
          If End_Time <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} < '{1}' ", IdxColumnName.CreateTime.ToString, End_Time)
            Else
              strWhere = String.Format("{0} AND {1} < '{2}' ", strWhere, IdxColumnName.CreateTime.ToString, End_Time)
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
              Dim Info As clsSTORE_ITEM = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetGroupDataDictionaryByTime(ByVal PlatForm As String, ByVal Start_Time As String, ByVal End_Time As String) As Dictionary(Of String, clsPLANFORM_LOTNO_LIST)
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

          If Start_Time <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} > '{1}' ", IdxColumnName.CreateTime.ToString, Start_Time)
            Else
              strWhere = String.Format("{0} AND {1} > '{2}' ", strWhere, IdxColumnName.CreateTime.ToString, Start_Time)
            End If
          End If
          If End_Time <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} < '{1}' ", IdxColumnName.CreateTime.ToString, End_Time)
            Else
              strWhere = String.Format("{0} AND {1} < '{2}' ", strWhere, IdxColumnName.CreateTime.ToString, End_Time)
            End If
          End If
          Dim strSQL As String = String.Empty
          Dim rs As DataSet = Nothing
          Dim DatasetMessage As New DataSet
          strSQL = String.Format("Select PlanForm,LotNo,count(LotNo) as Count from {1} {2} group by PLATFORM,LOTNO order by PLATFORM,LOTNO",
          strSQL,
          TableName,
          strWhere
          )
          SendMessageToLog(strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          DBTool.SQLExcute(strSQL, DatasetMessage)
          If DatasetMessage.Tables.Item(0).Rows.Count > 0 Then
            For RowIndex = 0 To DatasetMessage.Tables.Item(0).Rows.Count - 1
              Dim Info As clsPLANFORM_LOTNO_LIST = Nothing
              If PLANFORM_LOTNO_LISTManagement.SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryBySearch_Data(ByVal PlatForm As String, ByVal BarCode1 As String, ByVal BarCode2 As String, ByVal BarCode3 As String) As Dictionary(Of String, clsSTORE_ITEM)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_ITEM)
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
          If BarCode1 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode1.ToString, BarCode1)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode1.ToString, BarCode1)
            End If
          End If
          If BarCode2 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode2.ToString, BarCode2)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode2.ToString, BarCode2)
            End If
          End If
          If BarCode3 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} = '{1}' ", IdxColumnName.BarCode3.ToString, BarCode3)
            Else
              strWhere = String.Format("{0} AND {1} = '{2}' ", strWhere, IdxColumnName.BarCode3.ToString, BarCode3)
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
              Dim Info As clsSTORE_ITEM = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Public Shared Function GetDataDictionaryBySearch_FamilyData(ByVal PlatForm As String, ByVal BarCode1 As String) As Dictionary(Of String, clsSTORE_ITEM)
    Try
      Dim ret_dic As New Dictionary(Of String, clsSTORE_ITEM)
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
          If BarCode1 <> "" Then
            If strWhere = "" Then
              strWhere = String.Format("WHERE {0} like '%{1}%' ", IdxColumnName.BarCode1.ToString, BarCode1)
            Else
              strWhere = String.Format("{0} AND {1} like '%{2}%' ", strWhere, IdxColumnName.BarCode1.ToString, BarCode1)
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
              Dim Info As clsSTORE_ITEM = Nothing
              If SetInfoFromDB(Info, DatasetMessage.Tables.Item(0).Rows(RowIndex)) = True Then
                If Info IsNot Nothing Then
                  If ret_dic.ContainsKey(Info.gid) = False Then
                    ret_dic.Add(Info.gid, Info)
                  End If
                Else
                  SendMessageToLog("Get clsSTORE_ITEM Info Is Nothing", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
                End If
              Else
                SendMessageToLog("Get clsSTORE_ITEM Info Failed", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
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
  Private Shared Function SetInfoFromDB(ByRef Info As clsSTORE_ITEM, ByRef RowData As DataRow) As Boolean
    Try
      If RowData IsNot Nothing Then
        Dim SERIAL_NO = "" & RowData.Item(IdxColumnName.SERIAL_NO.ToString)
        Dim PlatForm = "" & RowData.Item(IdxColumnName.PlatForm.ToString)
        Dim LotNo = "" & RowData.Item(IdxColumnName.LotNo.ToString)
        Dim BarCode1 = "" & RowData.Item(IdxColumnName.BarCode1.ToString)
        Dim BarCode2 = "" & RowData.Item(IdxColumnName.BarCode2.ToString)
        Dim BarCode3 = "" & RowData.Item(IdxColumnName.BarCode3.ToString)
        Dim BarCode4 = "" & RowData.Item(IdxColumnName.BarCode4.ToString)
        Dim CreateTime = "" & RowData.Item(IdxColumnName.CreateTime.ToString)
        Info = New clsSTORE_ITEM(SERIAL_NO, PlatForm, LotNo, BarCode1, BarCode2, BarCode3, BarCode4, CreateTime)
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
