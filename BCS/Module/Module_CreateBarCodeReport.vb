
Imports System.Drawing
Imports GenCode128
Imports System.Diagnostics

Module Module_CreateBarCodeReport
  Public Function O_Process_Message(ByVal PlatForm As String,
                                    ByVal LotNo As String,
                                    ByVal Start_time As String,
                                    ByVal End_Time As String,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try

      Dim dicStore_Item As New Dictionary(Of String, clsSTORE_ITEM)
      Dim dicAddStore_Head As New Dictionary(Of String, clsSTORE_HEAD)
      Dim dicAddStore_Item As New Dictionary(Of String, clsSTORE_ITEM)

      Dim lstSQL As New List(Of String)
      Dim lstLCSSql As New List(Of String)
      Dim lstHistroySQL As New List(Of String)

      '檢查資料
      If Check_Data(PlatForm, LotNo, Start_time, End_Time, dicStore_Item, ret_strResultMsg) = False Then
        Return False
      End If
      ''取得更新資料
      If Get_Data(PlatForm, LotNo, dicStore_Item, ret_strResultMsg) = False Then
        Return False
      End If
      ''取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, dicAddStore_Head, dicAddStore_Item, lstSQL) = False Then
        Return False
      End If
      ''執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, lstSQL) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''檢查資料
  Private Function Check_Data(ByRef ret_PlatForm As String,
                              ByRef ret_LotNo As String,
                              ByRef ret_StartTime As String,
                              ByRef ret_EndTime As String,
                              ByRef ret_dicStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      ret_dicStore_Item = STORE_ITEMManagement.GetDataDictionaryByItemReport(ret_PlatForm, ret_LotNo, ret_StartTime, ret_EndTime)
      If ret_dicStore_Item.Any = False Then
        ret_strResultMsg = "查無條碼"
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''取得更新資料
  Private Function Get_Data(ByVal ret_PlatForm As String,
                            ByVal ret_LotNo As String,
                            ByVal ret_dicStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                            ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Now_Time As String = GetNewTime_DBFormat()

      Dim Store_Head_Key = clsSTORE_HEAD.Get_Combination_Key(ret_PlatForm, ret_LotNo, ret_BarCode1)
      Dim dicSTORE_HEAD = STORE_HEADManagement.GetDataDictionaryByKEY(ret_PlatForm, ret_LotNo, ret_BarCode1)
      If dicSTORE_HEAD.Any = False Then
        Dim objStore_Head = New clsSTORE_HEAD(ret_PlatForm, ret_LotNo, ret_BarCode1, "", Now_Time)
        If ret_dicAddStore_Head.ContainsKey(objStore_Head.gid) = False Then
          ret_dicAddStore_Head.Add(objStore_Head.gid, objStore_Head)
        End If
      End If

      Dim objStore_Item = New clsSTORE_ITEM(ret_PlatForm, ret_LotNo, ret_BarCode1, ret_BarCode2, ret_BarCode3, ret_BarCode4, Now_Time)
      If ret_dicAddStore_Item.ContainsKey(objStore_Item.gid) = False Then
        ret_dicAddStore_Item.Add(objStore_Item.gid, objStore_Item)
      End If

      If ret_strResultMsg.Length > 0 Then
        Return False
      Else
        Return True
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicAddStore_Head As Dictionary(Of String, clsSTORE_HEAD),
                           ByRef ret_dicAddStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得新增SQL
      For Each obj As clsSTORE_HEAD In ret_dicAddStore_Head.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Store_Head SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj As clsSTORE_ITEM In ret_dicAddStore_Item.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Store_Item SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      'If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
      '  '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
      '  ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        'ret_strResultMsg = "WMS Update DB Failed"
        ret_strResultMsg = "WMS 更新资料库失败"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      '修改記憶體資料
      '1.新增
      'For Each objNew As clsALARM In ret_dicAddAlarm.Values
      '  objNew.Add_Relationship(gMain.objHandling)
      'Next
      ''2.删除
      'For Each objALARM As clsALARM In ret_dicDeleteAlarm.Values
      '  objALARM.Remove_Relationship()
      'Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Enum enuArrival_Status
    BCR = 1
    Palletizing = 2
  End Enum
End Module
