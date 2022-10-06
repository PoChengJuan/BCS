Public Class clsUUID
  Private ShareName As String = "UUID"
  Private ShareKey As String = ""

  Private _gid As String
  Private _UUID_NO As String '編號名稱
  Private _UUID_SEQ As Long '流水號
  Private _IDLENGTH As Long '流水號長度
  Private _APPEND As String '編碼規則
  Private _COMMENTS As String '備註
  Private _RESETABLE As Boolean '超過上限是否自動Reset
  Private _UPDATE_DATE As String '更新日期
  'Public objWMS As clsWMSObject
  Private objUUIDLock As New Object

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property UUID_NO() As String
    Get
      Return _UUID_NO
    End Get
    Set(ByVal value As String)
      _UUID_NO = value
    End Set
  End Property
  Public Property UUID_SEQ() As Long
    Get
      Return _UUID_SEQ
    End Get
    Set(ByVal value As Long)
      _UUID_SEQ = value
    End Set
  End Property
  Public Property IDLENGTH() As Long
    Get
      Return _IDLENGTH
    End Get
    Set(ByVal value As Long)
      _IDLENGTH = value
    End Set
  End Property
  Public Property APPEND() As String
    Get
      Return _APPEND
    End Get
    Set(ByVal value As String)
      _APPEND = value
    End Set
  End Property
  Public Property COMMENTS() As String
    Get
      Return _COMMENTS
    End Get
    Set(ByVal value As String)
      _COMMENTS = value
    End Set
  End Property
  Public Property RESETABLE() As Boolean
    Get
      Return _RESETABLE
    End Get
    Set(ByVal value As Boolean)
      _RESETABLE = value
    End Set
  End Property
  Public Property UPDATE_DATE() As String
    Get
      Return _UPDATE_DATE
    End Get
    Set(ByVal value As String)
      _UPDATE_DATE = value
    End Set
  End Property




  '物件建立時執行的事件
  Public Sub New(ByVal UUID_NO As String, ByVal UUID_SEQ As Long, ByVal IDLENGTH As Long, ByVal APPEND As String, ByVal COMMENTS As String, ByVal RESETABLE As Boolean, ByVal UPDATE_DATE As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(UUID_NO)
      _gid = key
      _UUID_NO = UUID_NO
      _UUID_SEQ = UUID_SEQ
      _IDLENGTH = IDLENGTH
      _APPEND = APPEND
      _COMMENTS = COMMENTS
      _RESETABLE = RESETABLE
      _UPDATE_DATE = UPDATE_DATE
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  '物件結束時觸發的事件，用來清除物件的內容
  'Protected Overrides Sub Finalize()
  '  Class_Terminate_Renamed()
  '  MyBase.Finalize()
  'End Sub
  'Private Sub Class_Terminate_Renamed()
  '  '目的:結束物件
  '  objWMS = Nothing
  'End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal UUID_NO As String) As String
    Try
      Dim key As String = UUID_NO
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  'Public Sub Add_Relationship(ByRef objWMS As clsWMSObject)
  '  Try
  '    '挷定Customer和WMS的關係
  '    If objWMS IsNot Nothing Then
  '      Me.objWMS = objWMS
  '      objWMS.O_Add_UUID(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub
  'Public Sub Remove_Relationship()
  '  Try
  '    '解除Block和WMS的關係
  '    If Me.objWMS IsNot Nothing Then
  '      Me.objWMS.O_Remove_UUID(Me)
  '    End If
  '  Catch ex As Exception
  '    SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
  '  End Try
  'End Sub

  Public Function Update_UUID(ByRef obj As clsUUID) As Boolean
    Try
      Dim key As String = obj._gid
      If key <> _gid Then
        SendMessageToLog("Key can not Update, old_Key=" & _gid & " ,new_key=" & key, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      _UUID_SEQ = obj.UUID_SEQ
      _IDLENGTH = obj.IDLENGTH
      _APPEND = obj.APPEND
      _COMMENTS = obj.COMMENTS
      _RESETABLE = obj.RESETABLE
      _UPDATE_DATE = obj.UPDATE_DATE
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Public Function Get_NewUUID(Optional ByRef ret_ResultMsg As String = "") As String
    Try
      SyncLock objUUIDLock
        Dim rtnSEQ As String = "" '-回覆的UUID
        For Each item In _APPEND.Split(",")
          Select Case item.Chars(0)
            Case "@"
              If item.Contains("UUID") Then
                rtnSEQ += _UUID_SEQ.ToString.PadLeft(_IDLENGTH.ToString.Length, "0")
              Else  '除了UUID目前只設定日期
                Try '-防止錯誤日期格式
                  Dim NewDate As String = GetNewTime_ByDataTimeFormat(item.TrimStart("@"))
                  If NewDate <> _UPDATE_DATE Then
                    _UPDATE_DATE = NewDate
                    _UUID_SEQ = 1
                  End If
                  rtnSEQ += GetNewTime_ByDataTimeFormat(item.TrimStart("@"))
                Catch ex As Exception
                  ret_ResultMsg = "UUID APPEND Error, UUID_NO=" & _UUID_NO
                  SendMessageToLog(ret_ResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
                End Try
              End If
            Case Else
              rtnSEQ += item
          End Select
        Next
        '-檢查是否過限制(如果該次已經超過限制，表示該UUID不可以使用)
        If _UUID_SEQ > _IDLENGTH Then
          rtnSEQ = ""
          ret_ResultMsg = "Exceeding the defined range(超出設定的範圍，請重新定義), UUID_NO=" & _UUID_NO
          SendMessageToLog(ret_ResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        End If
        '-數值+1
        _UUID_SEQ = _UUID_SEQ + 1
        '-檢查是否過限制
        If _UUID_SEQ > _IDLENGTH Then
          If _RESETABLE = True Then
            _UUID_SEQ = 1
          End If
        End If
        '在資料庫Update UUID的資料
        If BCS_M_UUIDManagement.UpdateWMS_M_UUIDData(Me) Then
          Return rtnSEQ
        Else
          ret_ResultMsg = "Update UUID to DB Failed ,TableName = " & BCS_M_UUIDManagement.TableName & " ,UUID_NO=" & _UUID_NO
          SendMessageToLog(ret_ResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        End If
        Return rtnSEQ
      End SyncLock
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Class
