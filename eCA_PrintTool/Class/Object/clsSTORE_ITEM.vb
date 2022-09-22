Public Class clsSTORE_ITEM
  Private ShareName As String = "STORE_ITEM"
  Private ShareKey As String = ""
  Private _gid As String
  Private _PlatForm As String '平台 
  Private _LotNo As String '賣場 
  Private _Store_ID As String '店代號 
  Private _BarCode1 As String '條碼1 
  Private _BarCode2 As String '條碼2 
  Private _BarCode3 As String '條碼3 
  Private _CreateTime As String '建立時間 


  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property PlatForm() As String
    Get
      Return _PlatForm
    End Get
    Set(ByVal value As String)
      _PlatForm = value
    End Set
  End Property
  Public Property LotNo() As String
    Get
      Return _LotNo
    End Get
    Set(ByVal value As String)
      _LotNo = value
    End Set
  End Property
  Public Property Store_ID() As String
    Get
      Return _Store_ID
    End Get
    Set(ByVal value As String)
      _Store_ID = value
    End Set
  End Property
  Public Property BarCode1() As String
    Get
      Return _BarCode1
    End Get
    Set(ByVal value As String)
      _BarCode1 = value
    End Set
  End Property
  Public Property BarCode2() As String
    Get
      Return _BarCode2
    End Get
    Set(ByVal value As String)
      _BarCode2 = value
    End Set
  End Property
  Public Property BarCode3() As String
    Get
      Return _BarCode3
    End Get
    Set(ByVal value As String)
      _BarCode3 = value
    End Set
  End Property
  Public Property CreateTime() As String
    Get
      Return _CreateTime
    End Get
    Set(ByVal value As String)
      _CreateTime = value
    End Set
  End Property


  '物件建立時執行的事件
  Public Sub New(ByVal PlatForm As String, ByVal LotNo As String, ByVal Store_ID As String, ByVal BarCode1 As String, ByVal BarCode2 As String, ByVal BarCode3 As String, ByVal CreateTime As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(PlatForm, LotNo, Store_ID)
      _gid = key
      _PlatForm = PlatForm
      _LotNo = LotNo
      _Store_ID = Store_ID
      _BarCode1 = BarCode1
      _BarCode2 = BarCode2
      _BarCode3 = BarCode3
      _CreateTime = CreateTime
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub

  '物件結束時觸發的事件，用來清除物件的內容
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub
  Private Sub Class_Terminate_Renamed()
    '目的:結束物件
  End Sub

  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal PlatForm As String, ByVal LotNo As String, ByVal Store_ID As String) As String
    Try
      Dim key As String = PlatForm & LinkKey & LotNo & LinkKey & Store_ID
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsSTORE_ITEM
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  '取得要Insert的SQL
  Public Function O_Add_Insert_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = STORE_ITEMManagement.GetInsertSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Update的SQL
  Public Function O_Add_Update_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = STORE_ITEMManagement.GetUpdateSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要Delete的SQL
  Public Function O_Add_Delete_SQLString(ByRef lstSQL As List(Of String)) As Boolean
    Try
      Dim strSQL As String = STORE_ITEMManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function


End Class
