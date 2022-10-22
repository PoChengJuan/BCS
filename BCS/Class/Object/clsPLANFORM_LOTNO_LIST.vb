 Public Class clsPLANFORM_LOTNO_LIST
 Private ShareName As String = "PLANFORM_LOTNO_LIST"
 Private ShareKey As String = ""
 Private _gid As String
 Private _PlanForm As String '平台 
 Private _LotNo As String '賣場 
 Private _Count As Long '數量 
  
  
 Public Property gid() As String
 Get
 Return _gid
 End Get
 Set(ByVal value As String)
 _gid = value
 End Set
 End Property
 Public Property PlanForm() As String
 Get
 Return _PlanForm
 End Get
 Set(ByVal value As String)
 _PlanForm = value
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
 Public Property Count() As Long
 Get
 Return _Count
 End Get
 Set(ByVal value As Long)
 _Count = value
 End Set
 End Property
  
  
 '物件建立時執行的事件
 Public Sub New( ByVal PlanForm As String, ByVal LotNo As String, ByVal Count As Long)
 MyBase.New()
 Try
 Dim key As String = Get_Combination_Key(PlanForm, LotNo)
 _gid = key
 _PlanForm = PlanForm
 _LotNo = LotNo
 _Count = Count
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
 Public Shared Function Get_Combination_Key( ByVal PlanForm As String, ByVal LotNo As String) As String
 Try
 Dim key As String = PlanForm & LinkKey & LotNo
 Return key
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return ""
 End Try
 End Function
 Public Function Clone() As clsPLANFORM_LOTNO_LIST
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
 Dim strSQL As String = PLANFORM_LOTNO_LISTManagement.GetInsertSQL(Me)
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
 Dim strSQL As String = PLANFORM_LOTNO_LISTManagement.GetUpdateSQL(Me)
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
 Dim strSQL As String = PLANFORM_LOTNO_LISTManagement.GetDeleteSQL(Me)
 lstSQL.Add(strSQL)
 Return True
 Catch ex As Exception
 SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
 Return False
 End Try
 End Function
  
  
 End Class
