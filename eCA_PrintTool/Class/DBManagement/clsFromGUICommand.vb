Public Class clsFromGUICommand
  Private ShareName As String = "FromGUICommand"
  Private ShareKey As String = ""
	Private Const LinkKey As String = "*"
	Private _gid As String
  Private _UUID As String
  Private _Send_System As enuSystemType
  Private _Receive_System As enuSystemType
  Private _Function_ID As String
  Private _SEQ As Long
  Private _User_ID As String
  Private _Create_Time As String
  Private _Message As String
  Private _Result As String
  Private _Result_Message As String
	Private _Wait_UUID As String
	Private _CLIENT_ID As String
	Private _IP As String

	Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property UUID() As String
    Get
      Return _UUID
    End Get
    Set(ByVal value As String)
      _UUID = value
    End Set
  End Property
  Public Property Send_System() As enuSystemType
    Get
      Return _Send_System
    End Get
    Set(ByVal value As enuSystemType)
      _Send_System = value
    End Set
  End Property
  Public Property Receive_System() As enuSystemType
    Get
      Return _Receive_System
    End Get
    Set(ByVal value As enuSystemType)
      _Receive_System = value
    End Set
  End Property
  Public Property Function_ID() As String
    Get
      Return _Function_ID
    End Get
    Set(ByVal value As String)
      _Function_ID = value
    End Set
  End Property
  Public Property SEQ() As Long
    Get
      Return _SEQ
    End Get
    Set(ByVal value As Long)
      _SEQ = value
    End Set
  End Property
  Public Property User_ID() As String
    Get
      Return _User_ID
    End Get
    Set(ByVal value As String)
      _User_ID = value
    End Set
  End Property
  Public Property Create_Time() As String
    Get
      Return _Create_Time
    End Get
    Set(ByVal value As String)
      _Create_Time = value
    End Set
  End Property
  Public Property Message() As String
    Get
      Return _Message
    End Get
    Set(ByVal value As String)
      _Message = value
    End Set
  End Property
  Public Property Result() As String
    Get
      Return _Result
    End Get
    Set(ByVal value As String)
      _Result = value
    End Set
  End Property
  Public Property Result_Message() As String
    Get
      Return _Result_Message
    End Get
    Set(ByVal value As String)
      _Result_Message = value
    End Set
  End Property
	Public Property Wait_UUID() As String
		Get
			Return _Wait_UUID
		End Get
		Set(ByVal value As String)
			_Wait_UUID = value
		End Set
	End Property

	Public Property CLIENT_ID() As String
		Get
			Return _CLIENT_ID
		End Get
		Set(ByVal value As String)
			_CLIENT_ID = value
		End Set
	End Property
	Public Property IP() As String
		Get
			Return _IP
		End Get
		Set(ByVal value As String)
			_IP = value
		End Set
	End Property



	'物件建立時執行的事件
	Public Sub New(ByVal UUID As String,
								 ByVal Send_System As EnuSystemType,
								 ByVal Receive_System As EnuSystemType,
								 ByVal Function_ID As String,
								 ByVal SEQ As Long,
								 ByVal User_ID As String,
								 ByVal Create_Time As String,
								 ByVal Message As String,
								 ByVal Result As String,
								 ByVal Result_Message As String,
								 ByVal Wait_UUID As String,
								 ByVal CLIENT_ID As String,
								 ByVal IP As String)
		MyBase.New()
		Try
			Dim key As String = Get_Combination_Key(UUID, Function_ID, SEQ)
			_gid = key
			_UUID = UUID
			_Send_System = Send_System
			_Receive_System = Receive_System
			_Function_ID = Function_ID
			_SEQ = SEQ
			_User_ID = User_ID
			_Create_Time = Create_Time
			_Message = Message
			_Result = Result
			_Result_Message = Result_Message
			_Wait_UUID = Wait_UUID
			_CLIENT_ID = CLIENT_ID
			_IP = IP
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
  Public Shared Function Get_Combination_Key(ByVal UUID As String,
                                             ByVal Function_ID As String,
                                             ByVal SEQ As Long) As String
    Try
      Dim key As String = UUID & LinkKey & Function_ID & LinkKey & SEQ
      Return key
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
  Public Function Clone() As clsFromGUICommand
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
      Dim strSQL As String = WMS_T_GUI_CommandManagement.GetInsertSQL(Me)
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
      Dim strSQL As String = WMS_T_GUI_CommandManagement.GetUpdateSQL(Me)
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
      Dim strSQL As String = WMS_T_GUI_CommandManagement.GetDeleteSQL(Me)
      lstSQL.Add(strSQL)
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
End Class
