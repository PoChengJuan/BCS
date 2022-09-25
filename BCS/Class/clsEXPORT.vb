Public Class clsEXPORT
  Private _gid As String
  Private _EXPORT_ID As String '匯出的編號(每一筆任務，一個流水號)
  Private _SAMPLE_FILE_NAME As String '匯出的範本文件名稱
  Private _CREATE_TIME As String '建立時間
  Private _FINISH_TIME As String '完成時間
  Private _EXPORT_TYPE As Double '匯出的類型1:匯出2:列印
  Private _PRINTER_NO As String '如果匯出的類型是列印，需要填入印表機編號

  Public Property gid() As String
    Get
      Return _gid
    End Get
    Set(ByVal value As String)
      _gid = value
    End Set
  End Property
  Public Property EXPORT_ID() As String
    Get
      Return _EXPORT_ID

    End Get
    Set(ByVal value As String)
      _EXPORT_ID = value
    End Set
  End Property
  Public Property SAMPLE_FILE_NAME() As String
    Get
      Return _SAMPLE_FILE_NAME
    End Get
    Set(ByVal value As String)
      _SAMPLE_FILE_NAME = value
    End Set
  End Property
  Public Property CREATE_TIME() As String
    Get
      Return _CREATE_TIME
    End Get
    Set(ByVal value As String)
      _CREATE_TIME = value
    End Set
  End Property
  Public Property FINISH_TIME() As String
    Get
      Return _FINISH_TIME
    End Get
    Set(ByVal value As String)
      _FINISH_TIME = value
    End Set
  End Property
  Public Property EXPORT_TYPE() As Double
    Get
      Return _EXPORT_TYPE
    End Get
    Set(ByVal value As Double)
      _EXPORT_TYPE = value
    End Set
  End Property
  Public Property PRINTER_NO() As String
    Get
      Return _PRINTER_NO
    End Get
    Set(ByVal value As String)
      _PRINTER_NO = value
    End Set
  End Property
	Public Sub New(ByVal EXPORT_ID As String, ByVal SAMPLE_FILE_NAME As String, ByVal CREATE_TIME As String, ByVal FINISH_TIME As String,
								 ByVal EXPORT_TYPE As Double,
								 ByVal PRINTER_NO As String)
		MyBase.New()
		Try
			Dim key As String = Get_Combination_Key(EXPORT_ID)
			_gid = key
			_EXPORT_ID = EXPORT_ID
			_SAMPLE_FILE_NAME = SAMPLE_FILE_NAME
			_CREATE_TIME = CREATE_TIME
			_FINISH_TIME = FINISH_TIME
			_EXPORT_TYPE = EXPORT_TYPE
			_PRINTER_NO = PRINTER_NO
		Catch ex As Exception
		End Try
	End Sub
	'=================Public Function=======================
	'傳入指定參數取得Key值
	Public Shared Function Get_Combination_Key(ByVal EXPORT_ID As String) As String
    Try
      Dim key As String = EXPORT_ID
      Return key
    Catch ex As Exception
      Return ""
    End Try
  End Function
  Public Function Clone() As clsEXPORT
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      Return Nothing
    End Try
  End Function
End Class
