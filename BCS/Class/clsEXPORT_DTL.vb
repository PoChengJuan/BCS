Public Class clsEXPORT_DTL
  Private _gid As String
  Private _EXPORT_ID As String '匯出的編號(每一筆任務，一個流水號)
  Private _TABLE_INDEX As Double '使用的範本文件的第幾個Table(從1開始)
  Private _COLUMN_INDEX As Double '使用的範本文件的第幾個Column(從1開始)
  Private _ROW_INDEX As Double '使用的範本文件的第幾個Row(從1開始)
  Private _VALUE As String '寫入欄位的資料
  Private Const LinkKey As String = "*"
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
  Public Property TABLE_INDEX() As Double
    Get
      Return _TABLE_INDEX
    End Get
    Set(ByVal value As Double)
      _TABLE_INDEX = value
    End Set
  End Property
  Public Property COLUMN_INDEX() As Double
    Get
      Return _COLUMN_INDEX
    End Get
    Set(ByVal value As Double)
      _COLUMN_INDEX = value
    End Set
  End Property
  Public Property ROW_INDEX() As Double
    Get
      Return _ROW_INDEX
    End Get
    Set(ByVal value As Double)
      _ROW_INDEX = value
    End Set
  End Property
  Public Property VALUE() As String
    Get
      Return _VALUE
    End Get
    Set(ByVal value As String)
      _VALUE = value
    End Set
  End Property


  Public Sub New(ByVal EXPORT_ID As String, ByVal TABLE_INDEX As Double, ByVal COLUMN_INDEX As Double, ByVal ROW_INDEX As Double, ByVal VALUE As String)
    MyBase.New()
    Try
      Dim key As String = Get_Combination_Key(EXPORT_ID, TABLE_INDEX, COLUMN_INDEX, ROW_INDEX)
      _gid = key
      _EXPORT_ID = EXPORT_ID
      _TABLE_INDEX = TABLE_INDEX
      _COLUMN_INDEX = COLUMN_INDEX
      _ROW_INDEX = ROW_INDEX
      _VALUE = VALUE
    Catch ex As Exception
    End Try
  End Sub
  '=================Public Function=======================
  '傳入指定參數取得Key值
  Public Shared Function Get_Combination_Key(ByVal EXPORT_ID As String, ByVal TABLE_INDEX As Double, ByVal COLUMN_INDEX As Double, ByVal ROW_INDEX As Double) As String
    Try
      Dim key As String = EXPORT_ID & LinkKey & TABLE_INDEX & LinkKey & COLUMN_INDEX & LinkKey & ROW_INDEX
      Return key
    Catch ex As Exception
      Return ""
    End Try
  End Function
  Public Function Clone() As clsEXPORT_DTL
    Try
      Return Me.MemberwiseClone()
    Catch ex As Exception
      Return Nothing
    End Try
  End Function
End Class
