Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports ThoughtWorks.QRCode.Codec
Imports ThoughtWorks.QRCode.Codec.Data
Imports DataMatrix.net
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports GenCode128
Imports ZXing


Module Mod_HelpFunc
#Region "Boolean轉換"
  ''' <summary>
  ''' 處理數字轉Boolean
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  Public Function IntegerConvertToBoolean(ByVal value As String) As Boolean
    Try
      '-檢查是否為數字
      If IsNumeric(value) = False Then
        SendMessageToLog(value & "非數字", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      '-若不為0則為真 '-鬆的規範
      If Convert.ToInt32(value) = 0 Then
        Return False
      End If
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理數字轉Boolean
  ''' </summary>
  ''' <param name="bool"></param>
  ''' <returns></returns>
  Public Function BooleanConvertToInteger(ByVal bool As Boolean) As Integer
    Try
      If bool Then
        Return 1
      Else
        Return 0
      End If
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return 0
    End Try
  End Function
#End Region
  ''' <summary>
  ''' 傳入要更新的欄位和值，回傳要更新的欄位字串
  ''' </summary>
  ''' <param name="dicChangeColumnValue"></param>
  ''' <param name="ret_strUpdateColumnValue"></param>
  ''' <returns></returns>
  Public Function O_Get_UpdateColumnSQL(Of IdxColumnName)(ByRef dicChangeColumnValue As Dictionary(Of String, String),
                                                          ByRef ret_strUpdateColumnValue As String) As Boolean
    Try
      '取得可以寫入SQL的所有欄位
      Dim lstColumnName As List(Of String) = [Enum].GetNames(GetType(IdxColumnName)).ToList
      For Each Column As String In dicChangeColumnValue.Keys
        If lstColumnName.Contains(Column) = True Then
          Dim Value = dicChangeColumnValue.Item(Column)
          '取得寫入DB的欄位和值
          If ret_strUpdateColumnValue = "" Then
            ret_strUpdateColumnValue = String.Format("{0}='{1}'", Column, ReplaceStringForSQL(Value))
          Else
            ret_strUpdateColumnValue = String.Format("{0},{1}='{2}'", ret_strUpdateColumnValue, Column, ReplaceStringForSQL(Value))
          End If
        End If
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
#Region "處理時間"
  ''' <summary>
  '''  取得現在的時間，並自動轉成傳入的指定格式
  ''' </summary>
  ''' <param name="DateTimeFormat"></param>
  ''' <returns></returns>
  Public Function GetNewTime_ByDataTimeFormat(ByVal DateTimeFormat As String) As String
    Try
      Return Now.ToString(DateTimeFormat)
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return ""
    End Try
  End Function
#End Region
  '以下應該都是用不到要取代掉的
  ''' <summary>
  ''' 修改Insert字串中包含特殊符號的部份、將字串單引號改為雙引號，用於寫入資料庫欄位時的修正  
  ''' </summary>
  ''' <param name="name"></param>
  ''' <returns></returns>
  Public Function ReplaceStringForSQL(ByVal name As String) As String
    Try
      name = name.Replace("'", "''")
      'name = name.Replace(" & ", "'||'&'||'") SQL Server不用加
      Return name

    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
End Module
