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


Module ModuleFunction
  ''' <summary>
  ''' 處理 ='' 轉 is null(針對Oracle和SQLServer有不同的調整方式)
  ''' </summary>
  ''' <param name="DB_Type"></param>
  ''' <param name="lstSQL"></param>
  ''' <param name="NewSQL"></param>
  ''' <returns></returns>
  Public Function SQLCorrect(ByRef DB_Type As Short, ByVal lstSQL As List(Of String), ByRef NewSQL As List(Of String)) As Boolean
    Try
      Dim str As String
      For Each str In lstSQL
        Dim NewStr = ""
        SQLCorrect(DB_Type, str, NewStr)
        NewSQL.Add(NewStr)
      Next
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''' <summary>
  ''' 處理 ='' 轉 is null(針對Oracle和SQLServer有不同的調整方式)
  ''' </summary>
  ''' <param name="SQL"></param>
  ''' <param name="NewSQL"></param>
  ''' <returns></returns>
  Public Function SQLCorrect(ByRef DB_Type As Short, ByVal SQL As String, ByRef NewSQL As String) As Boolean
    Try
      '只進行Where條件部份的轉換
      Dim NewStr = ""
      Dim whereFlag = False
      Select Case DB_Type
        Case 0  'Oracle
          '把WITH(NOLOCK)取代掉
          SQL = SQL.Replace("WITH(NOLOCK)", "")
          '把Where條件中有=''和<>''的取代掉
          For Each splitstr In SQL.Split(" ")
            If whereFlag Then '轉
              If splitstr.IndexOf("=''") <> -1 Then
                NewStr += splitstr.Replace("=''", " is null") & " "
              ElseIf splitstr.IndexOf("<>''") <> -1 Then
                NewStr += splitstr.Replace("<>''", " is not null") & " "
              Else
                NewStr += splitstr & " "
              End If
            Else '不用轉
              NewStr += splitstr & " "
            End If
            '找到where where後的=''轉成 not null
            If splitstr.ToUpper() = "WHERE" Then
              whereFlag = True
            End If
          Next
        Case 1  'SQL Server
          'SQL Server不用進行轉換
          NewStr = SQL
      End Select
      NewSQL = NewStr
      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
