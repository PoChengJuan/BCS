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

  Public Function SQLCorrect(ByVal SQL As String, ByRef NewSQL As String) As Boolean
    Try
      Dim NewStr = ""
      Dim whereFlag = False
      For Each splitstr In SQL.Split(" ")
        If whereFlag Then '轉
          If splitstr.Contains("S_CONTAINER") = False Then
            NewStr += splitstr.Replace("=''", " is null ") & " "
          Else
            NewStr += splitstr & ""
          End If
        Else '不用轉
          NewStr += splitstr & " "
        End If
        '找到where where後的=''轉成 not null
        If splitstr.ToUpper() = "WHERE" Then
          whereFlag = True
        End If
      Next

      NewStr = NewStr.Replace("'NULL'", "NULL")
      NewStr = NewStr.Replace("'Null'", "NULL")
      NewStr = NewStr.Replace("'null'", "NULL")


      NewSQL = NewStr

      Return True
    Catch ex As Exception
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
