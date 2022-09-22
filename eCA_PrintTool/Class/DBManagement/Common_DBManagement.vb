Public Class Common_DBManagement
  Public Shared DBTool As eCA_DBTool.clsDBTool = Nothing
  Public Sub New()

  End Sub

  Public Shared Function BatchUpdate(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return False
      If lstSQL.Count = 0 Then Return True
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      Dim rtnMsg As String = DBTool.BatchUpdate(lstSQL)
      If rtnMsg.StartsWith("OK") Then
        SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Else
        SendMessageToLog(rtnMsg, eCALogTool.ILogTool.enuTrcLevel.lvError)
        Return False
      End If
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Public Shared Function AddQueued(ByVal strSQL As String) As Boolean
    Try
      If strSQL.Length = 0 Then Return -1

      SendMessageToLog("SQL:" & strSQL, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)


      DBTool.O_AddQueue_File((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, strSQL)

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Public Shared Function AddQueued(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return -1
      If lstSQL.Count = 0 Then Return 0
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      'DBTool.BatchUpdateToFile((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, lstSQL)
      DBTool.O_AddQueue_File((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, lstSQL)

      'For i = 0 To lstSQL.Count - 1
      '  DBTool.O_AddQueue_File((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, lstSQL(i))
      'Next
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Public Shared Function AddQueued_BatchUpdate(ByRef lstSQL As List(Of String)) As Boolean
    Try
      If lstSQL Is Nothing Then Return -1
      If lstSQL.Count = 0 Then Return 0
      For i = 0 To lstSQL.Count - 1
        SendMessageToLog("SQL:" & lstSQL(i), eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next
      DBTool.BatchUpdateToFile((New System.Diagnostics.StackTrace).GetFrame(1).GetMethod.Name, lstSQL)
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Public Shared Sub O_thr_Write_QueueLog()
    Const SleepTime As Integer = 500
    While True
      Try
        Dim ResultQueue = DBTool.gSQLResultQueue
        While ResultQueue.Count > 0
          Dim Queue = ResultQueue.Dequeue
          'If Queue.sSQLCommandKey.Equals("SYSTEM_UID") Or Queue.sSQLCommandKey.Equals("WMS_SYSTEM_STATUS") Then '-排除一直更新的DB 可改為config設定
          '  'Continue For
          'Else
          Dim str = "[Queue] Log:" & Queue.sLog
          SendMessageToLog(str, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
          'End If
        End While
      Catch ex As Exception
        SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Finally
        Threading.Thread.Sleep(SleepTime)
      End Try
    End While
    SendMessageToLog("O_thr_Write_QueueLog End", eCALogTool.ILogTool.enuTrcLevel.lvWARN)
  End Sub
End Class
