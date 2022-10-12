Public Class FormMsg
  Private Shared dicFormType As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

  Sub New(ByVal Msg_Str As String)
    Try
      InitializeComponent()
      '設定Form的標頭
      Me.lb_ItemCnt_Str.Text = Msg_Str.ToString

    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Shared Function CreateForm(ByVal FormName As String, ByVal Msg_Str As String) As FormMsg
    Try
      '當該MessageName的畫面被開啟時把該MessageName記錄起來，防止被重覆開啟
      If dicFormType.ContainsKey(FormName) Then
        Dim newForm As FormMsg = CType(dicFormType.Item(FormName), FormMsg)
        newForm.lb_ItemCnt_Str.Text = Msg_Str
        newForm.Focus()
        Return newForm
      Else
        Dim newForm As FormMsg = New FormMsg(Msg_Str)
        dicFormType.Add(FormName, newForm)
        Return newForm
      End If


    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function
  Private Sub FrmMessageTest_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
    Try
      If dicFormType.ContainsKey(Me.Name) Then
        dicFormType.Remove(Me.Name)
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
End Class