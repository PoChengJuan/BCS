Public Class FormAbout
  Private Shared dicFormType As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
  Sub New()
    Try
      InitializeComponent()
      '設定Form的標頭
      Me.lb_Version.Text = gVersion.ToString

    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Public Shared Function CreateForm(ByVal FormName As String) As FormAbout
    Try
      '當該MessageName的畫面被開啟時把該MessageName記錄起來，防止被重覆開啟
      If dicFormType.ContainsKey(FormName) Then
        Dim newForm As FormAbout = CType(dicFormType.Item(FormName), FormAbout)
        newForm.Focus()
        Return newForm
      Else
        Dim newForm As FormAbout = New FormAbout()
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