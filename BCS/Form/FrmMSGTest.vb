Public Class FrmMSGTest
	Public MessageName As EnuMSGFunctionID
	'判別是否被開啟過的參數
	Private Shared dicFormType As Dictionary(Of EnuMSGFunctionID, Object) = New Dictionary(Of EnuMSGFunctionID, Object)
	Private Sub FrmMSGTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
	Me.Height = 500 + gb_ResultMessage.Height + 6
  End Sub
	Sub New(ByVal New_MessageName As EnuMSGFunctionID)
		Try
			MessageName = New_MessageName
			InitializeComponent()
			'設定Form的標頭
			Me.Text = MessageName.ToString
		Catch ex As Exception
			MsgBox(ex.ToString)
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
		End Try
	End Sub
	Private Sub btn_SendMessage_Click(sender As Object, e As EventArgs) Handles btn_SendMessage.Click
		Try
			Dim strXMLMsg As String = txt_SendMessage.Text
			Dim ret_ResultMsg As String = ""
			Dim strLog As String = String.Format("XML String ={0}", strXMLMsg)
			SendMessageToLog(strLog, eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			Select Case MessageName
				Case EnuMSGFunctionID.T10F1S1_PrintCarrierLabel
          If O_ProcessT10F1S1_PrintCarrierLabel(strXMLMsg, ret_ResultMsg) = True Then
            txt_ResultMessage.Text = ret_ResultMsg
          Else
            txt_ResultMessage.Text = ret_ResultMsg
          End If
        Case EnuMSGFunctionID.T10F1S2_PrintItemLabel
          If O_ProcessT10F1S2_PrintItemLabel(strXMLMsg, ret_ResultMsg) = True Then
            txt_ResultMessage.Text = ret_ResultMsg
          Else
            txt_ResultMessage.Text = ret_ResultMsg
          End If
        Case EnuMSGFunctionID.T10F1S21_PrintShippingDTL
          If O_ProcessT10F1S21_PrintShippingDTL(strXMLMsg, ret_ResultMsg) = True Then
            txt_ResultMessage.Text = ret_ResultMsg
          Else
            txt_ResultMessage.Text = ret_ResultMsg
          End If
      End Select


		Catch ex As Exception
			MsgBox(ex.ToString)
		End Try
	End Sub



	Public Shared Function CreateForm(ByVal MessageName As EnuMSGFunctionID) As FrmMSGTest
		Try
			'當該MessageName的畫面被開啟時把該MessageName記錄起來，防止被重覆開啟
			If dicFormType.ContainsKey(MessageName) Then
				Dim newForm As FrmMSGTest = CType(dicFormType.Item(MessageName), FrmMSGTest)
				newForm.Focus()
				Return newForm
			Else
				Dim newForm As FrmMSGTest = New FrmMSGTest(MessageName)
				dicFormType.Add(MessageName, newForm)
				Return newForm
			End If
		Catch ex As Exception
			MsgBox(ex.ToString)
			SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
			Return Nothing
		End Try
	End Function

  Private Sub FrmMSGTest_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
	dicFormType.Remove(MessageName)
	Me.Dispose()
  End Sub
End Class