Imports System.Windows.Forms

Public Class FormShop
  Private Shared dicFormType As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

  Public Shared Function CreateForm(ByVal FormName As String) As FormShop
    Try
      '當該MessageName的畫面被開啟時把該MessageName記錄起來，防止被重覆開啟
      If dicFormType.ContainsKey(FormName) Then
        Dim newForm As FormShop = CType(dicFormType.Item(FormName), FormShop)
        newForm.Focus()
        Return newForm
      Else
        Dim newForm As FormShop = New FormShop()
        dicFormType.Add(FormName, newForm)
        Return newForm
      End If


    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return Nothing
    End Try
  End Function

  Private Sub btn_SHOP_SAVE_Click(sender As Object, e As EventArgs) Handles btn_SHOP_SAVE.Click
    Try
      'Dim _form = FormShop.CreateForm(Me.Name)
      'If _form IsNot Nothing Then
      '  _form.Close()
      'End If
      Dim SHOP_NO = tb_SHOP_NO.Text.ToUpper
      Dim SHOP_MEMO = tb_SHOP_MEMO.Text
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim lstSQL As New List(Of String)
      Dim strResultMsg = ""

      '檢查
      If SHOP_NO = "" Then
        strResultMsg = "請輸入賣家編號"
        MsgBox(strResultMsg)
      End If
      '資料
      Dim objSHOP As clsBCS_M_SHOP = Nothing
      Dim objUpdateSHOP As clsBCS_M_SHOP = Nothing
      Dim dicSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY(SHOP_NO)
      If dicSHOP.Any Then
        objUpdateSHOP = dicSHOP.First.Value.Clone
      Else
        objSHOP = New clsBCS_M_SHOP(SHOP_NO, SHOP_MEMO, Now_Time)
      End If

      '取得SQL
      If objSHOP IsNot Nothing Then
        If objSHOP.O_Add_Insert_SQLString(lstSQL) = False Then
          strResultMsg = "Get Insert SHOP SQL Failed"
          MsgBox(strResultMsg)
          Return
        End If
      ElseIf objUpdateSHOP IsNot Nothing Then
        objUpdateSHOP.MEMO = tb_SHOP_MEMO.Text
        If objUpdateSHOP.O_Add_Update_SQLString(lstSQL) = False Then
          strResultMsg = "Get Update SHOP SQL Failed"
          MsgBox(strResultMsg)
          Return
        End If
      End If


      '寫入DB
      If Common_DBManagement.BatchUpdate(lstSQL) = False Then
        '更新DB失敗則回傳False
        'ret_strResultMsg = "WMS Update DB Failed"
        strResultMsg = "網路連線失敗"
        MsgBox(strResultMsg)
        Return
      Else
        FrmMain.CB_BarCode_SHOP.Items.Clear()
        FrmMain.CB_BarCode_SHOP.ResetText()
        Dim dicNewSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY("")
        For Each obj In dicNewSHOP.Values
          FrmMain.CB_BarCode_SHOP.Items.Add(obj.SHOP)
        Next
        FrmMain.CB_Report_SHOP.Items.Clear()
        FrmMain.CB_Report_SHOP.ResetText()

        For Each objSHOP In dicNewSHOP.Values
          FrmMain.CB_Report_SHOP.Items.Add(objSHOP.SHOP)
        Next
        MsgBox("儲存成功")
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)

    End Try
  End Sub
  Private Sub btn_SHOP_DELETE_Click(sender As Object, e As EventArgs) Handles btn_SHOP_DELETE.Click
    Dim dicDeleteSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY(tb_SHOP_NO.Text)
    Dim lstSQL As New List(Of String)
    Dim strResultMsg = ""

    Dim SHOP_NO = tb_SHOP_NO.Text

    If SHOP_NO = "" Then
      strResultMsg = "未輸入賣家編號"
      MsgBox(strResultMsg)
      Return
    End If
    For Each obj In dicDeleteSHOP.Values
      If obj.O_Add_Delete_SQLString(lstSQL) = False Then
        strResultMsg = "Get Delete SHOP SQL Failed"
        MsgBox(strResultMsg)
        Return
      End If
    Next

    If Common_DBManagement.BatchUpdate(lstSQL) = False Then
      '更新DB失敗則回傳False
      'ret_strResultMsg = "WMS Update DB Failed"
      strResultMsg = "網路連線失敗"
      MsgBox(strResultMsg)
      Return
    Else
      FrmMain.CB_BarCode_SHOP.Items.Clear()
      FrmMain.CB_BarCode_SHOP.ResetText()
      Dim dicNewSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY("")
      For Each obj In dicNewSHOP.Values
        FrmMain.CB_BarCode_SHOP.Items.Add(obj.SHOP)
      Next

      FrmMain.CB_Report_SHOP.Items.Clear()
      FrmMain.CB_Report_SHOP.ResetText()

      For Each objSHOP In dicNewSHOP.Values
        FrmMain.CB_Report_SHOP.Items.Add(objSHOP.SHOP)
      Next

      MsgBox("刪除成功")
    End If
  End Sub
  Private Sub FrmMessageTest_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
    Try
      If dicFormType.ContainsKey(Me.Name) Then
        dicFormType.Remove(Me.Name)
      End If
    Catch ex As Exception
      MsgBox(ex.ToString)
      'SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
    End Try
  End Sub
  Private Sub tb_SHOP_NO_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tb_SHOP_NO.KeyPress
    Try
      If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
        Dim dicSHOP = BCS_M_SHOPManagement.GetDataDictionaryByKEY(tb_SHOP_NO.Text)
        If dicSHOP.Any Then
          tb_SHOP_MEMO.Text = dicSHOP.First.Value.MEMO
        Else
          tb_SHOP_MEMO.Text = ""
        End If
      End If

    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
End Class