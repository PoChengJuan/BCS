
Imports System.Drawing
Imports GenCode128
Imports System.Diagnostics
Imports System.IO
Imports NPOI.XSSF.UserModel
Imports NPOI.XWPF.Usermodel
Imports NPOI.SS.UserModel

Module Module_CreateBarCodeReport
  Public Function O_Process_Message(ByVal PlatForm As String,
                                    ByVal LotNo As String,
                                    ByVal Start_time As String,
                                    ByVal End_Time As String,
                                    ByRef ret_FileName_PDF As String,
                                    ByRef ret_strResultMsg As String) As Boolean
    Try

      Dim dicStore_Item As New Dictionary(Of String, clsSTORE_ITEM)
      Dim dicAddStore_Head As New Dictionary(Of String, clsSTORE_HEAD)
      Dim dicAddStore_Item As New Dictionary(Of String, clsSTORE_ITEM)

      Dim lstSQL As New List(Of String)
      Dim lstLCSSql As New List(Of String)
      Dim lstHistroySQL As New List(Of String)

      Dim File_Name = ""
      Dim File_Path = ""
      '檢查資料
      If Check_Data(PlatForm, LotNo, Start_time, End_Time, dicStore_Item, ret_strResultMsg) = False Then
        Return False
      End If
      ''取得更新資料
      If Get_Data(PlatForm, LotNo, dicStore_Item, File_Path, File_Name, ret_strResultMsg) = False Then
        Return False
      End If
      ''取得要更新到DB的SQL
      If Get_SQL(ret_strResultMsg, dicAddStore_Head, dicAddStore_Item, lstSQL) = False Then
        Return False
      End If
      ''執行資料更新
      If Execute_DataUpdate(ret_strResultMsg, lstSQL) = False Then
        Return False
      End If
      Dim File_Name_PDF = File_Name.Replace("xlsx", "pdf")
      'Dim ret_FileName_PDF = ""
      Dim ret_MSG = ""
      If eCA_Excel.ExcelToPDF2(File_Name, ret_FileName_PDF, ret_MSG) = False Then
        ret_strResultMsg += " 轉pdf失败"
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  ''檢查資料
  Private Function Check_Data(ByRef ret_PlatForm As String,
                              ByRef ret_LotNo As String,
                              ByRef ret_StartTime As String,
                              ByRef ret_EndTime As String,
                              ByRef ret_dicStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                              ByRef ret_strResultMsg As String) As Boolean
    Try
      ret_dicStore_Item = STORE_ITEMManagement.GetDataDictionaryByItemReport(ret_PlatForm, ret_LotNo, ret_StartTime, ret_EndTime)
      If ret_dicStore_Item.Any = False Then
        ret_strResultMsg = "查無條碼"
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  ''取得更新資料
  Private Function Get_Data(ByVal ret_PlatForm As String,
                            ByVal ret_LotNo As String,
                            ByVal ret_dicStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                            ByRef ret_File_Path As String,
                            ByRef ret_File_Name As String,
                            ByRef ret_strResultMsg As String) As Boolean
    Try
      Dim Now_Time As String = GetNewTime_DBFormat()
      Dim Now_Date As String = GetNewDate_DBFormat()

      '取得檔名漂水號
      Dim UUID_NO_Str = ret_PlatForm
      Dim UUID_NO = ""
      If ret_PlatForm = "7-11BarCode" Then
        UUID_NO = enuUUID_No.Seven_BarCode_SERIAL_NO.ToString
      ElseIf ret_PlatForm = "7-11QRCode" Then
        UUID_NO = enuUUID_No.Seven_QRCode_SERIAL_NO.ToString
      ElseIf ret_PlatForm = "OK Mart" Then
        UUID_NO = enuUUID_No.OK_SERIAL_NO.ToString
      ElseIf ret_PlatForm = "Family" Then
        UUID_NO = enuUUID_No.FAMILY_SERIAL_NO.ToString
      End If
      Dim dicUUID As Dictionary(Of String, clsUUID) = BCS_M_UUIDManagement.GetclsUUIDListByUUID_NO(UUID_NO)


      'If objHandling.O_Get_UUID(enuUUID_No.Seven_SERIAL_NO.ToString, dicUUID) = False Then
      '  ret_strResultMsg = "Get UUID False"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      If dicUUID.Any = False Then
        ret_strResultMsg = "Get UUID False"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If
      Dim objUUID = dicUUID.Values(0)
      Dim UUID = objUUID.Get_NewUUID
      Dim FileName = UUID_NO_Str & "-" & ret_LotNo & "-" & UUID & ".xlsx"
      Dim FilePath = gFileRootPath & Now_Date & "\" & UUID_NO_Str & "\"
      ret_File_Name = FilePath & FileName
      ret_File_Path = FilePath
      If System.IO.Directory.Exists(ret_File_Path) = False Then
        System.IO.Directory.CreateDirectory(ret_File_Path)
      End If
      Dim fs = New FileStream(FilePath & FileName, FileMode.Create)
      Dim workbook As XSSFWorkbook = New XSSFWorkbook()
      Dim sheet As XSSFSheet = workbook.CreateSheet("Sheet1") ' 新增試算表 Sheet名稱
      Dim footer = sheet.Footer
      Dim Header = sheet.Header
      'footer.Center = "第" & "&P" & "頁"
      Dim xlStyle As XSSFCellStyle = workbook.CreateCellStyle()

      Dim Row_Cnt = 0
      Dim Row As XSSFRow = sheet.CreateRow(Row_Cnt)

      'Row.CreateCell(0).SetCellValue("平台：" & ret_PlatForm)
      'Row.CreateCell(3).SetCellValue("賣場(Lot)：" & ret_LotNo)
      Header.Left = "平台：" & ret_PlatForm & "  賣場(Lot)：" & ret_LotNo
      Dim cnt = 0
      'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
      '  Return False
      'End If
      Dim lst_dicStore_Item As New List(Of clsSTORE_ITEM)
      For Each obj In ret_dicStore_Item.Values
        lst_dicStore_Item.Add(obj)
      Next


      Dim flg_lastOne = False
      Dim flg_Change_Page = False
      Dim Total_BarCode = 0
      Dim BarCode_Cnt = 0
      Dim page_cnt = 2
      Dim flg_Change_Sheet = False
      SendMessageToLog("開始時間：" & Now_Time, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      For Item_Cnt As Integer = 0 To lst_dicStore_Item.Count - 1 Step +2

        'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "3") = False Then
        '  Return False
        'End If
        Row_Cnt = Row_Cnt + 1
        Row = sheet.CreateRow(Row_Cnt)

        If ret_PlatForm = "7-11BarCode" Then
          If (lst_dicStore_Item.Count - Item_Cnt) < 2 Then
            flg_lastOne = True
          End If

#Region "7-11BarCode"
#Region "BarCode1"
          '=============================BarCode1====================================


          CreateBarCode(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper, 0, workbook, sheet, Row_Cnt)

          Row.CreateCell(0).SetCellValue("")
          Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
          Dim colorRgb = New Byte() {255, 255, 255}
          Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Barcode_Style.FillPattern = FillPattern.SolidForeground


          Dim Font As IFont = workbook.CreateFont()
          Font.FontName = "Code 128"
          Font.FontHeightInPoints = 12
          Barcode_Style.SetFont(Font)
          Row.Cells(0).CellStyle = Barcode_Style
          If flg_lastOne = False Then
            '##############################################################
            CreateBarCode(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper, 5, workbook, sheet, Row_Cnt)
            Row.CreateCell(5).SetCellValue("")

            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Barcode_Style.FillPattern = FillPattern.SolidForeground

            Font.FontName = "Code 128"
            Font.FontHeightInPoints = 36
            Barcode_Style.SetFont(Font)
            Row.Cells(1).CellStyle = Barcode_Style
            '##############################################################
          End If

          '=============================================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper)

          Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Num_Style.FillPattern = FillPattern.SolidForeground


          Dim Num_Font As IFont = workbook.CreateFont()
          Num_Font.FontName = "Calibri"
          Num_Font.FontHeightInPoints = 12
          Num_Style.SetFont(Num_Font)
          Row.Cells(0).CellStyle = Num_Style

          If flg_lastOne = False Then
            '##############################################################
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper)

            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground

            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(1).CellStyle = Num_Style
            '##############################################################
          End If

          '=============================================================
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
#End Region
#Region "BarCode2"
          '=============================BarCode2====================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          CreateBarCode(lst_dicStore_Item.Item(Item_Cnt).BarCode2.ToUpper, 0, workbook, sheet, Row_Cnt)

          Row.CreateCell(0).SetCellValue("")
          'Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Barcode_Style.FillPattern = FillPattern.SolidForeground


          'Dim Font As IFont = workbook.CreateFont()
          Font.FontName = "Code 128"
          Font.FontHeightInPoints = 36
          Barcode_Style.SetFont(Font)
          Row.Cells(0).CellStyle = Barcode_Style
          If flg_lastOne = False Then
            If Item_Cnt = 148 Then
              Dim a = 0
            End If
            '##############################################################
            CreateBarCode(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode2.ToUpper, 5, workbook, sheet, Row_Cnt)
            Row.CreateCell(5).SetCellValue("")

            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Barcode_Style.FillPattern = FillPattern.SolidForeground

            Font.FontName = "Code 128"
            Font.FontHeightInPoints = 36
            Barcode_Style.SetFont(Font)
            Row.Cells(1).CellStyle = Barcode_Style
            '##############################################################
          End If

          '=============================================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode2.ToUpper)

          'Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Num_Style.FillPattern = FillPattern.SolidForeground


          'Dim Num_Font As IFont = workbook.CreateFont()
          Num_Font.FontName = "Calibri"
          Num_Font.FontHeightInPoints = 12
          Num_Style.SetFont(Num_Font)
          Row.Cells(0).CellStyle = Num_Style

          If flg_lastOne = False Then
            '##############################################################
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode2.ToUpper)

            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground

            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(1).CellStyle = Num_Style
            '##############################################################
          End If

          '=============================================================
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
#End Region
#Region "BarCode3"
          '=============================BarCode3====================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          CreateBarCode(lst_dicStore_Item.Item(Item_Cnt).BarCode3.ToUpper, 0, workbook, sheet, Row_Cnt)

          Row.CreateCell(0).SetCellValue("")
          'Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Barcode_Style.FillPattern = FillPattern.SolidForeground


          'Dim Font As IFont = workbook.CreateFont()
          Font.FontName = "Code 128"
          Font.FontHeightInPoints = 36
          Barcode_Style.SetFont(Font)
          Row.Cells(0).CellStyle = Barcode_Style
          If flg_lastOne = False Then
            '##############################################################
            CreateBarCode(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode3.ToUpper, 5, workbook, sheet, Row_Cnt)

            Row.CreateCell(5).SetCellValue("")

            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Barcode_Style.FillPattern = FillPattern.SolidForeground

            Font.FontName = "Code 128"
            Font.FontHeightInPoints = 36
            Barcode_Style.SetFont(Font)
            Row.Cells(1).CellStyle = Barcode_Style
            '##############################################################
          End If

          '=============================================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode3.ToUpper)

          'Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Num_Style.FillPattern = FillPattern.SolidForeground


          'Dim Num_Font As IFont = workbook.CreateFont()
          Num_Font.FontName = "Calibri"
          Num_Font.FontHeightInPoints = 12
          Num_Style.SetFont(Num_Font)
          Row.Cells(0).CellStyle = Num_Style

          If flg_lastOne = False Then
            '##############################################################
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode3.ToUpper)

            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground

            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(1).CellStyle = Num_Style
            '##############################################################
          End If

          '=============================================================
          If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            Return False
          End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
#End Region
#Region "BarCode4"
          'Dim _VALUE = GenCode128.Code128Rendering.MakeBarcodeImage(objStore_item.BarCode4, 2, True)
          Dim _VALUE = CodeEncoderFromString(lst_dicStore_Item.Item(Item_Cnt).BarCode4.ToUpper, "QRCODE")
          Dim BYTE_Array As Byte() = BmpToBytes(_VALUE)
          Dim picInd = workbook.AddPicture(BYTE_Array, NPOI.SS.UserModel.PictureType.JPEG)

          Dim anchor As XSSFClientAnchor = New XSSFClientAnchor()
          anchor.Dx1 = 1
          anchor.Dy1 = 1


          anchor.Col1 = 0
          anchor.Row1 = Row_Cnt
          Dim drawing As XSSFDrawing = sheet.CreateDrawingPatriarch
          Dim pict As XSSFPicture = drawing.CreatePicture(anchor, picInd)
          pict.Resize()
          BarCode_Cnt = BarCode_Cnt + 1
          '=============================================================


          If flg_lastOne = False Then
            '##############################################################
            _VALUE = CodeEncoderFromString(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode4.ToUpper, "QRCODE")
            BYTE_Array = BmpToBytes(_VALUE)
            picInd = workbook.AddPicture(BYTE_Array, NPOI.SS.UserModel.PictureType.JPEG)

            anchor = New XSSFClientAnchor()
            anchor.Dx1 = 1
            anchor.Dy1 = 1


            anchor.Col1 = 5
            anchor.Row1 = Row_Cnt
            drawing = sheet.CreateDrawingPatriarch
            pict = drawing.CreatePicture(anchor, picInd)
            pict.Resize()
            BarCode_Cnt = BarCode_Cnt + 1
            '##############################################################
          End If

          If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            Return False
          End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If

#End Region
          '切換Sheet，釋放記憶體
          'If Item_Cnt Mod 46 = 0 And Item_Cnt <> 0 Then
          If BarCode_Cnt >= 6 Then
            Dim sheet_Str = "Sheet" & page_cnt.ToString
            Header = sheet.Header
            Header.Left = "平台：" & ret_PlatForm & "  賣場(Lot)：" & ret_LotNo
            page_cnt = page_cnt + 1
            Row_Cnt = 0
            BarCode_Cnt = 0
            sheet = workbook.CreateSheet(sheet_Str)

            'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "3") = False Then
            '  Return False
            'End If


          Else
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            'If BarCode_Cnt >= 6 Then
            '  Total_BarCode = Total_BarCode + BarCode_Cnt
            '  If lst_dicStore_Item.Count <> Total_BarCode Then  '防止剛好滿頁時，不會再多換一頁
            '    If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '      Return False
            '    End If
            '    sheet.SetRowBreak(Row_Cnt)
            '    BarCode_Cnt = 0
            '    If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '      Return False
            '    End If
            '  End If
            'Else
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            'End If
          End If


#End Region
        ElseIf ret_PlatForm = "OK Mart" Then
          If (lst_dicStore_Item.Count - Item_Cnt) < 2 Then
            flg_lastOne = True
          End If

#Region "OK Mart"
#Region "BarCode1"
          '=============================BarCode1====================================


          CreateBarCode(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper, 0, workbook, sheet, Row_Cnt)

          Row.CreateCell(0).SetCellValue("")
          Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
          Dim colorRgb = New Byte() {255, 255, 255}
          Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Barcode_Style.FillPattern = FillPattern.SolidForeground


          Dim Font As IFont = workbook.CreateFont()
          Font.FontName = "Code 128"
          Font.FontHeightInPoints = 36
          Barcode_Style.SetFont(Font)
          Row.Cells(0).CellStyle = Barcode_Style
          If flg_lastOne = False Then
            '##############################################################
            CreateBarCode(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper, 5, workbook, sheet, Row_Cnt)
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper)

            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Barcode_Style.FillPattern = FillPattern.SolidForeground

            Font.FontName = "Code 128"
            Font.FontHeightInPoints = 36
            Barcode_Style.SetFont(Font)
            Row.Cells(1).CellStyle = Barcode_Style
            '##############################################################
          End If

          '=============================================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper)

          Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Num_Style.FillPattern = FillPattern.SolidForeground


          Dim Num_Font As IFont = workbook.CreateFont()
          Num_Font.FontName = "Calibri"
          Num_Font.FontHeightInPoints = 12
          Num_Style.SetFont(Num_Font)
          Row.Cells(0).CellStyle = Num_Style

          If flg_lastOne = False Then
            '##############################################################
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper)

            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground

            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(1).CellStyle = Num_Style
            '##############################################################
          End If

          '=============================================================
          If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            Return False
          End If
          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
#End Region
#Region "BarCode2"
          '=============================BarCode2====================================
          CreateBarCode(lst_dicStore_Item.Item(Item_Cnt).BarCode2.ToUpper, 0, workbook, sheet, Row_Cnt)

          Row.CreateCell(0).SetCellValue("")
          'Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Barcode_Style.FillPattern = FillPattern.SolidForeground


          'Dim Font As IFont = workbook.CreateFont()
          Font.FontName = "Code 128"
          Font.FontHeightInPoints = 36
          Barcode_Style.SetFont(Font)
          Row.Cells(0).CellStyle = Barcode_Style
          If flg_lastOne = False Then
            '##############################################################
            CreateBarCode(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode2.ToUpper, 5, workbook, sheet, Row_Cnt)
            Row.CreateCell(5).SetCellValue("")

            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Barcode_Style.FillPattern = FillPattern.SolidForeground

            Font.FontName = "Code 128"
            Font.FontHeightInPoints = 36
            Barcode_Style.SetFont(Font)
            Row.Cells(1).CellStyle = Barcode_Style
            '##############################################################
          End If

          '=============================================================
          Row_Cnt = Row_Cnt + 1
          Row = sheet.CreateRow(Row_Cnt)
          Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode2.ToUpper)

          'Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
          'Dim colorRgb = New Byte() {255, 255, 255}
          Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
          Num_Style.FillPattern = FillPattern.SolidForeground


          'Dim Num_Font As IFont = workbook.CreateFont()
          Num_Font.FontName = "Calibri"
          Num_Font.FontHeightInPoints = 12
          Num_Style.SetFont(Num_Font)
          Row.Cells(0).CellStyle = Num_Style
          BarCode_Cnt = BarCode_Cnt + 1
          If flg_lastOne = False Then
            '##############################################################
            Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode2.ToUpper)

            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground

            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(1).CellStyle = Num_Style
            BarCode_Cnt = BarCode_Cnt + 1
            '##############################################################
          End If

          '=============================================================
          If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            Return False
          End If

          'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
          '  Return False
          'End If
          '切換Sheet，釋放記憶體
          'If Item_Cnt Mod 50 = 0 And Item_Cnt <> 0 Then
          If BarCode_Cnt >= 10 Then
            Dim sheet_Str = "Sheet" & page_cnt.ToString
            Header = sheet.Header
            Header.Left = "平台：" & ret_PlatForm & "  賣場(Lot)：" & ret_LotNo
            page_cnt = page_cnt + 1
            Row_Cnt = 0
            BarCode_Cnt = 0
            sheet = workbook.CreateSheet(sheet_Str)
            'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "3") = False Then
            '  Return False
            'End If
          Else
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            'If BarCode_Cnt >= 10 Then
            '  Total_BarCode = Total_BarCode + BarCode_Cnt
            '  If lst_dicStore_Item.Count <> Total_BarCode Then  '防止剛好滿頁時，不會再多換一頁
            '    sheet.SetRowBreak(Row_Cnt)
            '    BarCode_Cnt = 0
            '    If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '      Return False
            '    End If
            '  End If
            '  'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '  '  Return False
            '  'End If

            'Else
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            'End If
          End If






#End Region
#End Region
        ElseIf ret_PlatForm = "Family" Or ret_PlatForm = "7-11QRCode" Then
          If (lst_dicStore_Item.Count - Item_Cnt) < 2 Then
            flg_lastOne = True
          End If
#Region "Family"
#Region "BarCode4"
          'Dim _VALUE = GenCode128.Code128Rendering.MakeBarcodeImage(objStore_item.BarCode4, 2, True)
          Dim _VALUE = CodeEncoderFromString(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper, "QRCODE")
          Dim BYTE_Array As Byte() = BmpToBytes(_VALUE)
          Dim picInd = workbook.AddPicture(BYTE_Array, NPOI.SS.UserModel.PictureType.JPEG)

          Dim anchor As XSSFClientAnchor = New XSSFClientAnchor()
          anchor.Dx1 = 1
          anchor.Dy1 = 1


          anchor.Col1 = 0
          anchor.Row1 = Row_Cnt
          Dim drawing As XSSFDrawing = sheet.CreateDrawingPatriarch
          Dim pict As XSSFPicture = drawing.CreatePicture(anchor, picInd)
          pict.Resize()
          BarCode_Cnt = BarCode_Cnt + 1
          '=============================================================

          If flg_lastOne = False Then
            '##############################################################
            _VALUE = CodeEncoderFromString(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper, "QRCODE")
            BYTE_Array = BmpToBytes(_VALUE)
            picInd = workbook.AddPicture(BYTE_Array, NPOI.SS.UserModel.PictureType.JPEG)

            anchor = New XSSFClientAnchor()
            anchor.Dx1 = 1
            anchor.Dy1 = 1


            anchor.Col1 = 5
            anchor.Row1 = Row_Cnt
            drawing = sheet.CreateDrawingPatriarch
            pict = drawing.CreatePicture(anchor, picInd)
            pict.Resize()
            BarCode_Cnt = BarCode_Cnt + 1
            '##############################################################
          End If
          If ret_PlatForm = "7-11QRCode" Then
            '=============================================================
            Dim Barcode_Style As XSSFCellStyle = workbook.CreateCellStyle()
            Dim colorRgb = New Byte() {255, 255, 255}
            Barcode_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Row_Cnt = Row_Cnt + 4
            Row = sheet.CreateRow(Row_Cnt)
            Row.CreateCell(0).SetCellValue(lst_dicStore_Item.Item(Item_Cnt).BarCode1.ToUpper)

            Dim Num_Style As XSSFCellStyle = workbook.CreateCellStyle()
            'Dim colorRgb = New Byte() {255, 255, 255}
            Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
            Num_Style.FillPattern = FillPattern.SolidForeground


            Dim Num_Font As IFont = workbook.CreateFont()
            Num_Font.FontName = "Calibri"
            Num_Font.FontHeightInPoints = 12
            Num_Style.SetFont(Num_Font)
            Row.Cells(0).CellStyle = Num_Style

            If flg_lastOne = False Then
              '##############################################################
              Row.CreateCell(5).SetCellValue(lst_dicStore_Item.Item(Item_Cnt + 1).BarCode1.ToUpper)

              Num_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
              Num_Style.FillPattern = FillPattern.SolidForeground

              Num_Font.FontName = "Calibri"
              Num_Font.FontHeightInPoints = 12
              Num_Style.SetFont(Num_Font)
              Row.Cells(1).CellStyle = Num_Style
              '##############################################################
            End If
          Else
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
          End If




          'If Item_Cnt Mod 58 = 0 And Item_Cnt <> 0 Then
          If BarCode_Cnt >= 12 Then
            Dim sheet_Str = "Sheet" & page_cnt.ToString
            Header = sheet.Header
            Header.Left = "平台：" & ret_PlatForm & "  賣場(Lot)：" & ret_LotNo
            page_cnt = page_cnt + 1
            Row_Cnt = 0
            BarCode_Cnt = 0
            sheet = workbook.CreateSheet(sheet_Str)
            'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "3") = False Then
            '  Return False
            'End If
            'If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "2") = False Then
            '  Return False
            'End If

            'Dim sheet_Str = "Sheet" & page_cnt.ToString
            'Header = sheet.Header
            ''Header.Left = "平台：" & ret_PlatForm & "  賣場(Lot)：" & ret_LotNo
            'page_cnt = page_cnt + 1
            'Row_Cnt = 0
            'BarCode_Cnt = 0
            'sheet = workbook.CreateSheet(sheet_Str)
            ''If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "3") = False Then
            ''  Return False
            ''End If
          Else
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
              Return False
            End If
            'If BarCode_Cnt >= 12 Then
            '  Total_BarCode = Total_BarCode + BarCode_Cnt
            '  If lst_dicStore_Item.Count <> Total_BarCode Then  '防止剛好滿頁時，不會再多換一頁
            '    If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "2") = False Then
            '      Return False
            '    End If
            '    sheet.SetRowBreak(Row_Cnt)
            '    BarCode_Cnt = 0
            '    If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg, "2") = False Then
            '      Return False
            '    End If
            '  End If

            'Else
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            '  If Insert_Empty(workbook, sheet, Row, Row_Cnt, ret_strResultMsg) = False Then
            '    Return False
            '  End If
            'End If
          End If


#End Region
#End Region
        End If
        'sheet.SetRowBreak(32)
        Dim Now_Time_new As String = GetNewTime_DBFormat()

        SendMessageToLog("第" & Item_Cnt & "筆時間：" & Now_Time_new, eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
      Next








      workbook.Write(fs)
      fs.Close()


      If ret_strResultMsg.Length > 0 Then
        Return False
      Else
        Return True
      End If
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '取得要新增的SQL語句
  Private Function Get_SQL(ByRef ret_strResultMsg As String,
                           ByRef ret_dicAddStore_Head As Dictionary(Of String, clsSTORE_HEAD),
                           ByRef ret_dicAddStore_Item As Dictionary(Of String, clsSTORE_ITEM),
                           ByRef lstSql As List(Of String)) As Boolean
    Try
      '取得新增SQL
      For Each obj As clsSTORE_HEAD In ret_dicAddStore_Head.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Store_Head SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next
      For Each obj As clsSTORE_ITEM In ret_dicAddStore_Item.Values
        If obj.O_Add_Insert_SQLString(lstSql) = False Then
          ret_strResultMsg = "Get Insert Store_Item SQL Failed"
          SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '執行SQL語句，並進行記憶體資料更新
  Private Function Execute_DataUpdate(ByRef ret_strResultMsg As String,
                                      ByRef lstSql As List(Of String)) As Boolean
    Try
      '更新所有的SQL
      'If lstSql.Any = False Then '检查是否有要更新的SQL 如果没有检查是否有要给别人的命令
      '  '如果没有要给别人的命令 则回失败 (Message没做任何事!!)
      '  ret_strResultMsg = "Update SQL count is 0 and Send 0 Message to other system. Message do nothing!! Please Check!! ; 此笔命令无更新资料库，亦无发送其他命令给其它系统，请确认命令是否有问题。"
      '  SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      '  Return False
      'End If
      If Common_DBManagement.BatchUpdate(lstSql) = False Then
        '更新DB失敗則回傳False
        'ret_strResultMsg = "WMS Update DB Failed"
        ret_strResultMsg = "WMS 更新资料库失败"
        SendMessageToLog(ret_strResultMsg, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      End If

      '修改記憶體資料
      '1.新增
      'For Each objNew As clsALARM In ret_dicAddAlarm.Values
      '  objNew.Add_Relationship(gMain.objHandling)
      'Next
      ''2.删除
      'For Each objALARM As clsALARM In ret_dicDeleteAlarm.Values
      '  objALARM.Remove_Relationship()
      'Next
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Enum enuArrival_Status
    BCR = 1
    Palletizing = 2
  End Enum
  Public Function Insert_Empty(ByRef workbook As XSSFWorkbook, ByRef sheet As XSSFSheet,
                               ByRef ROW As XSSFRow, ByRef Row_Cnt As String,
                               ByRef ret_strResultMsg As String,
                               Optional ByVal Debug_Str As String = "1") As Boolean
    Try
      Dim colorRgb = New Byte() {255, 255, 255}

      Row_Cnt = Row_Cnt + 1
      ROW = sheet.CreateRow(Row_Cnt)
      ROW.CreateCell(0).SetCellValue(Debug_Str)

      Dim Empty_Style As XSSFCellStyle = workbook.CreateCellStyle()
      'Dim colorRgb = New Byte() {255, 255, 255}
      Empty_Style.SetFillForegroundColor(New XSSFColor(colorRgb))
      Empty_Style.FillPattern = FillPattern.SolidForeground
      'Barcode_Style.BorderTop = BorderStyle.Thin
      'Barcode_Style.BorderBottom = BorderStyle.Thin
      'Barcode_Style.BorderLeft = BorderStyle.Thin
      'Barcode_Style.BorderRight = BorderStyle.Thin

      Dim Empty_Font As IFont = workbook.CreateFont()
      Empty_Font.FontName = "Calibri"
      Empty_Font.FontHeightInPoints = 10
      Empty_Style.SetFont(Empty_Font)
      ROW.Cells(0).CellStyle = Empty_Style
      Return True
    Catch ex As Exception
      ret_strResultMsg = ex.ToString
      SendMessageToLog(ex.ToString, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  Private Sub CreateBarCode(ByVal data As String, ByVal Col As String, ByRef workbook As XSSFWorkbook, ByRef sheet As XSSFSheet, ByRef Row_Cnt As Integer, Optional ByVal fisrtBarCode As Boolean = False)
    Dim _VALUE1 As Bitmap = Nothing
    If fisrtBarCode = True Then
      _VALUE1 = New Bitmap(GenCode128.Code128Rendering.MakeBarcodeImage(data, 1, False))
    Else
      _VALUE1 = New Bitmap(GenCode128.Code128Rendering.MakeBarcodeImage(data, 1, False), 200, 20)
    End If

    'Dim _VALUE = CodeEncoderFromString(lst_dicStore_Item.Item(Item_Cnt).BarCode4, "QRCODE")
    Dim BYTE_Array1 As Byte() = BmpToBytes(_VALUE1)
    Dim picInd1 = workbook.AddPicture(BYTE_Array1, NPOI.SS.UserModel.PictureType.JPEG)

    Dim anchor1 As XSSFClientAnchor = New XSSFClientAnchor()
    anchor1.Dx1 = 1
    anchor1.Dy1 = 1


    anchor1.Col1 = Col
    anchor1.Row1 = Row_Cnt
    Dim drawing1 As XSSFDrawing = sheet.CreateDrawingPatriarch
    Dim pict1 As XSSFPicture = drawing1.CreatePicture(anchor1, picInd1)
    pict1.Resize()
  End Sub
End Module
