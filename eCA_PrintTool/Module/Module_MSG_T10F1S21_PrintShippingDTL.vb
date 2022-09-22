Module Module_MSG_T10F1S21_PrintShippingDTL

  Public Function O_ProcessT10F1S21_PrintShippingDTL(ByVal strMessage As String,
                                ByRef ret_strResult_Message As String,
                                Optional ByRef ret_strWait_UUID As String = "") As Boolean
    Try
      SendMessageToLog("Start Prcoess T10F1S21_PrintShippingDTL", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Dim PRINTER_NO = ""
      '解Message的Xml
      Dim Receive_Msg As MSG_T10F1S21_PrintShippingDTL = Nothing
      Receive_Msg = ParseXmlStringToClass(Of MSG_T10F1S21_PrintShippingDTL)(strMessage, ret_strResult_Message)
      If Receive_Msg Is Nothing Then
        SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        ret_strResult_Message = "ParseXmlString To Class Failed"
        Return False
      End If

      'PRINTER_NAME
      Dim ret_ResultMessage As String = Nothing
      'A4紙
      If PrintSetPrinter(ret_ResultMessage, PRINTER_NAME_A4, "") = False Then
        ret_ResultMessage = "找不到印表機"
        SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
        Return False
      Else
        ret_ResultMessage = "OK"
        SendMessageToLog(ret_ResultMessage, eCALogTool.ILogTool.enuTrcLevel.lvWARN)
      End If
      '-將資料轉成列印
      Dim PrintEnable = Receive_Msg.Body.PrintInfo.PRINT_ENABLE '是否列印 0:不列印/1:要列印
      Dim To_PDF = Receive_Msg.Body.PrintInfo.TO_PDF '是否轉PDF 0:不轉/1:要轉
      Dim PrintType = Receive_Msg.Body.PrintInfo.PRINT_TYPE '列印的類型 1:紙箱/2:物流箱
      Dim User_ID = Receive_Msg.Header.ClientInfo.UserID
      Dim UUID = Receive_Msg.Header.UUID
      Dim PrintFormatSettingPath = PrintFormatFile

      '開始組織欄位
      'Dim EXPORT_ID As String = GetNewTime_DBFormat()
      'Dim SAMPLE_FILE_NAME As String = ""
      'If PrintType = "1" Then
      '	SAMPLE_FILE_NAME = EnuPrintName.BoxBarcode.ToString
      'ElseIf PrintType = "2" Then
      '	SAMPLE_FILE_NAME = EnuPrintName.Logistics_BoxBarcode.ToString
      'Else
      '	ret_strResult_Message = "Error SAMPLE_FILE_NAME"
      '	Return False
      'End If
      'Dim CREATE_TIME As String = GetNewTime_DBFormat()
      'Dim FINISH_TIME As String = ""
      'Dim EXPORT_TYPE As Double = PrintEnable


      'Dim ret_dicAdd_EXPORT As New Dictionary(Of String, eCA_PrintTool.clsEXPORT)
      'Dim ret_dicAdd_EXPORT_DTL As New Dictionary(Of String, eCA_PrintTool.clsEXPORT_DTL)


      ''建立列印資訊
      'Dim New_objExport As New eCA_PrintTool.clsEXPORT(EXPORT_ID, SAMPLE_FILE_NAME, CREATE_TIME, FINISH_TIME, EXPORT_TYPE, PRINTER_NO)
      'If ret_dicAdd_EXPORT.ContainsKey(New_objExport.gid) = False Then
      '	ret_dicAdd_EXPORT.Add(New_objExport.gid, New_objExport)
      'End If



      ''組列印DTL
      'Dim New_objExport_DTL As eCA_PrintTool.clsEXPORT_DTL = Nothing
      'Dim TABLE_INDEX_1 As Double = 1 '第幾層 只目前支援三層(從1開始)
      'Dim TABLE_INDEX_2 As Double = 2 '第幾層 只目前支援三層(從1開始)
      'Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用		
      'Dim VALUE_INDEX As Double = 0 '該表中的第幾個位置(從1開始)
      'Dim VALUE As String = ""
      For Each objOrderInfo In Receive_Msg.Body.OrderList.OrderInfo
        'For Each objItemLabel In Receive_Msg.Body.LabelList.LabelInfo
        Dim ret_dicAdd_EXPORT As New Dictionary(Of String, eCA_PrintTool.clsEXPORT)
        Dim ret_dicAdd_EXPORT_DTL As New Dictionary(Of String, eCA_PrintTool.clsEXPORT_DTL)
        Dim EXPORT_ID As String = GetNewTime_DBFullTimeUUIDFormat()
        Dim SAMPLE_FILE_NAME As String = EnuPrintName.ShippingDetails.ToString
        'If PrintType = "A" Then
        '  SAMPLE_FILE_NAME = EnuPrintName.ItemLabel_A.ToString
        'ElseIf PrintType = "S" Then
        '  SAMPLE_FILE_NAME = EnuPrintName.ItemLabel_S.ToString
        'ElseIf PrintType = "G" Then
        '  SAMPLE_FILE_NAME = EnuPrintName.ItemLabel_G.ToString
        'Else
        '  ret_strResult_Message = "Error SAMPLE_FILE_NAME"
        '  Return False
        'End If
        Dim CREATE_TIME As String = GetNewTime_DBFormat()
        Dim FINISH_TIME As String = ""
        Dim EXPORT_TYPE As Double = PrintEnable
        '建立列印資訊
        Dim New_objExport As New eCA_PrintTool.clsEXPORT(EXPORT_ID, SAMPLE_FILE_NAME, CREATE_TIME, FINISH_TIME, EXPORT_TYPE, PRINTER_NO)
        If ret_dicAdd_EXPORT.ContainsKey(New_objExport.gid) = False Then
          ret_dicAdd_EXPORT.Add(New_objExport.gid, New_objExport)
        End If


        '組列印DTL
        Dim New_objExport_DTL As eCA_PrintTool.clsEXPORT_DTL = Nothing
        Dim TABLE_INDEX_1 As Double = 1 '第幾層 只目前支援三層(從1開始)
        Dim TABLE_INDEX_2 As Double = 2 '第幾層 只目前支援三層(從1開始)
        Dim SHEET_INDEX_1 As Double = 0 '第幾層(從1開始) '單層列印使用		
        Dim VALUE_INDEX As Double = 0 '該表中的第幾個位置(從1開始)
        Dim VALUE As String = ""

        '面料
        If SAMPLE_FILE_NAME = EnuPrintName.ShippingDetails.ToString Then
          For i = 1 To 9
            VALUE = "" '初始化
            VALUE_INDEX += 1
            Select Case i
              Case 1 'BarCode
                'VALUE = objOrderInfo.PO_ID
                VALUE = Code128b(objOrderInfo.PO_ID)
              Case 2 '收件人
                VALUE = objOrderInfo.TAG2
              Case 3 '連絡電話
                VALUE = objOrderInfo.TAG3
              Case 4 '地區
                VALUE = objOrderInfo.TAG5
              Case 5 '代收金額
                VALUE = objOrderInfo.TAG4
              Case 6 '客戶需求
                VALUE = objOrderInfo.TAG6
              Case 7 '調播單號
                VALUE = objOrderInfo.PO_ID
              Case 8 '接單日期
                VALUE = Now.ToString("yyyy/MM/dd")
              Case 9 '製單日期
                VALUE = objOrderInfo.TAG7
            End Select
            New_objExport_DTL = New eCA_PrintTool.clsEXPORT_DTL(EXPORT_ID, TABLE_INDEX_1, SHEET_INDEX_1, VALUE_INDEX, VALUE)
            If ret_dicAdd_EXPORT_DTL.ContainsKey(New_objExport_DTL.gid) = False Then
              ret_dicAdd_EXPORT_DTL.Add(New_objExport_DTL.gid, New_objExport_DTL)
            End If
          Next
          VALUE_INDEX = 0
          'SHEET_INDEX_1 = 1
          For Each objOrderDTLInfo In objOrderInfo.OrderDTLList.OrderDTLInfo
            For i = 1 To 6
              VALUE = "" '初始化
              VALUE_INDEX += 1
              Select Case i
                Case 1 '出貨類型
                  VALUE = objOrderDTLInfo.D_TAG5
                Case 2 '明細項次
                  VALUE = objOrderDTLInfo.PO_SERIAL_NO
                Case 3 '品號
                  VALUE = objOrderDTLInfo.D_TAG1
                Case 4 '品名
                  VALUE = objOrderDTLInfo.D_TAG2 & vbCrLf.ToString & objOrderDTLInfo.D_TAG3
                Case 5 '數量
                  VALUE = objOrderDTLInfo.QTY
                Case 6 '單位
                  VALUE = objOrderDTLInfo.D_TAG4
              End Select
              New_objExport_DTL = New eCA_PrintTool.clsEXPORT_DTL(EXPORT_ID, TABLE_INDEX_2, SHEET_INDEX_1, VALUE_INDEX, VALUE)
              If ret_dicAdd_EXPORT_DTL.ContainsKey(New_objExport_DTL.gid) = False Then
                ret_dicAdd_EXPORT_DTL.Add(New_objExport_DTL.gid, New_objExport_DTL)
              End If
            Next
          Next

        End If

        '-處理方式
        Dim ret_reportfilename = ""
        If To_PDF = "1" Then '-要轉PDF
          If PrintEnable = "1" Then '1:要列印
            ret_strResult_Message = "Not Print Type Function"
            SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
            Return False

          Else '-不要列印-會只存成PDF文件
            If eCA_PrintTool.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, "1", ret_strResult_Message) = False Then
              ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
              Return False
            End If
            '如果處理正確 則回寫完整擋案路徑在ResultMsg
            If ret_strResult_Message = "" Then
              ret_strResult_Message = ret_reportfilename
            End If
            'Return True
          End If


        Else '-不要轉PDF
          If PrintEnable = "1" Then '1:要列印
            If eCA_PrintTool.PrintStart(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, ret_strResult_Message, False) = False Then
              ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
              Return False
            Else
              'Return True
            End If

          Else '-不要列印
            ret_strResult_Message = "不列印不轉PDF 是來亂的嗎?"
            SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
            'Return True
          End If
        End If
        Threading.Thread.Sleep(1000)
      Next


      ''-處理方式
      'Dim ret_reportfilename = ""
      'If To_PDF = "1" Then '-要轉PDF
      '	If PrintEnable = "1" Then '1:要列印
      '		ret_strResult_Message = "Not Print Type Function"
      '		SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '		Return False

      '	Else '-不要列印-會只存成PDF文件
      '		If eCA_PrintTool.ReportPDFFile(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, ret_strResult_Message) = False Then
      '			ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
      '			Return False
      '		End If
      '		'如果處理正確 則回寫完整擋案路徑在ResultMsg
      '		If ret_strResult_Message = "" Then
      '			ret_strResult_Message = ret_reportfilename
      '		End If
      '		Return True
      '	End If


      'Else '-不要轉PDF
      '	If PrintEnable = "1" Then '1:要列印
      '		If eCA_PrintTool.PrintStart(ret_dicAdd_EXPORT, ret_dicAdd_EXPORT_DTL, SamplePath, ExportPath, PrintFormatSettingPath, ret_reportfilename, User_ID, UUID, ret_strResult_Message, False) = False Then
      '			ret_strResult_Message = ModifyStringApostrophe(ret_strResult_Message)
      '			Return False
      '		Else
      '			Return True
      '		End If

      '	Else '-不要列印
      '		ret_strResult_Message = "不列印不轉PDF 是來亂的嗎?"
      '		SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      '		Return True
      '	End If
      'End If


      Return True
    Catch ex As Exception
      ret_strResult_Message = ex.Message
      SendMessageToLog(ret_strResult_Message, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

End Module
