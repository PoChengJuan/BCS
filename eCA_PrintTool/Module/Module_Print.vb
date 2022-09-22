Imports NPOI.HSSF.UserModel
Imports System.IO
Imports NDde
Imports NDde.Client
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports System.Management
Imports System.Collections
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.Mvc
Imports NPOItest.Models.Sevices
Imports NPOI.XWPF.UserModel



Public Module Module_Print
  Private ClsConfig As clsConfigTool = Nothing
  Public Function PrintStart(ByRef tmp_dicExport As Dictionary(Of String, clsEXPORT),
                                ByRef tmp_dicExport_DTL As Dictionary(Of String, clsEXPORT_DTL),
                                ByVal SamplePath As String,
                                ByVal ExportPath As String,
                                ByVal PrintFormatSettingPath As String,
                                ByRef ret_ReportFileName As String,
                                ByRef USER_ID As String,
                                ByRef UUID As String,
                                ByRef ret_ResultMsg As String,
                             ByVal convertToPDF As Boolean,
                                Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0, Optional onPrintFlag As Boolean = False) As Boolean
    Try
      ClsConfig = New clsConfigTool(PrintFormatSettingPath)
      If onPrintFlag Then '-路徑列印
        If OnlyStartPrint(SamplePath, USER_ID, UUID, convertToPDF, ret_ResultMsg) Then
          Return True
        Else
          Return False
        End If

      End If


      '-防止傳入的dicExport和dicExport_DTL是否是Nothing或是空的
      If (tmp_dicExport Is Nothing OrElse tmp_dicExport.Count = 0) Or (tmp_dicExport_DTL Is Nothing OrElse tmp_dicExport_DTL.Count = 0) Then
        ret_ResultMsg = "tmp_dicExport Or tmp_dicExport_DTL Is Empty"
        Return False
      End If
      Dim EXPORT_TYPE = ""
      Dim SAMPLE_FILE_NAME = ""
      Dim _QUEUE As New Queue(Of List(Of List(Of String)))
      Dim _list As New List(Of String)
      For Each obj In tmp_dicExport_DTL.Values
        _list.Add(obj.VALUE)
      Next
      For Each objExport In tmp_dicExport.Values
        ' If dicDelete_Export.ContainsKey(objExport.gid) Then Continue For
        Dim _printData = objExport
        'get Table

        If tmp_dicExport_DTL.Count <> 0 Then
          Dim _GTable = From item In tmp_dicExport_DTL.Values Group item By item.TABLE_INDEX Into Group
          For Each _item In _GTable
            Dim _addAlllist As New List(Of List(Of String))
            Dim _GLine = (From item In _item.Group Group item By item.COLUMN_INDEX Into Group).OrderBy(Function(p) Convert.ToInt32(p.COLUMN_INDEX))   '-根據行分組
            Dim _dic = _GLine.ToDictionary(Function(p) p.COLUMN_INDEX, Function(k) k.Group)
            For Each it In _dic
              Dim _Data = it.Value.OrderBy(Function(p) Convert.ToInt32(p.ROW_INDEX))
              Dim _addlist As New List(Of String)
              _addAlllist.Add(_addlist)
              For i = 0 To _Data.Count - 1
                _addlist.Add(_Data(i).VALUE)
              Next
            Next
            _QUEUE.Enqueue(_addAlllist)
          Next
          EXPORT_TYPE = _printData.EXPORT_TYPE
          SAMPLE_FILE_NAME = _printData.SAMPLE_FILE_NAME
        End If
      Next
      If EXPORT_TYPE = "0" Then '-不啟動列印 單純寫進料
        If StartWrite(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, "1", convertToPDF, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
      ElseIf EXPORT_TYPE = "1" Then '-啟動列印
        If StartPrint(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, convertToPDF, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
      End If
      Return True
    Catch ex As Exception

    End Try
  End Function
  ''' <summary>
  ''' 傳入要列印的資料，回傳PDFFile並可進行列印
  ''' </summary>
  ''' <param name="tmp_dicExport"></param>
  ''' <param name="tmp_dicExport_DTL"></param>
  ''' <param name="SamplePath"></param>
  ''' <param name="ExportPath"></param>
  ''' <param name="PrintFormatSettingPath"></param>
  ''' <param name="ret_ReportFileName"></param>
  ''' <param name="USER_ID"></param>
  ''' <param name="UUID"></param>
  ''' <param name="ret_ResultMsg"></param>
  ''' <returns></returns>	
  Public Function ReportPDFFile(ByRef tmp_dicExport As Dictionary(Of String, clsEXPORT),
                                ByRef tmp_dicExport_DTL As Dictionary(Of String, clsEXPORT_DTL),
                                ByVal SamplePath As String,
                                ByVal ExportPath As String,
                                ByVal PrintFormatSettingPath As String,
                                ByRef ret_ReportFileName As String,
                                ByRef USER_ID As String,
                                ByRef UUID As String,
                                ByVal Serial_No As String,
                                ByRef ret_ResultMsg As String,
                                Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      ClsConfig = New clsConfigTool(PrintFormatSettingPath)
      '-防止傳入的dicExport和dicExport_DTL是否是Nothing或是空的
      If (tmp_dicExport Is Nothing OrElse tmp_dicExport.Count = 0) Or (tmp_dicExport_DTL Is Nothing OrElse tmp_dicExport_DTL.Count = 0) Then
        ret_ResultMsg = "tmp_dicExport Or tmp_dicExport_DTL Is Empty"
        Return False
      End If
      Dim EXPORT_TYPE = ""
      Dim SAMPLE_FILE_NAME = ""
      Dim _QUEUE As New Queue(Of List(Of List(Of String)))
      Dim _list As New List(Of String)
      For Each objExport In tmp_dicExport.Values
        ' If dicDelete_Export.ContainsKey(objExport.gid) Then Continue For
        Dim _printData = objExport
        'get Table
        For Each obj In tmp_dicExport_DTL.Values
          _list.Add(obj.VALUE)
        Next
        If tmp_dicExport_DTL.Count <> 0 Then
          Dim _GTable = From item In tmp_dicExport_DTL.Values Group item By item.TABLE_INDEX Into Group
          For Each _item In _GTable
            Dim _addAlllist As New List(Of List(Of String))
            Dim _GLine = (From item In _item.Group Group item By item.COLUMN_INDEX Into Group).OrderBy(Function(p) Convert.ToInt32(p.COLUMN_INDEX))   '-根據行分組
            Dim _dic = _GLine.ToDictionary(Function(p) p.COLUMN_INDEX, Function(k) k.Group)
            For Each it In _dic
              Dim _Data = it.Value.OrderBy(Function(p) Convert.ToInt32(p.ROW_INDEX))
              Dim _addlist As New List(Of String)
              _addAlllist.Add(_addlist)
              For i = 0 To _Data.Count - 1
                _addlist.Add(_Data(i).VALUE)
              Next
            Next
            _QUEUE.Enqueue(_addAlllist)
          Next
          EXPORT_TYPE = _printData.EXPORT_TYPE
          SAMPLE_FILE_NAME = _printData.SAMPLE_FILE_NAME
          'excute
          'If _printData.EXPORT_TYPE = "0" Then '-不啟動列印 單純寫進料
          '  If StartWrite(_list, _QUEUE, _printData.SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, ckeckExist) = False Then
          '    Return False
          '  End If
          'ElseIf _printData.EXPORT_TYPE = "1" Then '-啟動列印
          '  If StartPrint(_list, _QUEUE, _printData.SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, ckeckExist) = False Then
          '    Return False
          '  End If
          'End If
        End If
      Next
      If EXPORT_TYPE = "0" Then '-不啟動列印 單純寫進料
        If StartWrite(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, Serial_No, True, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
      ElseIf EXPORT_TYPE = "1" Then '-啟動列印
        'If StartPrint(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
        '  Return False
        'End If
        '寫進PDF
        If StartWrite(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, Serial_No, True, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
        '列印PDF
        If PrintPDF(ret_ReportFileName, ret_ResultMsg) = False Then
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      ret_ResultMsg = ex.Message
      Return False
    End Try
  End Function
  Public Function ReportPDFFileForMassive(ByRef tmp_dicExport As Dictionary(Of String, clsEXPORT),
                                ByRef tmp_dicExport_DTL As Dictionary(Of String, clsEXPORT_DTL),
                                ByVal SamplePath As String,
                                ByVal ExportPath As String,
                                ByVal PrintFormatSettingPath As String,
                                ByRef ret_ReportFileName As String,
                                ByRef USER_ID As String,
                                ByRef UUID As String,
                                ByRef ret_ResultMsg As String,
                                Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      ClsConfig = New clsConfigTool(PrintFormatSettingPath)
      '-防止傳入的dicExport和dicExport_DTL是否是Nothing或是空的
      If (tmp_dicExport Is Nothing OrElse tmp_dicExport.Count = 0) Or (tmp_dicExport_DTL Is Nothing OrElse tmp_dicExport_DTL.Count = 0) Then
        ret_ResultMsg = "tmp_dicExport Or tmp_dicExport_DTL Is Empty"
        Return False
      End If
      Dim EXPORT_TYPE = ""
      Dim SAMPLE_FILE_NAME = ""
      Dim _QUEUE As New Queue(Of List(Of List(Of String)))
      Dim _list As New List(Of String)
      For Each objExport In tmp_dicExport.Values
        ' If dicDelete_Export.ContainsKey(objExport.gid) Then Continue For
        Dim _printData = objExport
        'get Table
        For Each obj In tmp_dicExport_DTL.Values
          _list.Add(obj.VALUE)
        Next
        If tmp_dicExport_DTL.Count <> 0 Then
          Dim _GTable = From item In tmp_dicExport_DTL.Values Group item By item.TABLE_INDEX Into Group
          For Each _item In _GTable
            Dim _addAlllist As New List(Of List(Of String))
            Dim _GLine = (From item In _item.Group Group item By item.COLUMN_INDEX Into Group).OrderBy(Function(p) Convert.ToInt32(p.COLUMN_INDEX))   '-根據行分組
            Dim _dic = _GLine.ToDictionary(Function(p) p.COLUMN_INDEX, Function(k) k.Group)
            For Each it In _dic
              Dim _Data = it.Value.OrderBy(Function(p) Convert.ToInt32(p.ROW_INDEX))
              Dim _addlist As New List(Of String)
              _addAlllist.Add(_addlist)
              For i = 0 To _Data.Count - 1
                _addlist.Add(_Data(i).VALUE)
              Next
            Next
            _QUEUE.Enqueue(_addAlllist)
          Next
          EXPORT_TYPE = _printData.EXPORT_TYPE
          SAMPLE_FILE_NAME = _printData.SAMPLE_FILE_NAME
          'excute
          'If _printData.EXPORT_TYPE = "0" Then '-不啟動列印 單純寫進料
          '  If StartWrite(_list, _QUEUE, _printData.SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, ckeckExist) = False Then
          '    Return False
          '  End If
          'ElseIf _printData.EXPORT_TYPE = "1" Then '-啟動列印
          '  If StartPrint(_list, _QUEUE, _printData.SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, ckeckExist) = False Then
          '    Return False
          '  End If
          'End If
        End If
      Next
      If EXPORT_TYPE = "0" Then '-不啟動列印 單純寫進料
        If StartWriteForMassive(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
      ElseIf EXPORT_TYPE = "1" Then '-啟動列印
        If StartPrint(_list, _QUEUE, SAMPLE_FILE_NAME, SamplePath, ExportPath, USER_ID, UUID, True, ret_ReportFileName, ret_ResultMsg, Width, Height) = False Then
          Return False
        End If
      End If
      Return True
    Catch ex As Exception
      ret_ResultMsg = ex.Message
      Return False
    End Try
  End Function



  ''' <summary>
  ''' 傳入既有的Excle檔案名稱，產生PDF檔，並回傳檔案名稱
  ''' </summary>
  ''' <param name="ExcelFullFileName"></param>
  ''' <param name="PDFFullFileName"></param>
  ''' <param name="ret_ResultMsg"></param>
  ''' <returns></returns>
  Public Function ExcelToPDFFile(ByVal ExcelFullFileName As String,
                                 ByRef PDFFullFileName As String,
                                 ByRef ret_ResultMsg As String) As Boolean
    Try
      If eCA_Excel.ExcelToPDF(ExcelFullFileName, PDFFullFileName, ret_ResultMsg) = False Then
        Return False
      End If
      Return True
    Catch ex As Exception
      ret_ResultMsg = ex.Message
      Return False
    End Try
  End Function
  Public Function ExcelToPDFFile2(ByVal ExcelFullFileName As String,
                                 ByRef PDFFullFileName As String,
                                 ByRef ret_ResultMsg As String) As Boolean
    Try
      If eCA_Excel.ExcelToPDF(ExcelFullFileName, PDFFullFileName, ret_ResultMsg) = False Then
        Return False
      End If
      Return True
      'If eCA_Excel.ExcelToPDF2(ExcelFullFileName, PDFFullFileName, ret_ResultMsg) = False Then
      '  Return False
      'End If
      'Return True
    Catch ex As Exception
      ret_ResultMsg = ex.Message
      Return False
    End Try
  End Function

  Private Function StartWriteForMassive(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))),
                                     ByVal _SampleFileName As String,
                                     ByVal SamplePath As String,
                                     ByVal ExportPath As String,
                                     ByRef USER_ID As String,
                                     ByRef UUID As String,
                                     ByVal convertToPDF As Boolean,
                                     ByRef ret_ReportFilePath As String,
                                     ByRef ret_ResultMsg As String,
                                     Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      Dim _FileExportPath = ExportPath & Now.ToString("yyMMdd")  '---print來源 
      '1.檢查folder,original ,sample 必要存在 
      If CheckFolder(SamplePath, _FileExportPath) Then

        '2.檢查SamplePath資料夾內 是否有要列印的格式
        For Each fname As String In System.IO.Directory.GetFileSystemEntries(SamplePath)
          If Path.GetFileNameWithoutExtension(fname) = _SampleFileName Then
            '3.file 複製到Export資料夾內
            Dim _destFile = _FileExportPath & "\" & _SampleFileName & "_" & USER_ID & "_" & UUID & ".xls"
            File.Copy(fname, _destFile)
            '---main        
            'SendMessageToLog("复制完毕，开始建立excel资料", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            If ReadExcel2(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) = False Then
              'If ReadExcel(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) = False Then
              Return False
            Else '-成功
              If convertToPDF Then '-轉pdf
                If Width <> 0 AndAlso Height <> 0 Then
                  If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, Width, Height, "") = False Then
                    ret_ResultMsg += " 轉pdf失败"
                  End If
                Else
                  'If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, "") = False Then
                  '  ret_ResultMsg += " 轉pdf失败"
                  'End If
                  If eCA_Excel.ExcelToPDF2(_destFile, ret_ReportFilePath, "") = False Then
                    ret_ResultMsg += " 轉pdf失败"
                  End If
                End If
              End If
            End If
            Exit For
          End If
        Next
        Return True
      Else
        ret_ResultMsg = "Sample位置异常，请确认config。"
        Return False
      End If
    Catch ex As Exception
      ret_ResultMsg = ex.Message
    End Try
  End Function
  Private Function StartWrite(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))),
                                     ByVal _SampleFileName As String,
                                     ByVal SamplePath As String,
                                     ByVal ExportPath As String,
                                     ByRef USER_ID As String,
                                     ByRef UUID As String,
                                     ByVal Serial_No As String,
                                     ByVal convertToPDF As Boolean,
                                     ByRef ret_ReportFilePath As String,
                                     ByRef ret_ResultMsg As String,
                                     Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      Dim _FileExportPath = ExportPath & Now.ToString("yyMMdd")  '---print來源 
      '1.檢查folder,original ,sample 必要存在 
      If CheckFolder(SamplePath, _FileExportPath) Then

        '2.檢查SamplePath資料夾內 是否有要列印的格式
        For Each fname As String In System.IO.Directory.GetFileSystemEntries(SamplePath)
          If Path.GetFileNameWithoutExtension(fname) = _SampleFileName Then
            '3.file 複製到Export資料夾內
            Dim _destFile = _FileExportPath & "\" & _SampleFileName & "_" & USER_ID & "_" & UUID & "-" & Serial_No & ".xls"
            File.Copy(fname, _destFile)
            '---main        
            'SendMessageToLog("复制完毕，开始建立excel资料", eCALogTool.ILogTool.enuTrcLevel.lvDEBUG)
            If _SampleFileName = EnuPrintName.PackageinfoLabel.ToString Then
              If ReadExcelForForm(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) = False Then
                Return False
              Else '-成功
                If convertToPDF Then '-轉pdf
                  If Width <> 0 AndAlso Height <> 0 Then
                    If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, Width, Height, "") = False Then
                      ret_ResultMsg += " 轉pdf失败"
                    End If
                  Else
                    'If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, "") = False Then
                    '  ret_ResultMsg += " 轉pdf失败"
                    'End If
                    If eCA_Excel.ExcelToPDF2(_destFile, ret_ReportFilePath, "") = False Then
                      ret_ResultMsg += " 轉pdf失败"
                    End If
                  End If
                End If
              End If
            Else
              If ReadExcel(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) = False Then
                Return False
              Else '-成功
                If convertToPDF Then '-轉pdf
                  If Width <> 0 AndAlso Height <> 0 Then
                    If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, Width, Height, "") = False Then
                      ret_ResultMsg += " 轉pdf失败"
                    End If
                  Else
                    'If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, "") = False Then
                    '  ret_ResultMsg += " 轉pdf失败"
                    'End If
                    If eCA_Excel.ExcelToPDF2(_destFile, ret_ReportFilePath, "") = False Then
                      ret_ResultMsg += " 轉pdf失败"
                    End If
                  End If
                End If
              End If
            End If
            Exit For
          End If
        Next
        Return True
      Else
        ret_ResultMsg = "Sample位置异常，请确认config。"
        Return False
      End If
    Catch ex As Exception
      ret_ResultMsg = ex.Message
    End Try
  End Function
  Private Function OnlyStartPrint(ByVal SamplePath As String,
                                    ByRef USER_ID As String,
                                    ByRef UUID As String,
                                    ByVal convertToPDF As Boolean,
                                    ByRef ret_ResultMsg As String,
                                    Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      If printExcel(SamplePath, ClsConfig.ReadStringValueKey("Public", "EXCEL執行檔")) = False Then
        ret_ResultMsg = "列印失败"
      Else


      End If
    Catch ex As Exception

    End Try
  End Function
  Private Function StartPrint(ByVal _list As List(Of String), ByVal _queue As Queue(Of List(Of List(Of String))),
                                    ByVal _SampleFileName As String,
                                    ByVal SamplePath As String,
                                    ByVal ExportPath As String,
                                    ByRef USER_ID As String,
                                    ByRef UUID As String,
                                    ByVal convertToPDF As Boolean,
                                    ByRef ret_ReportFilePath As String,
                                    ByRef ret_ResultMsg As String,
                                    Optional ByVal Width As Double = 0, Optional ByVal Height As Double = 0) As Boolean
    Try
      Dim _FileExportPath = ExportPath & Now.ToString("yyMMdd") & "\" & Now.ToString("yyyyMMddHHmmssfff")  '---print來源 			
      '1.檢查folder,original ,sample 必要存在 
      If CheckFolder(SamplePath, _FileExportPath) Then
        '2.檢查SamplePath資料夾內 是否有要列印的格式
        For Each fname As String In System.IO.Directory.GetFileSystemEntries(SamplePath)
          If Path.GetFileNameWithoutExtension(fname) = _SampleFileName Then
            '3.file 複製到Export資料夾內
            Dim _destFile = _FileExportPath & "\" & _SampleFileName & "_" & USER_ID & "_" & UUID & ".xls"
            File.Copy(fname, _destFile)
            '---main          
            If _SampleFileName = EnuPrintName.PackageinfoLabel.ToString Then
              If ReadExcelForForm(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) Then
                If printExcel(_destFile, ClsConfig.ReadStringValueKey("Public", "EXCEL執行檔")) = False Then
                  ret_ResultMsg = "列印失败"
                  'Return False '-列印失敗
                End If
                If convertToPDF Then '-轉PDF
                  If Width <> 0 AndAlso Height <> 0 Then
                    If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, Width, Height, ret_ResultMsg) = False Then
                      ret_ResultMsg += " 轉PDF失败"
                    End If
                  Else
                    If eCA_Excel.ExcelToPDF2(_destFile, ret_ReportFilePath, ret_ResultMsg) = False Then
                      ret_ResultMsg += " 轉PDF失败"
                    End If
                  End If
                End If
              Else
                Return False
              End If

            Else
              If ReadExcel(_destFile, _SampleFileName, _queue, _list, ret_ResultMsg) Then
                If printExcel(_destFile, ClsConfig.ReadStringValueKey("Public", "EXCEL執行檔")) = False Then
                  ret_ResultMsg = "列印失败"
                  'Return False '-列印失敗
                End If
                If convertToPDF Then '-轉PDF
                  If Width <> 0 AndAlso Height <> 0 Then
                    If eCA_Excel.ExcelToPDF(_destFile, ret_ReportFilePath, Width, Height, ret_ResultMsg) = False Then
                      ret_ResultMsg += " 轉PDF失败"
                    End If
                  Else
                    If eCA_Excel.ExcelToPDF2(_destFile, ret_ReportFilePath, ret_ResultMsg) = False Then
                      ret_ResultMsg += " 轉PDF失败"
                    End If
                  End If
                End If
              Else
                Return False
              End If
            End If
            Exit For
          End If
        Next
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      ret_ResultMsg = ex.Message
      Return False
    End Try

  End Function
  Private Function CheckFolder(ByVal Original_samplePath As String, ByVal _samplePath As String) As Boolean
    Try
      '1檢查folder 是否存在
      If Not Directory.Exists(Original_samplePath) Then Return False
      If Not Directory.Exists(_samplePath) Then Directory.CreateDirectory(_samplePath) ' Return False

      ''2.清空sample Folder 內的的檔案
      'For Each fname As String In System.IO.Directory.GetFileSystemEntries(_samplePath)
      '  File.Delete(fname)
      'Next

      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Private Function ReadExcel2(ByVal sExcelPath As String, ByVal strName As String, ByVal _queue As Queue(Of List(Of List(Of String))),
                              ByVal _list As List(Of String), ByRef msgResult_Message As String) As Boolean
    Try
      Dim wk As HSSFWorkbook
      Using fs As New FileStream(sExcelPath, FileMode.Open, FileAccess.ReadWrite)
        wk = New HSSFWorkbook(fs) '-抓整個excel
        Dim hst = DirectCast(wk.GetSheetAt(0), HSSFSheet) '---抓一個sheet 理論上應該只有一個Sheet
        Dim sorce_hst = hst.CopySheet("Sample") '---抓一個sheet 理論上應該只有一個Sheet				

        '-根據有多少個要列印的資訊拆成多個Sheet	
        Dim Allist = _queue(0)
        Dim _count = 1000
        Dim _sheetCount = (Allist.Count \ _count) '-要幾個Sheet
        If (Allist.Count Mod _count) <> 0 Then _sheetCount = _sheetCount + 1

        '取的Cofig	
        Dim Public_dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary("Public")
        Dim PrintMode = ClsConfig.ReadStringValueDictionary("Public", strName)
        Dim _strSection = Public_dic(strName)
        Dim _dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_strSection)


        '1. 複製sheet
        Dim lstSheet As New List(Of HSSFSheet)
        For i = 0 To _sheetCount - 1  '---Table 生成      
          Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
          lstSheet.Add(hst_CopySheet)
        Next

        For i = 0 To _sheetCount - 1
          Dim tmp As New List(Of List(Of String))

          For j = 0 To _count - 1
            Dim _value = j + (_count * i)
            If Allist.Count <= _value Then Continue For
            tmp.Add(Allist(j + (_count * i)))
          Next

          ReadExcelTest(tmp, lstSheet(i), sorce_hst, wk, _dic, _strSection, _list, msgResult_Message)
        Next

        '4. 移除多餘的sheet
        Dim clear = False
        While clear = False
          For i = 0 To wk.Count - 1
            If wk(i).SheetName.Contains("Sheet_") = False Then
              wk.RemoveSheetAt(i)
              Exit For
            End If
            If i = wk.Count - 1 Then
              clear = True
            End If
          Next
        End While
      End Using


      '3.寫回檔案
      Dim file As New FileStream(sExcelPath, FileMode.Create)
      '產生檔案
      wk.Write(file)
      file.Close()
      Return True
    Catch ex As Exception
      msgResult_Message = ex.Message
      Return False
    End Try
  End Function
  Private Function ReadExcelTest(ByVal Allist As List(Of List(Of String)), ByRef hst As HSSFSheet, ByRef sorce_hst As HSSFSheet, ByRef wk As HSSFWorkbook,
                               ByVal _dic As Dictionary(Of String, String), ByVal _strSection As String, ByVal _list As List(Of String),
                                 ByRef msgResult_Message As String) As Integer
    Try
      Dim _count = 0
      Dim _Tablecount = 1
      Dim ruleTablecount = 0
      For Each item In _dic
        Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)

        If _info("TABLE") = "0" Then
          Dim hr = DirectCast(hst.GetRow(_info("Y") - 1), HSSFRow)
          Dim _cellcount = hr.Cells.Count
          'hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
          If _info("TYPE") = "STRING" Then
            hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
          ElseIf _info("TYPE") = "QRCODE" Or _info("TYPE") = "DATAMATRIX" Then
            'Dim _VALUE = CodeEncoderFromString(_list(_count), _info("TYPE")) '-要塞進去QRCode的值
            'Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), PictureType.JPEG)
            'Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
            ''Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1, _info("X") + 1, _info("Y") + 4)
            'Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1, _info("X") + 1, _info("Y") + 5)
            'anchor.AnchorType = 0
            'Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
            ''Reset the image to the original size.(若沒設定這個，需要注意HSSFClientAnchor最後兩個參數)
            ''picture.Resize()
          Else
            If _list(_count) <> "" Then
              hr.Cells(_info("X") - 1).SetCellValue(Convert.ToDouble(_list(_count)))
            End If
          End If


          _count = _count + 1

          'For i = 0 To wk.Count - 1
          '	If wk(i).SheetName.Contains("Sample") Then
          '		wk.RemoveSheetAt(i)
          '		Exit For
          '	End If
          'Next
        End If

        'Dim tmp1 = IIf(ruleTablecount = 2, 1, 0)
        Dim tmp2 = IIf(ruleTablecount = 2, 1, 0)

        If _info("TABLE").Contains("Table") Then  '---TABLE表    '---TABLE表          
          'For i = 0 To (Allist.Count - 1) ' - 1 '---Table 生成(全部n列，就新增n-1列，因為原本的範例就有一列了)         
          '	If _info("Y") = 0 Then Continue For
          '	hst.CopyRow(_info("Y") - 2 + _Tablecount + tmp1, _info("Y") - 1 + _Tablecount + tmp1) ' + _Tablecount)

          'Next

          ''---取得塞Table的位置 new         
          'Dim _locat As New List(Of clsExcelSet)
          'Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
          'For Each _item In _dicseatitem
          '	Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
          '	Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))

          '	_locat.Add(_point)
          'Next

          ''---塞table new
          'For i = 0 To Allist.Count - 1
          '	'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
          '	For j = 0 To Allist(i).Count - 1
          '		Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount + tmp2), HSSFRow)
          '		If _locat(j).TYPE = "STRING" Then
          '			'benny test
          '			If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
          '				Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
          '				Dim _test = hst.SheetConditionalFormatting
          '				Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
          '				Dim _font = _rule.CreateFontFormatting
          '				'_font.SetFontStyle(False, True)
          '				'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

          '				Dim _x = _locat(j).X - 1
          '				Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
          '				Dim _result = ""
          '				If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
          '				Dim _value = _result & _Y & ":" & _result & _Y
          '				Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
          '				hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
          '			End If
          '			'benny test											
          '			hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))

          '		Else
          '			If Allist(i)(j) <> "" Then
          '				hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToDouble(Allist(i)(j)))
          '				'benny test
          '				If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
          '					Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
          '					Dim _test = hst.SheetConditionalFormatting
          '					Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
          '					Dim _font = _rule.CreateFontFormatting
          '					'_font.SetFontStyle(False, True)
          '					'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

          '					Dim _x = _locat(j).X - 1
          '					Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
          '					Dim _result = ""
          '					If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
          '					Dim _value = _result & _Y & ":" & _result & _Y
          '					Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
          '					hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
          '				End If
          '				'benny test
          '			End If
          '		End If
          '	Next
          'Next
          ''若上一個Table有n列，則下一個Table位移n-1列(因為新增加了n-1列)
          '_Tablecount = _Tablecount + (Allist.Count - 1)
          'ruleTablecount += 1
        ElseIf _info("TABLE").Contains("Form") Then '-特殊列印
          Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
          Dim Interval = _info("FormInterval")  '-複製間距
          '刪除原本的所有合併儲存格
          Dim MergedCount = hst.NumMergedRegions
          For i = 0 To MergedCount
            hst.RemoveMergedRegion(0)
          Next
          ''複製格式
          'For i = 0 To Allist.Count - 1 - 1 '---Table 生成      
          'For j = 0 To Interval - 1
          '	hst.CopyRow(_info("Y") + j - 1, _info("Y") + j + Interval - 1)


          '||要採用跑一次就把所有Row複製出來
          Dim _lstTest As New List(Of String) '-for test
          For j = 0 To Interval - 1

            'test
            'If j = 13 Then
            '	Dim _WWW = ""
            'End If
            '--
            'Dim _Test = hst.CopyRow(_info("Y") + j - 1, _info("Y") + j + Interval - 1)
            Dim _row = hst.GetRow(_info("Y") + j - 1)

            For k = 0 To Allist.Count - 1 - 1
              _row.CopyRowTo(_info("Y") + j + Interval - 1 + (Interval * k))
              'TEST
              'If (_info("Y") + j + Interval - 1 + (Interval * k)) = 31 Then
              '	Dim _WWW = ""
              'End If
              ''TEST
              '_lstTest.Add(_info("Y") + j - 1 & "_" & _info("Y") + j + Interval - 1 + (Interval * k)) '-for test
            Next
          Next

          '-分頁線
          For k = 0 To Allist.Count - 1 - 1
            _RowBreak += 1
          Next

          'Next
          '_RowBreak += 1
          '設定格式化的條件
          Dim ConditionalCount = sorce_hst.SheetConditionalFormatting.NumConditionalFormattings
          For L = 0 To ConditionalCount - 1
            Dim cf = sorce_hst.SheetConditionalFormatting.GetConditionalFormattingAt(L)
            hst.SheetConditionalFormatting.AddConditionalFormatting(cf)
          Next

          'Next

          '設定長寬
          For i = 0 To _RowBreak - 1
            For j = 0 To Interval - 1
              hst.GetRow(_info("Y") + j + Interval - 1 + (i * Interval)).Height = hst.GetRow(_info("Y") + j - 1).Height
            Next
          Next

          '設定合併儲存格
          For i = 0 To Allist.Count - 1
            '取得第一組範例的合併儲存格資訊
            MergedCount = sorce_hst.NumMergedRegions
            For k = 0 To MergedCount - 1
              Dim MergedRegionAt = sorce_hst.GetMergedRegion(k)
              Dim toFirstRow = MergedRegionAt.FirstRow + Interval * i
              Dim toLastRow = MergedRegionAt.LastRow + Interval * i
              Dim toFirstCol = MergedRegionAt.FirstColumn
              Dim toLastCol = MergedRegionAt.LastColumn
              Dim toMergedRegionAt As New CellRangeAddress(toFirstRow, toLastRow, toFirstCol, toLastCol)
              hst.AddMergedRegion(toMergedRegionAt)
            Next

            '儲存格格式設定
            For j = 0 To Interval - 1
              'If j = 7 Then
              '	Dim _WWW = ""
              'End If
              If (Allist.Count - 1 - 1) > 1 Then 'benny 20200130
                Dim newRow As HSSFRow = hst.GetRow(_info("Y") + j + Interval - 1)
                Dim sourceRow As HSSFRow = sorce_hst.GetRow(_info("Y") + j - 1)
                For k = 0 To sourceRow.LastCellNum - 1
                  Dim oldCell As HSSFCell = sourceRow.GetCell(k)
                  If oldCell.ToString = "" Then Continue For
                  If newRow Is Nothing Then Continue For
                  Dim newCell As HSSFCell = newRow.GetCell(k)
                  If newCell Is Nothing Then Continue For
                  newCell.CellStyle.CloneStyleFrom(oldCell.CellStyle)
                  newCell.CellStyle.Alignment = oldCell.CellStyle.Alignment
                  newCell.CellStyle.BorderBottom = oldCell.CellStyle.BorderBottom
                  newCell.CellStyle.BorderLeft = oldCell.CellStyle.BorderLeft
                  newCell.CellStyle.BorderRight = oldCell.CellStyle.BorderRight
                  newCell.CellStyle.BorderTop = oldCell.CellStyle.BorderTop
                  newCell.CellStyle.TopBorderColor = oldCell.CellStyle.TopBorderColor
                  newCell.CellStyle.BottomBorderColor = oldCell.CellStyle.BottomBorderColor
                  newCell.CellStyle.RightBorderColor = oldCell.CellStyle.RightBorderColor
                  newCell.CellStyle.LeftBorderColor = oldCell.CellStyle.LeftBorderColor
                  newCell.CellStyle.FillBackgroundColor = oldCell.CellStyle.FillBackgroundColor
                  newCell.CellStyle.FillForegroundColor = oldCell.CellStyle.FillForegroundColor
                  newCell.CellStyle.DataFormat = oldCell.CellStyle.DataFormat
                  newCell.CellStyle.FillPattern = oldCell.CellStyle.FillPattern
                  newCell.CellStyle.IsHidden = oldCell.CellStyle.IsHidden
                  newCell.CellStyle.Indention = oldCell.CellStyle.Indention
                  newCell.CellStyle.IsLocked = oldCell.CellStyle.IsLocked
                  newCell.CellStyle.Rotation = oldCell.CellStyle.Rotation
                  newCell.CellStyle.VerticalAlignment = oldCell.CellStyle.VerticalAlignment
                  newCell.CellStyle.WrapText = oldCell.CellStyle.WrapText
                  newCell.CellStyle.SetFont(oldCell.CellStyle.GetFont(wk))
                Next
              End If
            Next
          Next


          '取得位置
          '---取得塞Table的位置 new         
          Dim _locat As New List(Of clsExcelSet)
          Dim _QR_Index As New List(Of clsExcelSet)
          Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
          For Each _item In _dicseatitem
            Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
            Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))
            If seat_dic("TYPE") = "QRCODE" Then
              Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
              _QR_Index.Add(_QR_Index_point)
            Else
              Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
              _QR_Index.Add(_QR_Row_point)
            End If
            _locat.Add(_point)
          Next


          '填入資料
          For i = 0 To Allist.Count - 1
            If i = 0 Then
              For j = 0 To _locat.Count - 1
                Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 2 + i + _Tablecount), HSSFRow)
                If _locat(j).TYPE = "STRING" Then
                  hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                  'benny test
                  If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化										
                    Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                    Dim _test = hst.SheetConditionalFormatting
                    Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                    Dim _font = _rule.CreateFontFormatting
                    '_font.SetFontStyle(False, True)
                    'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                    Dim _x = _locat(j).X - 1
                    Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                    'Dim _x = _locat(j).X - 1
                    'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                    Dim _result = ""
                    If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                    Dim _value = _result & _Y & ":" & _result & _Y
                    Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                    hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                  End If
                  '

                ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                  'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                  Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                  Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                  Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                  Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                  anchor.AnchorType = 0
                  Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                Else '-視為integer
                  hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                End If
              Next

            Else
              For j = 0 To _locat.Count - 1
                Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + (i * Interval)), HSSFRow)
                If _locat(j).TYPE = "STRING" Then
                  Try '-test
                    hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                  Catch ex As Exception
                    msgResult_Message = ex.Message
                  End Try

                  'benny test
                  If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                    Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                    Dim _test = hst.SheetConditionalFormatting
                    Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                    Dim _font = _rule.CreateFontFormatting
                    '_font.SetFontStyle(False, True)
                    'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                    Dim _x = _locat(j).X - 1
                    Dim _Y = _locat(j).Y + (i * Interval)
                    'Dim _x = _locat(j).X - 1
                    'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                    Dim _result = ""
                    If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                    Dim _value = _result & _Y & ":" & _result & _Y
                    Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                    hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                  End If
                  '

                ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                  'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                  Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                  Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                  Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                  Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1 + (i * Interval), _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y) + (i * Interval))
                  anchor.AnchorType = 0
                  Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                Else
                  hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                End If
              Next
            End If
          Next

          '轉換分頁線實際行數
          Dim _lstRowrBreak As New List(Of Integer)
          For i = 0 To _RowBreak - 1
            If i = 0 Then '-第一條分頁線例外處理
              _lstRowrBreak.Add((_info("Y") - 1 + Interval - 1) * (i + 1))

            Else
              _lstRowrBreak.Add((_info("Y") - 1 + Interval - 1) + (i * Interval))

            End If
          Next

          '設置分頁線
          If _lstRowrBreak.Count >= 1 Then
            For Each _item In _lstRowrBreak
              hst.SetRowBreak(_item)
            Next
          End If
          'test
          'Dim _ww = hst.RowBreaks.Count
          'Dim sheetname = hst.SheetName
          'test

          '移除複製出來的
          Dim clear = False
          While clear = False
            For i = 0 To wk.Count - 1
              If wk(i).SheetName.Contains("Sample") Then
                wk.RemoveSheetAt(i)
                Exit For
              End If
              If i = wk.Count - 1 Then
                clear = True
              End If
            Next
          End While

        ElseIf _info("TABLE").Contains("CopySheet") Then '-特殊列印
          Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
          Dim Interval = _info("FormInterval")  '-複製間距

          '取得位置
          '---取得塞Table的位置 new         
          Dim _locat As New List(Of clsExcelSet)
          Dim _QR_Index As New List(Of clsExcelSet)
          Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
          For Each _item In _dicseatitem
            Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
            Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
            If seat_dic("TYPE") = "QRCODE" OrElse seat_dic("TYPE") = "DATAMATRIX" Then
              Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
              _QR_Index.Add(_QR_Index_point)
            Else
              Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
              _QR_Index.Add(_QR_Row_point)
            End If
            _locat.Add(_point)
          Next

          ''複製Sheet
          Dim lstSheet As New List(Of HSSFSheet)
          For i = 0 To Allist.Count - 1  '---Table 生成      
            Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
            lstSheet.Add(hst_CopySheet)
          Next

          For i = 0 To lstSheet.Count - 1 '往sheet裡填資料
            '填入資料
            For j = 0 To _locat.Count - 1
              If Allist(i).Count = j Then Exit For
              Dim hr = DirectCast(lstSheet(i).GetRow(_locat(j).Y - 2 + _Tablecount), HSSFRow)
              If _locat(j).TYPE = "STRING" Then
                hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
              ElseIf _locat(j).TYPE = "QRCODE" Then
                'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                anchor.AnchorType = 0
                Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)

              ElseIf _locat(j).TYPE = "DATAMATRIX" Then
                Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                anchor.AnchorType = 0
                Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
              Else
                hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
              End If
            Next
          Next
          '移除複製出來的

          Dim clear = False
          While clear = False
            For i = 0 To wk.Count - 1
              If wk(i).SheetName.Contains("Sheet_") = False Then
                wk.RemoveSheetAt(i)
                Exit For
              End If
              If i = wk.Count - 1 Then
                clear = True
              End If
            Next
          End While
        End If

      Next
    Catch ex As Exception
      msgResult_Message = ex.Message
      Return False
    End Try
  End Function

  Private Function ReadExcelForForm(ByVal sExcelPath As String, ByVal strName As String, ByVal _queue As Queue(Of List(Of List(Of String))), ByVal _list As List(Of String), ByRef msgResult_Message As String) As Boolean
    Dim _count = 0
    Try
      '1.read
      Dim wk As HSSFWorkbook
      Using fs As New FileStream(sExcelPath, FileMode.Open, FileAccess.ReadWrite)
        wk = New HSSFWorkbook(fs) '-抓整個excel
        Dim hst = DirectCast(wk.GetSheetAt(0), HSSFSheet) '---抓一個sheet 理論上應該只有一個Sheet
        Dim sorce_hst = hst.CopySheet("Sample") '---抓一個sheet 理論上應該只有一個Sheet				


        '2.塞資料
        Dim Public_dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary("Public")
        Dim PrintMode = ClsConfig.ReadStringValueDictionary("Public", strName)
        Dim bln_NewMode As Boolean = False
        If PrintMode.ContainsKey("Mode") Then
          If PrintMode("Mode") = "NEW" Then
            bln_NewMode = True
          End If
        End If

        Dim _strSection = Public_dic(strName)
        Dim _dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_strSection)
        Dim _Tablecount = 0
        Dim ruleTablecount = 0

        If bln_NewMode Then
#Region "新版列印"
          '1. 複製sheet
          Dim lstSheet As New List(Of HSSFSheet)
          lstSheet.Clear()
          For i = 1 To _queue(0).Count  '---Table 生成      
            Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
            lstSheet.Add(hst_CopySheet)
          Next
          '2. 取得列印格式

          '3. 往每個sheet填資料
          Dim DisplaceLine = 0
          Dim reDisplaceLine = 0
          For i = 0 To lstSheet.Count - 1
            DisplaceLine = 0
            '根據範本取得本sheet需要分成幾層
            Dim k = -1
            For Each item In _dic

              k += 1
              Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)
              '3.1 取得塞Table的位置
              Dim _locat As New List(Of clsExcelSet)
              Dim _QR_Index As New List(Of clsExcelSet)
              Dim _Bar_Index As New List(Of clsExcelSet)
              Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
              For Each _item In _dicseatitem
                Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                If seat_dic("TYPE") = "QRCODE" Then
                  Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_QR_Index_point)
                ElseIf seat_dic("TYPE") = "BARCODE" Then
                  Dim _Bar_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_Bar_Index_point)
                Else
                  Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_QR_Row_point)
                End If
                _locat.Add(_point)
              Next


              '3.2 填入資料
              DisplaceLine = reDisplaceLine
              If _queue(k) Is Nothing Then Continue For

              For m = 0 To _queue(k)(i).Count - 1
                m -= 1
                Dim bln_check_CopyRow As Boolean = True '確認是否該複製
                For j = 0 To _locat.Count - 1
                  m += 1
                  If _queue(k)(i).Count = m Then Exit For

                  '位移Y軸
                  _Tablecount = Int(m / (_locat.Count)) + 1 + DisplaceLine
                  '檢查是否複製
                  If bln_check_CopyRow AndAlso m > _locat.Count - 1 Then
                    bln_check_CopyRow = False
                    '該層開始前 檢查是否需增加欄位
                    'lstSheet(i).CopyRow(_locat(j).Y - 2 + _Tablecount - 1, _locat(j).Y - 2 + _Tablecount)
                    'reDisplaceLine += 1
                  End If

                  Dim hr = DirectCast(lstSheet(i).GetRow(_locat(j).Y - 2 + _Tablecount), HSSFRow)
                  If _locat(j).TYPE = "STRING" Then
                    hr.Cells(_locat(j).X - 1).SetCellValue(_queue(k)(i)(m))
                  ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                    Dim _VALUE = CodeEncoderFromString(_queue(k)(i)(m), _locat(j).TYPE) '-要塞進去QRCode的值
                    Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                    Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                    Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                    anchor.AnchorType = 0
                    Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                  ElseIf _locat(j).TYPE = "BARCODE" Then
                    Dim _VALUE = GenCode128.Code128Rendering.MakeBarcodeImage(_queue(k)(i)(m), 2, True)
                    Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                    Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                    Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                    anchor.AnchorType = 0
                    Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                  ElseIf _locat(j).TYPE = "" Then
                    '如果為空則不寫入
                  Else
                    hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(_queue(k)(i)(m)))
                  End If
                Next
              Next
            Next
          Next

          '4. 移除多餘的sheet
          Dim clear = False
          While clear = False
            For i = 0 To wk.Count - 1
              If wk(i).SheetName.Contains("Sheet_") = False Then
                wk.RemoveSheetAt(i)
                Exit For
              End If
              If i = wk.Count - 1 Then
                clear = True
              End If
            Next
          End While

#End Region
        Else
#Region "舊版列印"
          '確認要做幾次
          Dim BaseRowFrom = 0
          Dim BaseRowCount = hst.LastRowNum
          Dim CountTime As Integer = 1
          If _dic.Count > 1 Then
            CountTime = Int16.Parse(_list.Count / _dic.Count)
          End If
          For x = 1 To CountTime - 1
            For y = 0 To BaseRowCount
              hst.GetRow(y).CopyRowTo(hst.LastRowNum + 1)
            Next
          Next
          For x = 0 To CountTime - 1
            _Tablecount = 0
            If BaseRowFrom <> 0 Then
              hst.SetRowBreak(BaseRowFrom - 1)
              _Tablecount = _Tablecount + BaseRowFrom
            End If
            For Each item In _dic
              Dim LastRowBreakIndex = 0
              If hst.RowBreaks.Any Then LastRowBreakIndex = hst.RowBreaks(hst.RowBreaks.Length - 1)
              Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)
              If _info("TABLE") = "0" Then
                '先確認當前Row是否會超出標籤範圍
                If CInt(PRINTER_Label_Limited) <> 0 Then
                  If _info("TYPE") = "STRING" AndAlso (_info("Y") - 1 + _Tablecount - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                    hst.SetRowBreak(_info("Y") - 1 + _Tablecount - 1)
                  ElseIf (_info("TYPE") = "QRCODE" Or _info("TYPE") = "DATAMATRIX") AndAlso (_info("Y") + 5 + _Tablecount - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                    hst.SetRowBreak(_info("Y") - 1 + _Tablecount - 1)
                  End If
                End If
                Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + _Tablecount), HSSFRow)
                Dim _cellcount = hr.Cells.Count
                'hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
                If _info("TYPE") = "STRING" Then
                  hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
                ElseIf _info("TYPE") = "QRCODE" Or _info("TYPE") = "DATAMATRIX" Then
                  Dim _VALUE = CodeEncoderFromString(_list(_count), _info("TYPE")) '-要塞進去QRCode的值
                  Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                  Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                  'Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1, _info("X") + 1, _info("Y") + 4)
                  'Gary:可能會影響到其他標籤列印要測試
                  Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1 + _Tablecount, _info("X"), _info("Y") + 5 + _Tablecount)
                  anchor.AnchorType = 0
                  Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                  'Reset the image to the original size.(若沒設定這個，需要注意HSSFClientAnchor最後兩個參數)
                  'picture.Resize()
                Else
                  If _list(_count) <> "" Then
                    hr.Cells(_info("X") - 1).SetCellValue(Convert.ToDouble(_list(_count)))
                  End If
                End If


                _count = _count + 1

                For i = 0 To wk.Count - 1
                  If wk(i).SheetName.Contains("Sample") Then
                    wk.RemoveSheetAt(i)
                    Exit For
                  End If
                Next
              ElseIf _info("TABLE").Contains("Table") AndAlso _list(_count) <> "" Then
                Dim TableLst = Split(_list(_count), ",")
                _count = _count + 1
                For i = 0 To (TableLst.Count - 2) ' - 2 '---Table 生成(全部n列，就新增n-1列，因為原本的範例就有一列了)         
                  If _info("Y") = 0 Then Continue For
                  Dim tmp1 = IIf(i = 0, 0, 1)
                  hst.CopyRow(_info("Y") + _Tablecount + tmp1, _info("Y") + _Tablecount + tmp1 + 1) ' + _Tablecount)


                Next
                '---取得塞Table的位置 new         
                Dim _locat As New List(Of clsExcelSet)
                Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                For Each _item In _dicseatitem
                  Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                  Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))

                  _locat.Add(_point)
                Next
                '---塞table new
                For i = 0 To TableLst.Count - 1
                  Dim tmp1 = IIf(i = 0, 0, 1)
                  Dim tmp2 = 1
                  'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
                  Dim TableLstDtl = Split(TableLst(i), "/")
                  For j = 0 To TableLstDtl.Count - 1
                    If hst.RowBreaks.Any Then LastRowBreakIndex = hst.RowBreaks(hst.RowBreaks.Length - 1)
                    If CInt(PRINTER_Label_Limited) <> 0 Then
                      If _locat(j).TYPE = "STRING" AndAlso (_locat(j).Y - 1 + i + _Tablecount - 1 + tmp2 - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                        hst.SetRowBreak(_locat(j).Y - 1 + i + _Tablecount - 1 + tmp2 - 1)
                      End If
                    End If
                    Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount - 1 + tmp2), HSSFRow)
                    If _locat(j).TYPE = "STRING" Then
                      'benny test
                      If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                        Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                        Dim _test = hst.SheetConditionalFormatting
                        Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                        Dim _font = _rule.CreateFontFormatting
                        '_font.SetFontStyle(False, True)
                        'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                        Dim _x = _locat(j).X - 1
                        Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                        Dim _result = ""
                        If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                        Dim _value = _result & _Y & ":" & _result & _Y
                        Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                        hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                      End If
                      'benny test											
                      hr.Cells(_locat(j).X - 1).SetCellValue(TableLstDtl(j))

                    Else
                      If TableLstDtl(j) <> "" Then
                        hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToDouble(TableLstDtl(j)))
                        'benny test
                        If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                          Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                          Dim _test = hst.SheetConditionalFormatting
                          Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                          Dim _font = _rule.CreateFontFormatting
                          '_font.SetFontStyle(False, True)
                          'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                          Dim _x = _locat(j).X - 1
                          Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                          Dim _result = ""
                          If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                          Dim _value = _result & _Y & ":" & _result & _Y
                          Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                          hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                        End If
                        'benny test
                      End If
                    End If
                  Next
                Next
                _Tablecount = _Tablecount + TableLst.Count - 1
              Else
                _count = _count + 1
              End If
              While _queue.Count > 0
                If _queue.Count = 0 Then Exit For
                Dim Allist = _queue.Dequeue '---取得資料
                Dim tmp1 = IIf(ruleTablecount = 2, 1, 0)
                Dim tmp2 = IIf(ruleTablecount = 2, 1, 0)

                If _info("TABLE").Contains("Table") Then  '---TABLE表    '---TABLE表          
                  For i = 0 To (Allist.Count - 1) ' - 1 '---Table 生成(全部n列，就新增n-1列，因為原本的範例就有一列了)         
                    If _info("Y") = 0 Then Continue For
                    hst.CopyRow(_info("Y") - 2 + _Tablecount + tmp1, _info("Y") - 1 + _Tablecount + tmp1) ' + _Tablecount)

                  Next

                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))

                    _locat.Add(_point)
                  Next

                  '---塞table new
                  For i = 0 To Allist.Count - 1
                    'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
                    For j = 0 To Allist(i).Count - 1
                      Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount + tmp2), HSSFRow)
                      If _locat(j).TYPE = "STRING" Then
                        'benny test
                        If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                          Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                          Dim _test = hst.SheetConditionalFormatting
                          Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                          Dim _font = _rule.CreateFontFormatting
                          '_font.SetFontStyle(False, True)
                          'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                          Dim _x = _locat(j).X - 1
                          Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                          Dim _result = ""
                          If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                          Dim _value = _result & _Y & ":" & _result & _Y
                          Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                          hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                        End If
                        'benny test											
                        hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))

                      Else
                        If Allist(i)(j) <> "" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToDouble(Allist(i)(j)))
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          'benny test
                        End If
                      End If
                    Next
                  Next
                  '若上一個Table有n列，則下一個Table位移n-1列(因為新增加了n-1列)
                  _Tablecount = _Tablecount + (Allist.Count - 1)
                  ruleTablecount += 1
                ElseIf _info("TABLE").Contains("Form") Then '-特殊列印
                  'Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
                  _Tablecount = 1
                  Dim Interval As Integer = _info("FormInterval")  '-複製間距
                  Dim ColumnRange As Integer = _info("ColumnRange") '確認列印寬度
                  '刪除原本的所有合併儲存格
                  Dim MergedCount = hst.NumMergedRegions
                  For i = 0 To MergedCount
                    hst.RemoveMergedRegion(0)
                  Next
                  ''複製格式
                  'For i = 0 To Allist.Count - 1 - 1 '---Table 生成      
                  'For j = 0 To Interval - 1
                  '	hst.CopyRow(_info("Y") + j - 1, _info("Y") + j + Interval - 1)


                  '||要採用跑一次就把所有Row複製出來
                  For j = 1 To Allist.Count - 1
                    For k = 0 To Interval - 1
                      hst.GetRow(k).CopyRowTo(hst.LastRowNum + 1)
                    Next
                  Next

                  '設定格式化的條件
                  Dim ConditionalCount = sorce_hst.SheetConditionalFormatting.NumConditionalFormattings
                  For L = 0 To ConditionalCount - 1
                    Dim cf = sorce_hst.SheetConditionalFormatting.GetConditionalFormattingAt(L)
                    hst.SheetConditionalFormatting.AddConditionalFormatting(cf)
                  Next

                  'Next

                  '設定長寬
                  For i = 0 To Allist.Count - 1 - 1
                    For j = 0 To Interval - 1
                      hst.GetRow(_info("Y") + j + Interval - 1 + (i * Interval)).Height = hst.GetRow(_info("Y") + j - 1).Height
                    Next
                  Next

                  '設定合併儲存格
                  For i = 0 To Allist.Count - 1
                    '取得第一組範例的合併儲存格資訊
                    MergedCount = sorce_hst.NumMergedRegions
                    For k = 0 To MergedCount - 1
                      Dim MergedRegionAt = sorce_hst.GetMergedRegion(k)
                      Dim toFirstRow = MergedRegionAt.FirstRow + Interval * i
                      Dim toLastRow = MergedRegionAt.LastRow + Interval * i
                      Dim toFirstCol = MergedRegionAt.FirstColumn
                      Dim toLastCol = MergedRegionAt.LastColumn
                      Dim toMergedRegionAt As New CellRangeAddress(toFirstRow, toLastRow, toFirstCol, toLastCol)
                      hst.AddMergedRegion(toMergedRegionAt)
                    Next

                    '儲存格格式設定
                    For j = 0 To Interval - 1
                      Dim newRow As HSSFRow = hst.GetRow(_info("Y") + j + Interval - 1)
                      Dim sourceRow As HSSFRow = sorce_hst.GetRow(_info("Y") + j - 1)
                      For k = 0 To sourceRow.LastCellNum - 1
                        Dim oldCell As HSSFCell = sourceRow.GetCell(k)
                        If oldCell.ToString = "" Then Continue For
                        If newRow Is Nothing Then Continue For
                        Dim newCell As HSSFCell = newRow.GetCell(k)
                        If newCell Is Nothing Then Continue For
                        newCell.CellStyle.CloneStyleFrom(oldCell.CellStyle)
                        newCell.CellStyle.Alignment = oldCell.CellStyle.Alignment
                        newCell.CellStyle.BorderBottom = oldCell.CellStyle.BorderBottom
                        newCell.CellStyle.BorderLeft = oldCell.CellStyle.BorderLeft
                        newCell.CellStyle.BorderRight = oldCell.CellStyle.BorderRight
                        newCell.CellStyle.BorderTop = oldCell.CellStyle.BorderTop
                        newCell.CellStyle.TopBorderColor = oldCell.CellStyle.TopBorderColor
                        newCell.CellStyle.BottomBorderColor = oldCell.CellStyle.BottomBorderColor
                        newCell.CellStyle.RightBorderColor = oldCell.CellStyle.RightBorderColor
                        newCell.CellStyle.LeftBorderColor = oldCell.CellStyle.LeftBorderColor
                        newCell.CellStyle.FillBackgroundColor = oldCell.CellStyle.FillBackgroundColor
                        newCell.CellStyle.FillForegroundColor = oldCell.CellStyle.FillForegroundColor
                        newCell.CellStyle.DataFormat = oldCell.CellStyle.DataFormat
                        newCell.CellStyle.FillPattern = oldCell.CellStyle.FillPattern
                        newCell.CellStyle.IsHidden = oldCell.CellStyle.IsHidden
                        newCell.CellStyle.Indention = oldCell.CellStyle.Indention
                        newCell.CellStyle.IsLocked = oldCell.CellStyle.IsLocked
                        newCell.CellStyle.Rotation = oldCell.CellStyle.Rotation
                        newCell.CellStyle.VerticalAlignment = oldCell.CellStyle.VerticalAlignment
                        newCell.CellStyle.WrapText = oldCell.CellStyle.WrapText
                        newCell.CellStyle.SetFont(oldCell.CellStyle.GetFont(wk))
                      Next
                    Next
                  Next


                  '取得位置
                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _QR_Index As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))
                    If seat_dic("TYPE") = "QRCODE" Then
                      Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Index_point)
                      _QR_Index_point = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y") - 1, "STRING")
                      _QR_Index.Add(_QR_Index_point)
                    Else
                      Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Row_point)
                    End If
                    _locat.Add(_point)
                  Next

                  Dim RowBreakIndex As Integer = 0
                  '填入資料
                  For i = 0 To Allist.Count - 1
                    If i = 0 Then
                      For j = 0 To _locat.Count - 1
                        Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 2 + i + _Tablecount), HSSFRow)
                        If _locat(j).TYPE = "STRING" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                          RowBreakIndex = hr.RowNum
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                            'Dim _x = _locat(j).X - 1
                            'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          '

                        ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                          'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                          Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                          Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                          Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                          Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                          anchor.AnchorType = 0
                          Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                        Else '-視為integer
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                        End If
                      Next

                    Else
                      '不是第一組就先插一條分頁線
                      hst.SetRowBreak(RowBreakIndex)

                      For j = 0 To _locat.Count - 1
                        Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + (i * Interval)), HSSFRow)
                        If _locat(j).TYPE = "STRING" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                          RowBreakIndex = hr.RowNum
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y + (i * Interval)
                            'Dim _x = _locat(j).X - 1
                            'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          '

                        ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                          'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                          Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                          Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                          Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                          Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1 + (i * Interval), _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y) + (i * Interval))
                          anchor.AnchorType = 0
                          Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                        Else
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                        End If
                      Next
                    End If
                  Next
                  '移除複製出來的
                  Dim clear = False
                  While clear = False
                    For i = 0 To wk.Count - 1
                      If wk(i).SheetName.Contains("Sample") Then
                        wk.RemoveSheetAt(i)
                        Exit For
                      End If
                      If i = wk.Count - 1 Then
                        clear = True
                      End If
                    Next
                  End While
                  wk.SetPrintArea(0, 0, ColumnRange, 0, RowBreakIndex)
                ElseIf _info("TABLE").Contains("CopySheet") Then '-特殊列印
                  Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
                  Dim Interval = _info("FormInterval")  '-複製間距

                  '取得位置
                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _QR_Index As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                    If seat_dic("TYPE") = "QRCODE" OrElse seat_dic("TYPE") = "DATAMATRIX" Then
                      Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Index_point)
                    Else
                      Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Row_point)
                    End If
                    _locat.Add(_point)
                  Next

                  ''複製Sheet
                  Dim lstSheet As New List(Of HSSFSheet)
                  For i = 0 To Allist.Count - 1  '---Table 生成      
                    Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
                    lstSheet.Add(hst_CopySheet)
                  Next

                  For i = 0 To lstSheet.Count - 1 '往sheet裡填資料
                    '填入資料
                    For j = 0 To _locat.Count - 1
                      If Allist(i).Count = j Then Exit For
                      Dim hr = DirectCast(lstSheet(i).GetRow(_locat(j).Y - 2 + _Tablecount), HSSFRow)
                      If _locat(j).TYPE = "STRING" Then
                        hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                      ElseIf _locat(j).TYPE = "QRCODE" Then
                        'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                        Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                        Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                        Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                        Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                        anchor.AnchorType = 0
                        Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)

                      ElseIf _locat(j).TYPE = "DATAMATRIX" Then
                        Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                        Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                        Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                        Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                        anchor.AnchorType = 0
                        Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                      Else
                        hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                      End If
                    Next
                  Next
                  '移除複製出來的

                  Dim clear = False
                  While clear = False
                    For i = 0 To wk.Count - 1
                      If wk(i).SheetName.Contains("Sheet_") = False Then
                        wk.RemoveSheetAt(i)
                        Exit For
                      End If
                      If i = wk.Count - 1 Then
                        clear = True
                      End If
                    Next
                  End While
                End If
              End While
            Next
            BaseRowFrom = BaseRowCount + _Tablecount + 1
          Next
#End Region
        End If


      End Using
      '3.寫回檔案
      Dim file As New FileStream(sExcelPath, FileMode.Create)
      '產生檔案
      wk.Write(file)
      file.Close()

      Return True
    Catch ex As Exception
      msgResult_Message = ex.Message
      Return False
    End Try
  End Function


  Private Function ReadExcel(ByVal sExcelPath As String, ByVal strName As String, ByVal _queue As Queue(Of List(Of List(Of String))), ByVal _list As List(Of String), ByRef msgResult_Message As String) As Boolean
    Dim _count = 0
    Try
      '1.read
      Dim wk As HSSFWorkbook
      Using fs As New FileStream(sExcelPath, FileMode.Open, FileAccess.ReadWrite)
        wk = New HSSFWorkbook(fs) '-抓整個excel
        Dim hst = DirectCast(wk.GetSheetAt(0), HSSFSheet) '---抓一個sheet 理論上應該只有一個Sheet
        Dim sorce_hst = hst.CopySheet("Sample") '---抓一個sheet 理論上應該只有一個Sheet				


        '2.塞資料
        Dim Public_dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary("Public")
        Dim PrintMode = ClsConfig.ReadStringValueDictionary("Public", strName)
        Dim bln_NewMode As Boolean = False
        If PrintMode.ContainsKey("Mode") Then
          If PrintMode("Mode") = "NEW" Then
            bln_NewMode = True
          End If
        End If

        Dim _strSection = Public_dic(strName)
        Dim _dic As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_strSection)
        Dim _Tablecount = 0
        Dim ruleTablecount = 0

        If bln_NewMode Then
#Region "新版列印"
          '1. 複製sheet
          Dim lstSheet As New List(Of HSSFSheet)
          lstSheet.Clear()
          For i = 1 To _queue(0).Count  '---Table 生成      
            Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
            lstSheet.Add(hst_CopySheet)
          Next
          '2. 取得列印格式

          '3. 往每個sheet填資料
          Dim DisplaceLine = 0
          Dim reDisplaceLine = 0
          For i = 0 To lstSheet.Count - 1
            DisplaceLine = 0
            '根據範本取得本sheet需要分成幾層
            Dim k = -1
            For Each item In _dic

              k += 1
              Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)
              '3.1 取得塞Table的位置
              Dim _locat As New List(Of clsExcelSet)
              Dim _QR_Index As New List(Of clsExcelSet)
              Dim _Bar_Index As New List(Of clsExcelSet)
              Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
              For Each _item In _dicseatitem
                Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                If seat_dic("TYPE") = "QRCODE" Then
                  Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_QR_Index_point)
                ElseIf seat_dic("TYPE") = "BARCODE" Then
                  Dim _Bar_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_Bar_Index_point)
                Else
                  Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                  _QR_Index.Add(_QR_Row_point)
                End If
                _locat.Add(_point)
              Next


              '3.2 填入資料
              DisplaceLine = reDisplaceLine
              If _queue(k) Is Nothing Then Continue For

              For m = 0 To _queue(k)(i).Count - 1
                m -= 1
                Dim bln_check_CopyRow As Boolean = True '確認是否該複製
                For j = 0 To _locat.Count - 1
                  m += 1
                  If _queue(k)(i).Count = m Then Exit For

                  '位移Y軸
                  _Tablecount = Int(m / (_locat.Count)) + 1 + DisplaceLine
                  '檢查是否複製
                  If bln_check_CopyRow AndAlso m > _locat.Count - 1 Then
                    bln_check_CopyRow = False
                    '該層開始前 檢查是否需增加欄位
                    'lstSheet(i).CopyRow(_locat(j).Y - 2 + _Tablecount - 1, _locat(j).Y - 2 + _Tablecount)
                    'reDisplaceLine += 1
                  End If

                  Dim hr = DirectCast(lstSheet(i).GetRow(_locat(j).Y - 2 + _Tablecount), HSSFRow)
                  If _locat(j).TYPE = "STRING" Then
                    hr.Cells(_locat(j).X - 1).SetCellValue(_queue(k)(i)(m))
                  ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                    If _queue(k)(i)(m) <> "" Then
                      Dim _VALUE = CodeEncoderFromString(_queue(k)(i)(m), _locat(j).TYPE) '-要塞進去QRCode的值
                      If _VALUE IsNot Nothing Then
                        Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                        Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                        Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                        anchor.AnchorType = 0
                        Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                      End If
                    End If
                  ElseIf _locat(j).TYPE = "BARCODE" Then
                    If _queue(k)(i)(m) <> "" Then
                      Dim _VALUE = GenCode128.Code128Rendering.MakeBarcodeImage(_queue(k)(i)(m), 2, True)
                      If _VALUE IsNot Nothing Then
                        Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                        Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                        Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                        anchor.AnchorType = 0
                        Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                      End If
                    End If
                  ElseIf _locat(j).TYPE = "" Then
                    '如果為空則不寫入
                  Else
                      hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(_queue(k)(i)(m)))
                  End If
                Next
              Next
            Next
          Next

          '4. 移除多餘的sheet
          Dim clear = False
          While clear = False
            For i = 0 To wk.Count - 1
              If wk(i).SheetName.Contains("Sheet_") = False Then
                wk.RemoveSheetAt(i)
                Exit For
              End If
              If i = wk.Count - 1 Then
                clear = True
              End If
            Next
          End While

#End Region
        Else
#Region "舊版列印"
          '確認要做幾次
          Dim BaseRowFrom = 0
          Dim BaseRowCount = hst.LastRowNum
          Dim CountTime As Integer = 1
          If _dic.Count > 1 Then
            CountTime = Int16.Parse(_list.Count / _dic.Count)
          End If
          For x = 1 To CountTime - 1
            For y = 0 To BaseRowCount
              hst.GetRow(y).CopyRowTo(hst.LastRowNum + 1)
              hst.GetRow(hst.LastRowNum).Height = hst.GetRow(y).Height
            Next
          Next
          For x = 0 To CountTime - 1
            _Tablecount = 0
            If BaseRowFrom <> 0 Then
              hst.SetRowBreak(BaseRowFrom - 1)
              _Tablecount = _Tablecount + BaseRowFrom
            End If
            For Each item In _dic
              Dim LastRowBreakIndex = 0
              If hst.RowBreaks.Any Then LastRowBreakIndex = hst.RowBreaks(hst.RowBreaks.Length - 1)
              Dim _info = ClsConfig.ReadStringValueDictionary(_strSection, item.Key)
              If _info("TABLE") = "0" Then
                '確認是否為有效資料
                If (_info("TYPE") Is Nothing OrElse _info("TYPE") = "NONE") Then Continue For
                '確認當前Row是否會超出標籤範圍
                If CInt(PRINTER_Label_Limited) <> 0 Then
                  If _info("TYPE") = "STRING" AndAlso (_info("Y") - 1 + _Tablecount - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                    hst.SetRowBreak(_info("Y") - 1 + _Tablecount - 1)
                  ElseIf (_info("TYPE") = "QRCODE" Or _info("TYPE") = "DATAMATRIX") AndAlso (Int16.Parse(_info("Y")) + Int16.Parse(_info("INDEX_Y")) + _Tablecount - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                    hst.SetRowBreak(_info("Y") - 1 + _Tablecount - 1)
                  End If
                End If
                Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + _Tablecount), HSSFRow)
                Dim _cellcount = hr.Cells.Count
                'hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
                If _info("TYPE") = "STRING" Then
                  hr.Cells(_info("X") - 1).SetCellValue(_list(_count))
                ElseIf _info("TYPE") = "QRCODE" Or _info("TYPE") = "DATAMATRIX" Or _info("TYPE") = "Code128" Or _info("TYPE") = "Code39" Or _info("TYPE") = "CODABAR" Then
                  If Not _list(_count) = "" Then
                    Dim _VALUE = CodeEncoderFromString(_list(_count), _info("TYPE"), _info("INDEX_X")) '-要塞進去QRCode的值
                    If _VALUE IsNot Nothing Then
                      Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                      Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                      'Gary:可能會影響到其他標籤列印要測試
                      Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1 + _Tablecount + hst.RowBreaks.Where(Function(f) f < _info("Y") - 1 + _Tablecount AndAlso f > BaseRowFrom).Count, Int16.Parse(_info("X")) + Int16.Parse(_info("INDEX_X")) - 1, Int16.Parse(_info("Y")) + Int16.Parse(_info("INDEX_Y")) - 1 + _Tablecount + hst.RowBreaks.Where(Function(f) f < _info("Y") - 1 + _Tablecount AndAlso f > BaseRowFrom).Count)
                      anchor.AnchorType = 0
                      Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                      'Reset the image to the original size.(若沒設定這個，需要注意HSSFClientAnchor最後兩個參數)
                      'picture.Resize()
                    End If
                  End If
                ElseIf _info("TYPE") = "IMAGE" Then
                  If Not _list(_count) = "" Then
                    Dim _VALUE As New Drawing.Bitmap(Windows.Forms.Application.StartupPath & _list(_count))
                    If _VALUE IsNot Nothing Then
                      Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                      Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                      'Gary:可能會影響到其他標籤列印要測試
                      Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _info("X") - 1, _info("Y") - 1 + _Tablecount + hst.RowBreaks.Where(Function(f) f < _info("Y") - 1 + _Tablecount AndAlso f > BaseRowFrom).Count, Int16.Parse(_info("X")) + Int16.Parse(_info("INDEX_X")) - 1, Int16.Parse(_info("Y")) + Int16.Parse(_info("INDEX_Y")) - 1 + _Tablecount + hst.RowBreaks.Where(Function(f) f < _info("Y") - 1 + _Tablecount AndAlso f > BaseRowFrom).Count)
                      anchor.AnchorType = 0
                      Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                      picture.Resize()
                      _VALUE.Dispose()
                      'Reset the image to the original size.(若沒設定這個，需要注意HSSFClientAnchor最後兩個參數)
                      'picture.Resize()
                    End If
                  End If
                ElseIf _info("TYPE") = "NEWPAGE" Then
                  hst.SetRowBreak(_info("Y") - 1 + _Tablecount - 1)
                Else
                  If _list(_count) <> "" Then
                    hr.Cells(_info("X") - 1).SetCellValue(Convert.ToDouble(_list(_count)))
                  End If
                End If


                _count = _count + 1

                For i = 0 To wk.Count - 1
                  If wk(i).SheetName.Contains("Sample") Then
                    wk.RemoveSheetAt(i)
                    Exit For
                  End If
                Next
              ElseIf _info("TABLE").Contains("Table") AndAlso _list(_count) <> "" Then
                Dim TableLst = Split(_list(_count), ",")
                _count = _count + 1
                For i = 0 To (TableLst.Count - 2) ' - 2 '---Table 生成(全部n列，就新增n-1列，因為原本的範例就有一列了)         
                  If _info("Y") = 0 Then Continue For
                  Dim tmp1 As Integer = IIf(i = 0, 0, 1)
                  hst.CopyRow(_info("Y") + _Tablecount + tmp1, _info("Y") + _Tablecount + tmp1 + 1) ' + _Tablecount)
                  '複製後,要把這以下的高度全部重設
                  For a = hst.LastRowNum To (Integer.Parse(_info("Y")) + _Tablecount + tmp1) Step -1
                    hst.GetRow(a).Height = hst.GetRow(a - 1).Height
                  Next

                Next
                '---取得塞Table的位置 new         
                Dim _locat As New List(Of clsExcelSet)
                Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                For Each _item In _dicseatitem
                  Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                  Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))

                  _locat.Add(_point)
                Next
                '---塞table new
                Dim EndBreak = 1
                For i = 0 To TableLst.Count - 1
                  Dim tmp1 = IIf(i = 0, 0, 1)
                  'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
                  Dim TableLstDtl = Split(TableLst(i), "/")
                  For j = 0 To TableLstDtl.Count - 1
                    EndBreak = _locat(j).Y + i + _Tablecount
                    If hst.RowBreaks.Any Then LastRowBreakIndex = hst.RowBreaks(hst.RowBreaks.Length - 1)
                    If CInt(PRINTER_Label_Limited) <> 0 Then
                      If _locat(j).TYPE = "STRING" AndAlso (_locat(j).Y - 1 + i + _Tablecount - LastRowBreakIndex > CInt(PRINTER_Label_Limited)) Then
                        hst.SetRowBreak(_locat(j).Y - 1 + i + _Tablecount - 1)
                      End If
                    End If
                    Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount), HSSFRow)
                    If _locat(j).TYPE = "STRING" Then
                      'benny test
                      If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                        Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                        Dim _test = hst.SheetConditionalFormatting
                        Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                        Dim _font = _rule.CreateFontFormatting
                        '_font.SetFontStyle(False, True)
                        'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                        Dim _x = _locat(j).X - 1
                        Dim _Y = _locat(j).Y + i + _Tablecount
                        Dim _result = ""
                        If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                        Dim _value = _result & _Y & ":" & _result & _Y
                        Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                        hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                      End If
                      'benny test											
                      hr.Cells(_locat(j).X - 1).SetCellValue(TableLstDtl(j))

                    Else
                      If TableLstDtl(j) <> "" Then
                        hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToDouble(TableLstDtl(j)))
                        'benny test
                        If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                          Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                          Dim _test = hst.SheetConditionalFormatting
                          Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                          Dim _font = _rule.CreateFontFormatting
                          '_font.SetFontStyle(False, True)
                          'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                          Dim _x = _locat(j).X - 1
                          Dim _Y = _locat(j).Y + i + _Tablecount
                          Dim _result = ""
                          If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                          Dim _value = _result & _Y & ":" & _result & _Y
                          Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                          hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                        End If
                        'benny test
                      End If
                    End If
                  Next
                Next
                If (Not hst.LastRowNum = EndBreak + 1) AndAlso (_info("TABLE") <> "PackageinfoMR_Table2") Then
                  hst.SetRowBreak(EndBreak - 1)
                End If
                _Tablecount = _Tablecount + TableLst.Count - 1
              Else
                _count = _count + 1
              End If
              While _queue.Count > 0
                If _queue.Count = 0 Then Exit For
                Dim Allist = _queue.Dequeue '---取得資料
                Dim tmp1 = IIf(ruleTablecount = 2, 1, 0)
                Dim tmp2 = IIf(ruleTablecount = 2, 1, 0)

                If _info("TABLE").Contains("Table") Then  '---TABLE表    '---TABLE表          
                  For i = 0 To (Allist.Count - 1) ' - 1 '---Table 生成(全部n列，就新增n-1列，因為原本的範例就有一列了)         
                    If _info("Y") = 0 Then Continue For
                    hst.CopyRow(_info("Y") - 2 + _Tablecount + tmp1, _info("Y") - 1 + _Tablecount + tmp1) ' + _Tablecount)

                  Next

                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))

                    _locat.Add(_point)
                  Next

                  '---塞table new
                  For i = 0 To Allist.Count - 1
                    'Dim hr = DirectCast(hst.GetRow(_info("Y") - 1 + i), HSSFRow)
                    For j = 0 To Allist(i).Count - 1
                      Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + i + _Tablecount + tmp2), HSSFRow)
                      If _locat(j).TYPE = "STRING" Then
                        'benny test
                        If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                          Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                          Dim _test = hst.SheetConditionalFormatting
                          Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                          Dim _font = _rule.CreateFontFormatting
                          '_font.SetFontStyle(False, True)
                          'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                          Dim _x = _locat(j).X - 1
                          Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                          Dim _result = ""
                          If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                          Dim _value = _result & _Y & ":" & _result & _Y
                          Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                          hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                        End If
                        'benny test											
                        hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))

                      Else
                        If Allist(i)(j) <> "" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToDouble(Allist(i)(j)))
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          'benny test
                        End If
                      End If
                    Next
                  Next
                  '若上一個Table有n列，則下一個Table位移n-1列(因為新增加了n-1列)
                  _Tablecount = _Tablecount + (Allist.Count - 1)
                  ruleTablecount += 1
                ElseIf _info("TABLE").Contains("Form") Then '-特殊列印
                  'Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
                  _Tablecount = 1
                  Dim Interval As Integer = _info("FormInterval")  '-複製間距
                  Dim ColumnRange As Integer = _info("ColumnRange") '確認列印寬度
                  '刪除原本的所有合併儲存格
                  Dim MergedCount = hst.NumMergedRegions
                  For i = 0 To MergedCount
                    hst.RemoveMergedRegion(0)
                  Next
                  ''複製格式
                  'For i = 0 To Allist.Count - 1 - 1 '---Table 生成      
                  'For j = 0 To Interval - 1
                  '	hst.CopyRow(_info("Y") + j - 1, _info("Y") + j + Interval - 1)


                  '||要採用跑一次就把所有Row複製出來
                  For j = 1 To Allist.Count - 1
                    For k = 0 To Interval - 1
                      hst.GetRow(k).CopyRowTo(hst.LastRowNum + 1)
                    Next
                  Next

                  '設定格式化的條件
                  Dim ConditionalCount = sorce_hst.SheetConditionalFormatting.NumConditionalFormattings
                  For L = 0 To ConditionalCount - 1
                    Dim cf = sorce_hst.SheetConditionalFormatting.GetConditionalFormattingAt(L)
                    hst.SheetConditionalFormatting.AddConditionalFormatting(cf)
                  Next

                  'Next

                  '設定長寬
                  For i = 0 To Allist.Count - 1 - 1
                    For j = 0 To Interval - 1
                      hst.GetRow(_info("Y") + j + Interval - 1 + (i * Interval)).Height = hst.GetRow(_info("Y") + j - 1).Height
                    Next
                  Next

                  '設定合併儲存格
                  For i = 0 To Allist.Count - 1
                    '取得第一組範例的合併儲存格資訊
                    MergedCount = sorce_hst.NumMergedRegions
                    For k = 0 To MergedCount - 1
                      Dim MergedRegionAt = sorce_hst.GetMergedRegion(k)
                      Dim toFirstRow = MergedRegionAt.FirstRow + Interval * i
                      Dim toLastRow = MergedRegionAt.LastRow + Interval * i
                      Dim toFirstCol = MergedRegionAt.FirstColumn
                      Dim toLastCol = MergedRegionAt.LastColumn
                      Dim toMergedRegionAt As New CellRangeAddress(toFirstRow, toLastRow, toFirstCol, toLastCol)
                      hst.AddMergedRegion(toMergedRegionAt)
                    Next

                    '儲存格格式設定
                    For j = 0 To Interval - 1
                      Dim newRow As HSSFRow = hst.GetRow(_info("Y") + j + Interval - 1)
                      Dim sourceRow As HSSFRow = sorce_hst.GetRow(_info("Y") + j - 1)
                      For k = 0 To sourceRow.LastCellNum - 1
                        Dim oldCell As HSSFCell = sourceRow.GetCell(k)
                        If oldCell.ToString = "" Then Continue For
                        If newRow Is Nothing Then Continue For
                        Dim newCell As HSSFCell = newRow.GetCell(k)
                        If newCell Is Nothing Then Continue For
                        newCell.CellStyle.CloneStyleFrom(oldCell.CellStyle)
                        newCell.CellStyle.Alignment = oldCell.CellStyle.Alignment
                        newCell.CellStyle.BorderBottom = oldCell.CellStyle.BorderBottom
                        newCell.CellStyle.BorderLeft = oldCell.CellStyle.BorderLeft
                        newCell.CellStyle.BorderRight = oldCell.CellStyle.BorderRight
                        newCell.CellStyle.BorderTop = oldCell.CellStyle.BorderTop
                        newCell.CellStyle.TopBorderColor = oldCell.CellStyle.TopBorderColor
                        newCell.CellStyle.BottomBorderColor = oldCell.CellStyle.BottomBorderColor
                        newCell.CellStyle.RightBorderColor = oldCell.CellStyle.RightBorderColor
                        newCell.CellStyle.LeftBorderColor = oldCell.CellStyle.LeftBorderColor
                        newCell.CellStyle.FillBackgroundColor = oldCell.CellStyle.FillBackgroundColor
                        newCell.CellStyle.FillForegroundColor = oldCell.CellStyle.FillForegroundColor
                        newCell.CellStyle.DataFormat = oldCell.CellStyle.DataFormat
                        newCell.CellStyle.FillPattern = oldCell.CellStyle.FillPattern
                        newCell.CellStyle.IsHidden = oldCell.CellStyle.IsHidden
                        newCell.CellStyle.Indention = oldCell.CellStyle.Indention
                        newCell.CellStyle.IsLocked = oldCell.CellStyle.IsLocked
                        newCell.CellStyle.Rotation = oldCell.CellStyle.Rotation
                        newCell.CellStyle.VerticalAlignment = oldCell.CellStyle.VerticalAlignment
                        newCell.CellStyle.WrapText = oldCell.CellStyle.WrapText
                        newCell.CellStyle.SetFont(oldCell.CellStyle.GetFont(wk))
                      Next
                    Next
                  Next


                  '取得位置
                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _QR_Index As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"), seat_dic("ConditionalFormattingIndex"))
                    If seat_dic("TYPE") = "QRCODE" Then
                      Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Index_point)
                    Else
                      Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Row_point)
                    End If
                    _locat.Add(_point)
                  Next

                  Dim RowBreakIndex As Integer = 0
                  '填入資料
                  For i = 0 To Allist.Count - 1
                    If i = 0 Then
                      For j = 0 To _locat.Count - 1
                        Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 2 + i + _Tablecount), HSSFRow)
                        If _locat(j).TYPE = "STRING" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                          RowBreakIndex = hr.RowNum
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y - 1 + i + _Tablecount + tmp2
                            'Dim _x = _locat(j).X - 1
                            'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          '

                        ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                          'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                          Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                          If _VALUE IsNot Nothing Then
                            Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                            Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                            Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                            anchor.AnchorType = 0
                            Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                          End If
                        Else '-視為integer
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                        End If
                      Next

                    Else
                      '不是第一組就先插一條分頁線
                      hst.SetRowBreak(RowBreakIndex)

                      For j = 0 To _locat.Count - 1
                        Dim hr = DirectCast(hst.GetRow(_locat(j).Y - 1 + (i * Interval)), HSSFRow)
                        If _locat(j).TYPE = "STRING" Then
                          hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                          RowBreakIndex = hr.RowNum
                          'benny test
                          If _locat(j).ConditionalFormattingIndex <> "" Then '-代表需要複製格式化
                            Dim _ConditionalFormattingIndex = _locat(j).ConditionalFormattingIndex
                            Dim _test = hst.SheetConditionalFormatting
                            Dim _rule = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetRule(0)
                            Dim _font = _rule.CreateFontFormatting
                            '_font.SetFontStyle(False, True)
                            'Dim _font = _test.GetConditionalFormattingAt(_ConditionalFormattingIndex).GetFormattingRanges

                            Dim _x = _locat(j).X - 1
                            Dim _Y = _locat(j).Y + (i * Interval)
                            'Dim _x = _locat(j).X - 1
                            'Dim _Y = _locat(j).Y - 2 + i + _Tablecount
                            Dim _result = ""
                            If TurnIntegerToX(_x, _result, msgResult_Message) = False Then Return False
                            Dim _value = _result & _Y & ":" & _result & _Y
                            Dim regions() As CellRangeAddress = {CellRangeAddress.ValueOf(_value)} '-更改
                            hst.SheetConditionalFormatting.AddConditionalFormatting(regions, _rule)
                          End If
                          '

                        ElseIf _locat(j).TYPE = "QRCODE" Or _locat(j).TYPE = "DATAMATRIX" Then
                          'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                          Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                          If _VALUE IsNot Nothing Then
                            Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                            Dim patriarch As HSSFPatriarch = hst.CreateDrawingPatriarch()
                            Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1 + (i * Interval), _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y) + (i * Interval))
                            anchor.AnchorType = 0
                            Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                          End If
                        Else
                          hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                        End If
                      Next
                    End If
                  Next
                  '移除複製出來的
                  Dim clear = False
                  While clear = False
                    For i = 0 To wk.Count - 1
                      If wk(i).SheetName.Contains("Sample") Then
                        wk.RemoveSheetAt(i)
                        Exit For
                      End If
                      If i = wk.Count - 1 Then
                        clear = True
                      End If
                    Next
                  End While
                  wk.SetPrintArea(0, 0, ColumnRange, 0, RowBreakIndex)
                ElseIf _info("TABLE").Contains("CopySheet") Then '-特殊列印
                  Dim _RowBreak As Integer = 0  '-儲存分頁線 共有幾個分頁線
                  Dim Interval = _info("FormInterval")  '-複製間距

                  '取得位置
                  '---取得塞Table的位置 new         
                  Dim _locat As New List(Of clsExcelSet)
                  Dim _QR_Index As New List(Of clsExcelSet)
                  Dim _dicseatitem As Dictionary(Of String, String) = ClsConfig.ReadKEYDictionary(_info("TABLE"))
                  For Each _item In _dicseatitem
                    Dim seat_dic = ClsConfig.ReadStringValueDictionary(_info("TABLE"), _item.Key)
                    Dim _point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                    If seat_dic("TYPE") = "QRCODE" OrElse seat_dic("TYPE") = "DATAMATRIX" Then
                      Dim _QR_Index_point As clsExcelSet = New clsExcelSet(seat_dic("INDEX_X"), seat_dic("INDEX_Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Index_point)
                    Else
                      Dim _QR_Row_point As clsExcelSet = New clsExcelSet(seat_dic("X"), seat_dic("Y"), seat_dic("TYPE"))
                      _QR_Index.Add(_QR_Row_point)
                    End If
                    _locat.Add(_point)
                  Next

                  ''複製Sheet
                  Dim lstSheet As New List(Of HSSFSheet)
                  For i = 0 To Allist.Count - 1  '---Table 生成      
                    Dim hst_CopySheet = hst.CopySheet("Sheet_" & i.ToString) '---複製一個新的Sheet
                    lstSheet.Add(hst_CopySheet)
                  Next

                  For i = 0 To lstSheet.Count - 1 '往sheet裡填資料
                    '填入資料
                    For j = 0 To _locat.Count - 1
                      If Allist(i).Count = j Then Exit For
                      Dim hr = DirectCast(lstSheet(i).GetRow(_locat(j).Y - 2 + _Tablecount), HSSFRow)
                      If _locat(j).TYPE = "STRING" Then
                        hr.Cells(_locat(j).X - 1).SetCellValue(Allist(i)(j))
                      ElseIf _locat(j).TYPE = "QRCODE" Then
                        'Dim _VALUE = CodeEncoderFromString(_list(_count)) '-要塞進去QRCode的值
                        Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                        If _VALUE IsNot Nothing Then
                          Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                          Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                          Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                          anchor.AnchorType = 0
                          Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                        End If
                      ElseIf _locat(j).TYPE = "DATAMATRIX" Then
                        Dim _VALUE = CodeEncoderFromString(Allist(i)(j), _locat(j).TYPE) '-要塞進去QRCode的值
                        If _VALUE IsNot Nothing Then
                          Dim pictureIndex = wk.AddPicture(BmpToBytes(_VALUE), NPOI.SS.UserModel.PictureType.JPEG)
                          Dim patriarch As HSSFPatriarch = lstSheet(i).CreateDrawingPatriarch()
                          Dim anchor As HSSFClientAnchor = New HSSFClientAnchor(0, 0, 0, 0, _locat(j).X - 1, _locat(j).Y - 1, _locat(j).X - 1 + Integer.Parse(_QR_Index(j).X), _locat(j).Y - 1 + Integer.Parse(_QR_Index(j).Y))
                          anchor.AnchorType = 0
                          Dim picture As HSSFPicture = patriarch.CreatePicture(anchor, pictureIndex)
                        End If
                      Else
                        hr.Cells(_locat(j).X - 1).SetCellValue(Convert.ToInt32(Allist(i)(j)))
                      End If
                    Next
                  Next
                  '移除複製出來的

                  Dim clear = False
                  While clear = False
                    For i = 0 To wk.Count - 1
                      If wk(i).SheetName.Contains("Sheet_") = False Then
                        wk.RemoveSheetAt(i)
                        Exit For
                      End If
                      If i = wk.Count - 1 Then
                        clear = True
                      End If
                    Next
                  End While
                End If
              End While
            Next
            If hst.RowBreaks.Any AndAlso blnPageTop Then
              Dim hr = DirectCast(hst.GetRow(BaseRowFrom), HSSFRow)
              For HeadIndex = hst.RowBreaks.Where(Function(g) g > BaseRowFrom).Count - 1 To 0 Step -1
                Dim i = hst.RowBreaks.Where(Function(g) g > BaseRowFrom).ToList
                hr.CopyRowTo(i(HeadIndex) + 1)
                '插完頁首後,要把這頁首以下的高度全部重設
                For a = hst.LastRowNum To i(HeadIndex) Step -1
                  hst.GetRow(a).Height = hst.GetRow(a - 1).Height
                Next
                _Tablecount = _Tablecount + 1
GoNext:
              Next
            End If

            BaseRowFrom = BaseRowCount + _Tablecount + 1
          Next
#End Region
        End If


      End Using
      '3.寫回檔案
      Dim file As New FileStream(sExcelPath, FileMode.Create)
      '產生檔案
      wk.Write(file)
      file.Close()

      Return True
    Catch ex As Exception
      msgResult_Message = ex.Message
      Return False
    End Try
  End Function
  Private Function TurnIntegerToX(ByRef _value As String, ByRef _result As String, ByRef msgResult_Message As String) As Boolean
    Try
      '-_value 轉成X軸
      _value = _value + 1
      Select Case _value
        Case "1"
          _result = "A"
        Case "2"
          _result = "B"
        Case "3"
          _result = "C"
        Case "4"
          _result = "D"
        Case "5"
          _result = "E"
        Case "6"
          _result = "F"
        Case "7"
          _result = "G"
        Case "8"
          _result = "H"
        Case "9"
          _result = "I"
        Case "10"
          _result = "J"
        Case "11"
          _result = "K"
        Case "12"
          _result = "L"
        Case "13"
          _result = "M"
        Case "14"
          _result = "N"
        Case "15"
          _result = "O"
        Case "16"
          _result = "P"
        Case "17"
          _result = "Q"
        Case "18"
          _result = "R"
        Case "19"
          _result = "S"
        Case "20"
          _result = "T"
        Case "21"
          _result = "U"
        Case "22"
          _result = "V"
        Case "23"
          _result = "W"
        Case "24"
          _result = "X"
        Case "25"
          _result = "Y"
        Case "26"
          _result = "Z"

        Case Else
          msgResult_Message = "TurnIntegerToX ERROR,X:" & _value
          Return False
      End Select

      Return True
    Catch ex As Exception
      msgResult_Message = ex.Message
      Return False
    End Try
  End Function
  Private Function printExcel(ByVal sFilePath As String, ByVal evFilePath As String) As Boolean
    'For demo
    'If DisablePrinterFlag = 1 Then Return
    ' 1. 判斷檔案是否存在
    Dim fi As New FileInfo(sFilePath)
    If fi.Exists = False Then
      Return False
    End If
    'For Each Proc In Process.GetProcessesByName("Excel")
    '  Proc.Kill()
    'Next
    '資料夾存在
    'If Directory.Exists(sFolderPath) Then
    '  If CleanFolder(sFolderPath) = False Then 'YES -CLEAN
    '    Return
    '  End If
    'Else
    '  Directory.CreateDirectory(sFolderPath) 'NO-新增資料夾
    'End If

    ''2.將目標excel 拆到指定folder 並依sheet 拆成不同的excel
    'If FolderInfo(sFilePath, sFolderPath) = False Then Return


    ' 3. 要列印 EXCEL 檔案的程式位置
    '    Excel Viewer 在 C:\Program Files\Microsoft Office\Office12\XLVIEW.exe
    'Dim evFilePath As String = "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE"


    ' 4. 初始化 DdeClint 類別物件 ddeClient
    '    DdeClint(Server 名稱,string topic 名稱)
    Dim ddeClient As New DdeClient("excel", "system")


    Dim process__1 As Process = Nothing

    Do
      Try
        ' 5. DDE Client 進行連結
        If ddeClient.IsConnected = False Then

          ddeClient.Connect()

        End If

      Catch generatedExceptionName As DdeException

        ' 6. 開啟 Excel Viewer
        Dim info As New ProcessStartInfo(evFilePath)

        info.WindowStyle = ProcessWindowStyle.Minimized

        info.UseShellExecute = True

        'Excel Viewer used --- 
        info.Arguments = sFilePath
        info.Arguments = String.Format("""{0}""", sFilePath)
        '---

        process__1 = Process.Start(info)


        process__1.WaitForInputIdle()

      End Try
    Loop While ddeClient.IsConnected = False AndAlso process__1.HasExited = False

    ' 7. DDE 處理
    Try


      'For Each fname As String In System.IO.Directory.GetFileSystemEntries(sFolderPath)
      ddeClient.Execute(String.Format("[Open(""{0}"")]", sFilePath), 60000)

      ' 開啟 EXCEL 檔案           
      ddeClient.Execute("[Print()]", 60000)
      'ddeClient.Execute("[Printview()]", 60000)

      ' 列印 EXCEL 檔案
      ddeClient.Execute("[Close()]", 60000)
      'Next

      'process__1.Kill()

      Return True

    Catch ex As Exception
      Return False
      ' process__1.Kill()
      ' MessageBox.Show(ex.Message)
    End Try

  End Function

  Public Function PrintSetPrinter(ByRef strmessage As String, ByVal printer As String, ByVal IP As String) As Boolean
    Try
      '-先確認是否有此標籤機,如果沒有就用預設的
      Shell(String.Format("rundll32 printui.dll,PrintUIEntry /y /n ""{0}""", printer))
      SendMessageToLog(String.Format("Set Default Printer to {0}", printer), eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
      Return True
    Catch ex As Exception
      strmessage = ex.Message
      SendMessageToLog(strmessage, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function
  '列印PDF, 因無法預測多久會印完，因此限制四分鐘沒印完時直接取消，關閉PDF
  Public Function PrintPDF(ByVal ret_reportfilename As String, ByRef strmessage As String) As Boolean
		Try
			'ret_reportfilename = "D:\Log\WMSLite\Export\201109\BoxBarcode_1_013791-1.pdf"
			Dim processInfo As ProcessStartInfo = New ProcessStartInfo(ret_reportfilename)
			processInfo.Verb = "print"
			processInfo.WindowStyle = ProcessWindowStyle.Hidden
			processInfo.CreateNoWindow = True
			'processInfo.UseShellExecute = False
			Dim process As Process = New Process()
			process.StartInfo = processInfo

			process.Start()

			Dim timeOut = DateTime.Now.AddSeconds(240)
			Dim printing = False '是否開始列印
			Dim done = False '是否列印完成
			'Dim done = False '是否列印完成
			'取純檔名部分，跟PrintQueue進行比對
			Dim pureFileName = Path.GetFileName(ret_reportfilename)
			'限定最大等待時間
			'While DateTime.Now.CompareTo(timeOut) < 0
			'  If printing = False Then
			'    '未開始列印前發現檔名相同的列印工作
			'    If CheckPrintQueue(pureFileName) Then
			'      printing = True
			'      SendMessageToLog(pureFileName & "列印中...", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			'    End If
			'  Else
			'    '已開始列印後，同檔名列印工作消失表示列印完成
			'    If CheckPrintQueue(pureFileName) = False Then
			'      done = True
			'      SendMessageToLog(pureFileName & "列印完成!!", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			'      Exit While

			'    End If
			'  End If
			'  System.Threading.Thread.Sleep(100)
			'End While
			While DateTime.Now.CompareTo(timeOut) < 0
				If printing = False Then
					'未開始列印前發現檔名相同的列印工作
					If CheckPrintQueue(pureFileName) Then
						printing = True
						SendMessageToLog(pureFileName & "列印中...", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
						Exit While
					End If
				End If
				System.Threading.Thread.Sleep(100)
			End While
			Try
				'若程序尚未關閉，強制關閉之
				If process.CloseMainWindow = False Then
					process.Kill()
				End If
			Catch ex As Exception

			End Try

			'If done = False Then
			'  SendMessageToLog("無法確認報表" & pureFileName & "列印狀態!!", eCALogTool.ILogTool.enuTrcLevel.lvTRACE)
			'End If

			'process.WaitForInputIdle()
			'process.Kill()
			'System.Threading.Thread.Sleep(30000)
			'If Not process.CloseMainWindow() Then

			'End If
			Return True
			'Console.WriteLine("Done")
		Catch ex As Exception
			strmessage = ex.Message
      SendMessageToLog(strmessage, eCALogTool.ILogTool.enuTrcLevel.lvError)
      Return False
    End Try
  End Function

  '需查詢 WMI 記得加入參照及 System.Management
  Public Function CheckPrintQueue(ByVal file As String) As Boolean

    '尋找PrintQueue有沒有檔案相同的列印工作
    Dim searchQuery = "SELECT * FROM Win32_PrintJob"
    Dim printJobs As ManagementObjectSearcher = New ManagementObjectSearcher(searchQuery)
    Dim jobs As ManagementObjectCollection = printJobs.Get()

    'Return True
    For Each obj As ManagementObject In jobs
      If obj.Properties("Document").Value.ToString() = file Then
        Return True
      End If
    Next
    Return False

    'Return printJobs.Any(o >= (CSharpImpl.__Assign(o.Properties("Document").Value, file)))
  End Function

End Module
