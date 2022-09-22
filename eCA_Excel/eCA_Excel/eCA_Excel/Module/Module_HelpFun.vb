Imports System.IO
Imports Spire.Xls
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Web.ui.page

Public Module Module_HelpFun
  Public Function ExcelToPDF(ByVal FullExcelFileName As String,
                             ByRef ret_FullPDFFileName As String,
                             Optional ByVal ret_Msg As String = "") As Boolean
    Try
      '-取得副檔名
      Dim workbook As New Spire.Xls.Workbook()
      workbook.LoadFromFile(FullExcelFileName)
      Dim sheet As Spire.Xls.Worksheet = workbook.Worksheets(0)
      ret_FullPDFFileName = FullExcelFileName.Replace(Path.GetExtension(FullExcelFileName), ".pdf")
      sheet.SaveToPdf(ret_FullPDFFileName)
      workbook = Nothing
      Return True
    Catch ex As Exception
      ret_Msg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ExcelToPDF(ByVal FullExcelFileName As String,
                           ByVal ret_FullPDFFileName As String,
                           Optional ByVal Width As Long = 0,
                           Optional ByVal Height As Long = 0,
                           Optional ByVal ret_Msg As String = "") As Boolean
    Try
      ''-取的副檔名
      Dim workbook As New Spire.Xls.Workbook()
      'workbook.LoadFromFile(FullExcelFileName)

      workbook.LoadFromStream(New MemoryStream(System.IO.File.ReadAllBytes(FullExcelFileName)))
      'Dim sheet As Worksheet = workbook.Worksheets(0)
      ret_FullPDFFileName = FullExcelFileName.Replace(Path.GetExtension(FullExcelFileName), ".pdf")
      'sheet.SaveToPdf(ret_FullPDFFileName)
      'workbook = Nothing
      'Return True
      Dim pdfDocument As Spire.Pdf.PdfDocument = New Spire.Pdf.PdfDocument()
      pdfDocument.PageSettings.Orientation = Spire.Pdf.PdfPageOrientation.Landscape
      pdfDocument.PageSettings.Width = Width '指定PDF的宽度
      pdfDocument.PageSettings.Height = Height '指定PDF的高度

      Dim SSS As Spire.Xls.Converter.PdfConverterSettings = New Spire.Xls.Converter.PdfConverterSettings
      SSS.TemplateDocument = pdfDocument
      Dim pdfConverter As Spire.Xls.Converter.PdfConverter = New Spire.Xls.Converter.PdfConverter(workbook)
      pdfDocument = pdfConverter.Convert(SSS)
      pdfDocument.SaveToFile(ret_FullPDFFileName, Spire.Pdf.FileFormat.PDF)
      Return True
    Catch ex As Exception
      ret_Msg = ex.ToString()
      Return False
    End Try
  End Function
  Public Function ExcelToPDF2(ByVal FullExcelFileName As String,
                          ByRef ret_FullPDFFileName As String,
                          Optional ByVal ret_Msg As String = "") As Boolean

    Dim xlsxDocument As Excel.Workbook = Nothing
    Dim appExcel = New Microsoft.Office.Interop.Excel.Application()
    Try
      ' Excel 檔案位置
      Dim sourcexlsx = FullExcelFileName
      ' PDF 儲存位置
      Dim targetpdf = FullExcelFileName.Replace(Path.GetExtension(FullExcelFileName), ".pdf")

      '建立 Excel application instance
      'Microsoft.Office.Interop.Excel.Application appExcel = New Microsoft.Office.Interop.Excel.Application();

      '開啟 Excel 檔案
      xlsxDocument = appExcel.Workbooks.Open(sourcexlsx)
        'var xlsxDocument = appExcel.Workbooks.Open(sourcexlsx);

        Dim Type = XlFixedFormatType.xlTypePDF
			Try
				'匯出為 pdf
				xlsxDocument.ExportAsFixedFormat(Type, targetpdf)
				'xlsxDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetpdf);
			Catch ex As Exception
				MsgBox("xlsxDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetpdf) " & vbCrLf & ex.ToString)
      End Try



      '關閉 Excel 檔
      xlsxDocument.Close()
      '結束 Excel
      appExcel.Quit()

      ret_FullPDFFileName = targetpdf
      Return True
    Catch ex As Exception
      ret_Msg = ex.ToString()
      '關閉 Excel 檔
      xlsxDocument.Close()
      '結束 Excel
      appExcel.Quit()
      MsgBox(ret_Msg)
      Return False
    End Try
  End Function


End Module
