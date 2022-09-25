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


Module ModuleQRCode

  ''' <summary>
  ''' 將文字轉成QRCode Bitmap
  ''' </summary>
  ''' <param name="_InputDataString">欲轉換的文字</param>
  ''' <returns></returns>
  Public Function CodeEncoderFromString(ByVal _InputDataString As String, ByVal EncoderType As String, Optional ByRef gLength As Integer = 0) As Bitmap
    Try
      'Dim ce As New DmtxImageEncoderOptions
      'Dim encoder = New DmtxImageEncoder
      Select Case EncoderType
        Case "QRCODE"
          Dim ce As New QRCodeEncoder()
          ce.QRCodeScale = 6
          ce.QRCodeVersion = 0
          Return ce.Encode(_InputDataString)
        Case "DATAMATRIX"
          Dim ce As New DmtxImageEncoderOptions
          Dim encoder = New DmtxImageEncoder
          ce.ModuleSize = 2
          ce.MarginSize = 0
          Return encoder.EncodeImage(_InputDataString, ce)
          'ce.QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE
        Case "Code128"
          Return Code128Rendering.MakeBarcodeImage(_InputDataString, gLength, False)
        Case "Code39"
          Dim br As New ZXing.BarcodeWriter
          br.Format = BarcodeFormat.CODE_39
          br.Options.Width = gLength
          br.Options.Height = 1
          Return br.Write(_InputDataString)
        Case "CODABAR"
          Dim br As New ZXing.BarcodeWriter
          br.Format = BarcodeFormat.CODABAR
          br.Options.Width = gLength
          br.Options.Height = 1
          Return br.Write(_InputDataString)
        Case Else
          Return Nothing
      End Select
      'ce.QRCodeScale = 6
      'ce.QRCodeVersion = 8
      'ce.QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.M
    Catch ex As Exception

    End Try

  End Function

  ''' <summary>
  ''' 根據QRCode Bitmap解碼成原本文字內容
  ''' </summary>
  ''' <param name="_BitMap">QRCode Bitmap</param>
  ''' <returns></returns>
  Public Function CodeDecoderFromBitMapForQRcode(ByVal _BitMap As Bitmap) As String
    Dim decoder As New QRCodeDecoder()
    Dim decoded_string As String = decoder.decode(New QRCodeBitmapImage(_BitMap))
    Return decoded_string
  End Function

  ''' <summary>
  ''' 根據DataMarix Bitmap解碼成原本文字內容
  ''' </summary>
  ''' <param name="_BitMap">QRCode Bitmap</param>
  ''' <returns></returns>
  Public Function CodeDecoderFromBitMapForDataMarix(ByVal _BitMap As Bitmap) As String
    Dim decoder As New DmtxImageDecoder
    'Dim decoded_string As String = decoder.decode(New QRCodeBitmapImage(_BitMap))
    Dim decoded_string As String = decoder.DecodeImage(_BitMap)(0)
    Return decoded_string
  End Function

  ''' <summary>
  ''' 將轉換後的Bitmap存成jpeg檔案
  ''' </summary>
  ''' <param name="_Bitmap">Bitmap來源</param>
  ''' <param name="_FilePath">檔案路徑(含檔名)</param>
  Public Sub CodeEncoderFromStringWriteToJpeg(ByVal _Bitmap As Bitmap, ByVal _FilePath As String)
		Dim image As Image = _Bitmap
		Dim fs As New FileStream(_FilePath, FileMode.OpenOrCreate, FileAccess.Write)
		image.Save(fs, ImageFormat.Jpeg)
		fs.Close()
		image.Dispose()
	End Sub

  ''' <summary>
  ''' 根據DataMarix Jpeg檔案解碼成原本文字內容
  ''' </summary>
  ''' <param name="_FilePath">QRCode Jpeg檔案位置(含檔名)</param>
  ''' <returns></returns>
  Public Function CodeDecoderFromFilePathForQRcode(ByVal _FilePath As String) As String
    If Not System.IO.File.Exists(_FilePath) Then
      Return Nothing
    End If
    Dim my_bitmap As New Bitmap(Image.FromFile(_FilePath))
    Dim decoder As New DmtxImageDecoder
    Dim decoded_string As String = decoder.DecodeImage(my_bitmap)(0)
    Return decoded_string
  End Function

  ''' <summary>
  ''' 根據QRCode Jpeg檔案解碼成原本文字內容
  ''' </summary>
  ''' <param name="_FilePath">QRCode Jpeg檔案位置(含檔名)</param>
  ''' <returns></returns>
  Public Function CodeDecoderFromFilePath(ByVal _FilePath As String) As String
    If Not System.IO.File.Exists(_FilePath) Then
      Return Nothing
    End If
    Dim my_bitmap As New Bitmap(Image.FromFile(_FilePath))
    Dim decoder As New QRCodeDecoder()
    Dim decoded_string As String = decoder.decode(New QRCodeBitmapImage(my_bitmap))
    Return decoded_string
  End Function

  ''' <summary>
  ''' 將Bitmap轉成BitmapImage(顯示在Image元件上使用，DEMO程式需求)
  ''' </summary>
  ''' <param name="_Bitmap"></param>
  ''' <returns></returns>
  Public Function BitmapToImageSource(ByVal _Bitmap As Bitmap) As BitmapImage
		Dim memory = New MemoryStream()
		_Bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp)
		memory.Position = 0
		Dim bitmap_image = New BitmapImage()
		bitmap_image.BeginInit()
		bitmap_image.StreamSource = memory
		bitmap_image.CacheOption = BitmapCacheOption.OnLoad
		bitmap_image.EndInit()
		memory.Dispose()

		Return bitmap_image
	End Function

	Public Function BmpToBytes(_Bitmap As Bitmap) As Byte()
		Using stream = New MemoryStream()
			_Bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
			Return stream.ToArray()
		End Using
	End Function

End Module
