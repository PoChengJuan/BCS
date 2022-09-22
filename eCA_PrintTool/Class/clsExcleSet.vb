Public Class clsExcelSet
  Public X As String
  Public Y As String
	Public TYPE As String
	Public ConditionalFormattingIndex As String
	Sub New(ByVal _x As String, ByVal _y As String, ByVal _type As String, Optional _ConditionalFormattingIndex As String = "")
		X = _x
		Y = _y
		TYPE = _type
		ConditionalFormattingIndex = _ConditionalFormattingIndex
	End Sub
End Class