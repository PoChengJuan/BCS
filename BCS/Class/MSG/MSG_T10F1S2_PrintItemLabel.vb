Imports System.Xml.Serialization
<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F1S2_PrintItemLabel
  Public Property Header As New clsHeader
  Public Property Body As New clsBody
  Public Property KeepData As String

  Public Class clsBody
    <XmlElement(ElementName:="PrintInfo")>
    Public Property PrintInfo As New clsPrintInfo
    Public Class clsPrintInfo
      Public Property PRINT_ENABLE As String
      Public Property TO_PDF As String
      Public Property PRINT_TYPE As String
    End Class
    <XmlElement(ElementName:="LabelList")>
    Public Property LabelList As New clsLabelList
    Public Class clsLabelList
      <XmlElement(ElementName:="LabelInfo")>
      Public Property LabelInfo As New List(Of clsLabelInfo)
      Public Class clsLabelInfo
        Public Property ITEM_LABEL_ID As String
        Public Property TAG1 As String
        Public Property TAG2 As String
        Public Property TAG3 As String
        Public Property TAG4 As String
        Public Property TAG5 As String
        Public Property TAG6 As String
        Public Property TAG7 As String
        Public Property TAG8 As String
        Public Property TAG9 As String
        Public Property TAG10 As String
        Public Property TAG11 As String
        Public Property TAG12 As String
        Public Property TAG13 As String
        Public Property TAG14 As String
        Public Property TAG15 As String
        Public Property TAG16 As String
        Public Property TAG17 As String
        Public Property TAG18 As String
        Public Property TAG19 As String
        Public Property TAG20 As String
        Public Property TAG21 As String
        Public Property TAG22 As String
        Public Property TAG23 As String
        Public Property TAG24 As String
        Public Property TAG25 As String
        Public Property TAG26 As String
        Public Property TAG27 As String
        Public Property TAG28 As String
        Public Property TAG29 As String
        Public Property TAG30 As String
        Public Property TAG31 As String
        Public Property TAG32 As String
        Public Property TAG33 As String
        Public Property TAG34 As String
        Public Property TAG35 As String
      End Class
    End Class
  End Class
End Class