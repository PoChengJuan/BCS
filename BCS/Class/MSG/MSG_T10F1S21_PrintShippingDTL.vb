Imports System.Xml.Serialization
<XmlRoot(ElementName:="Message")>
Public Class MSG_T10F1S21_PrintShippingDTL
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
    <XmlElement(ElementName:="OrderList")>
    Public Property OrderList As New clsOrderList
    Public Class clsOrderList
      <XmlElement(ElementName:="OrderInfo")>
      Public Property OrderInfo As New List(Of clsOrderInfo)
      Public Class clsOrderInfo
        Public Property PO_ID As String
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
        <XmlElement(ElementName:="OrderDTLList")>
        Public Property OrderDTLList As New clsOrderDTLList
        Public Class clsOrderDTLList
          <XmlElement(ElementName:="OrderDTLInfo")>
          Public Property OrderDTLInfo As List(Of clsOrderDTLInfo)
          Public Class clsOrderDTLInfo
            Public Property PO_SERIAL_NO As String
            Public Property SKU_NO As String
            Public Property QTY As String
            Public Property D_TAG1 As String
            Public Property D_TAG2 As String
            Public Property D_TAG3 As String
            Public Property D_TAG4 As String
            Public Property D_TAG5 As String
            Public Property D_TAG6 As String
            Public Property D_TAG7 As String
            Public Property D_TAG8 As String
            Public Property D_TAG9 As String
            Public Property D_TAG10 As String
            Public Property D_TAG11 As String
            Public Property D_TAG12 As String
            Public Property D_TAG13 As String
            Public Property D_TAG14 As String
            Public Property D_TAG15 As String
            Public Property D_TAG16 As String
            Public Property D_TAG17 As String
            Public Property D_TAG18 As String
            Public Property D_TAG19 As String
            Public Property D_TAG20 As String
          End Class
        End Class
      End Class
    End Class
  End Class
End Class
