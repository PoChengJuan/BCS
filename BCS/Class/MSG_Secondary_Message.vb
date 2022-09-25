Imports System.Xml.Serialization

<XmlRoot(ElementName:="Message")>
Public Class MSG_Secondary_Message
  Public Property Header As clsHeader
  Public Property Body As clsResultInfo
  Public Property KeepData As String

  Public Class clsResultInfo

    Public Property Result As String
    Public Property ResultMessage As String

  End Class
End Class