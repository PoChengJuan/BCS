Imports System.Xml
Imports System.IO
Public Class clsConfigTool
  Dim _xmlConfig As New XmlDocument
  Private FilePath As String

  Sub New(ByVal _FilePath As String)
    FilePath = _FilePath
  End Sub

  ''' <summary>
  ''' 回傳指定key的值
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定key</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ReadStringValueKey(ByVal strSection As String, ByVal strKey As String) As String
    Dim _rtnMsg As String = String.Empty
    Dim strValue As String = String.Empty
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)


    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then
              strValue = _xn2.Attributes(0).Value.ToString()
              Exit For
            End If
          End If
        Next
        Exit For
      End If
    Next

    Return strValue
  End Function

  ''' <summary>
  ''' 回傳指定區塊內所有KEY值
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <returns>key:Attributes NAME   value:Attributes VALUE </returns>
  ''' <remarks></remarks> 
  Public Function ReadKEYDictionary(ByVal strSection As String) As Dictionary(Of String, String)
    Dim strValue As New Dictionary(Of String, String)
    Dim _rtnMsg As String = String.Empty
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)

    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            strValue.Add(_xn2.Attributes(0).Name, _xn2.Attributes(0).Value)
          End If
        Next
        Exit For
      End If
    Next

    Return strValue
  End Function

  ''' <summary>
  ''' 回傳指定區塊 整個Attributes的 name, value
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定KEY</param>  
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ReadStringValueDictionary(ByVal strSection As String, ByVal strKey As String) As Dictionary(Of String, String)
    Dim strValue As New Dictionary(Of String, String)
    Dim _rtnMsg As String = String.Empty
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)


    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then
              For i As Integer = 0 To _xn2.Attributes.Count - 1
                strValue.Add(_xn2.Attributes(i).Name, _xn2.Attributes(i).Value)
              Next
              Exit For

            End If
          End If
        Next
        Exit For
      End If
    Next

    Return strValue
  End Function

  ''' <summary>
  ''' 回傳指定區塊 指定屬性的值
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定KEY</param>
  ''' <param name="strName">指定屬性</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ReadStringValueDesignation(ByVal strSection As String, ByVal strKey As String, ByVal strName As String) As String
    Dim _rtnMsg As String = String.Empty
    Dim strValue As String = String.Empty
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)

    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then

              For i As Integer = 0 To _xn2.Attributes.Count - 1
                If _xn2.Attributes(i).Name = strName Then strValue = _xn2.Attributes(i).Value.ToString()
              Next
              Exit For

            End If
          End If
        Next
        Exit For
      End If
    Next

    Return strValue

  End Function


  ''' <summary>
  ''' 將LIST內容逐一寫進指定的KEY區域
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定KEY</param>
  ''' <param name="DictionaryValue">寫進Attributes的value</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WriteStringValueDictionary(ByVal strSection As String, ByVal strKey As String, ByVal DictionaryValue As Dictionary(Of String, String)) As Boolean
    '---condition    
    If DictionaryValue.Count = 0 Then Return False

    Dim _rtnMsg As String = String.Empty
    Dim _newNode As XmlElement
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)

    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then

              For i As Integer = 0 To _xn2.Attributes.Count - 1
                _xn2.Attributes(i).Value = DictionaryValue(_xn2.Attributes(i).Name)
              Next

              _xmlConfig.Save(FilePath)
              Return True
            End If
          End If
        Next

        _newNode = _xmlConfig.CreateElement("Key")
        _newNode.SetAttribute(strKey, "0")
        _xn1.AppendChild(_newNode)
        _xmlConfig.Save(FilePath)
        Return True
      End If
    Next

    Dim _newElement As XmlElement = _xmlConfig.CreateElement(strSection)
    _newNode = _xmlConfig.CreateElement("Key")
    _newNode.SetAttribute(strKey, "0")
    _newElement.AppendChild(_newNode)
    _xe.AppendChild(_newElement)
    _xmlConfig.Save(FilePath)
    Return True
  End Function

  ''' <summary>
  ''' 寫進指定區塊 KEY值
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定key</param>
  ''' <param name="strValue">寫進value</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WriteStringValueKey(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String) As Boolean
    Dim _rtnMsg As String = String.Empty
    Dim _newNode As XmlElement
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)

    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then
              _xn2.Attributes(0).Value = strValue
              _xmlConfig.Save(FilePath)
              Return True
            End If
          End If
        Next

        _newNode = _xmlConfig.CreateElement("Key")
        _newNode.SetAttribute(strKey, strValue)
        _xn1.AppendChild(_newNode)
        _xmlConfig.Save(FilePath)
        Return True
      End If
    Next

    Dim _newElement As XmlElement = _xmlConfig.CreateElement(strSection)
    _newNode = _xmlConfig.CreateElement("Key")
    _newNode.SetAttribute(strKey, strValue)
    _newElement.AppendChild(_newNode)
    _xe.AppendChild(_newElement)
    _xmlConfig.Save(FilePath)
    Return True
  End Function


  ''' <summary>
  ''' 寫進指定區塊 指定屬性
  ''' </summary>
  ''' <param name="strSection">指定區塊</param>
  ''' <param name="strKey">指定key</param>
  ''' <param name="strName">指定屬性</param>
  ''' <param name="strValue">寫進value</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WirteStringValueDesignation(ByVal strSection As String, ByVal strKey As String, ByVal strName As String, ByVal strValue As String) As Boolean
    Dim _rtnMsg As String = String.Empty
    Dim _newNode As XmlElement
    Dim _xe As XmlElement

    If Not File.Exists(FilePath) Then
      If Not ConfigCreate(_rtnMsg) Then
        Throw New Exception(_rtnMsg)
      End If
    End If

    _xmlConfig.Load(FilePath)

    _xe = _xmlConfig.DocumentElement

    For Each _xn1 As XmlNode In _xe.ChildNodes
      If _xn1.Name = strSection Then
        For Each _xn2 As XmlNode In _xn1.ChildNodes
          If _xn2.Name = "Key" Then
            If _xn2.Attributes(0).Name = strKey Then
              For i As Integer = 0 To _xn2.Attributes.Count - 1
                If _xn2.Attributes(i).Name = strName Then _xn2.Attributes(i).Value = strValue
              Next

              _xmlConfig.Save(FilePath)
              Return True
            End If
          End If
        Next

        _newNode = _xmlConfig.CreateElement("Key")
        _newNode.SetAttribute(strKey, strValue)
        _xn1.AppendChild(_newNode)
        _xmlConfig.Save(FilePath)
        Return True
      End If
    Next

    Dim _newElement As XmlElement = _xmlConfig.CreateElement(strSection)
    _newNode = _xmlConfig.CreateElement("Key")
    _newNode.SetAttribute(strKey, strValue)
    _newElement.AppendChild(_newNode)
    _xe.AppendChild(_newElement)
    _xmlConfig.Save(FilePath)
    Return True
  End Function

  Private Function ConfigCreate(ByRef _rtnMsg As String) As Boolean
    Try
      If Not File.Exists(FilePath) Then
        Dim _newXmlConfig As New XmlDocument()
        Dim _newXmlDecl As XmlDeclaration = _newXmlConfig.CreateXmlDeclaration("1.0", "", Nothing)
        _newXmlConfig.InsertBefore(_newXmlDecl, _newXmlConfig.DocumentElement)
        Dim _newXmlEle As XmlElement = _newXmlConfig.CreateElement("Setting")
        _newXmlConfig.AppendChild(_newXmlEle)
        _newXmlConfig.Save(FilePath)
        _newXmlConfig = Nothing
      End If
    Catch ex As Exception
      _rtnMsg = ex.Message
      Return False
    End Try
    _rtnMsg = String.Empty
    Return True


  End Function

End Class
