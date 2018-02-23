Imports System.Xml
Public Class AFDInterface
  Private mvAddress As String
  Private mvTown As String
  Private mvCounty As String
  Private mvPostCode As String
  Private mvOrgName As String

  Public Sub New()
    mvAddress = String.Empty
    mvTown = String.Empty
    mvCounty = String.Empty
    mvPostCode = String.Empty
    mvOrgName = String.Empty
  End Sub

  Public Function DoSearch(ByVal pPostcode As String, Optional ByVal pAddress As String = "") As ListItemCollection
    Dim vXMLDoc As System.Xml.XmlDocument
    Dim vRoot As System.Xml.XmlElement
    Dim vItemNodes As System.Xml.XmlNodeList
    Dim vResults As New ListItemCollection
    Dim vIndex As Integer

    vXMLDoc = CType(AFDGetXMLDoc(False, pPostcode, "", pAddress), System.Xml.XmlDocument)
    vRoot = vXMLDoc.DocumentElement
    vItemNodes = vRoot.SelectNodes("Item")
    SetAFDAddress(vItemNodes(0))
    For vIndex = 0 To vItemNodes.Count - 1
      SetAFDAddress(vItemNodes(vIndex))
      vResults.Add(Replace(mvAddress, vbCrLf, ", "))
    Next
    Return vResults
  End Function

  Private Sub SetAFDAddress(ByRef pNode As System.Xml.XmlNode)
    Dim vDataNode As System.Xml.XmlNode
    Dim vProperty As String = ""
    Dim vStreet As String = ""
    Dim vLocality As String = ""

    vDataNode = pNode.SelectSingleNode("Organisation")
    If Not (vDataNode Is Nothing) Then mvOrgName = vDataNode.InnerText
    vDataNode = pNode.SelectSingleNode("Property")
    If Not (vDataNode Is Nothing) Then vProperty = vDataNode.InnerText
    vDataNode = pNode.SelectSingleNode("Street")
    If Not (vDataNode Is Nothing) Then vStreet = vDataNode.InnerText
    vDataNode = pNode.SelectSingleNode("Locality")
    If Not (vDataNode Is Nothing) Then vLocality = vDataNode.InnerText
    mvAddress = Replace(vProperty, ", ", vbCrLf)
    If Len(mvAddress) > 0 Then
      mvAddress = mvAddress & vbCrLf & vStreet
    Else
      mvAddress = vStreet
    End If
    If vLocality.Length > 0 Then mvAddress = mvAddress & vbCrLf & vLocality
    vDataNode = pNode.SelectSingleNode("Town")
    If Not (vDataNode Is Nothing) Then mvTown = UCase(vDataNode.InnerText)
    vDataNode = pNode.SelectSingleNode("PostalCounty")
    If Not (vDataNode Is Nothing) Then mvCounty = vDataNode.InnerText
    vDataNode = pNode.SelectSingleNode("Postcode")
    If Not (vDataNode Is Nothing) Then mvPostCode = UCase(vDataNode.InnerText)
  End Sub

  Private Function AFDGetXMLDoc(ByRef pList As Boolean, ByRef pPostCode As String, ByRef pKey As String, Optional ByRef pAddress As String = "", Optional ByRef pTown As String = "") As System.Xml.XmlDocument
    Dim vXMLDoc As System.Xml.XmlDocument
    Dim vXmlParams As String
    Dim vRoot As System.Xml.XmlElement
    Dim vDataNode As System.Xml.XmlNode
    Dim vItemNodes As System.Xml.XmlNodeList
    Dim vFields As String

    ' Initialise the Microsoft XML Document Object Model
    vXMLDoc = New System.Xml.XmlDocument
    If pList Then
      vFields = "List"
    Else
      vFields = "Standard"
    End If
    ' Build up the XML query string
    vXmlParams = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.afd_everywhere_server) & "/afddata.pce?"
    vXmlParams = vXmlParams & "Serial=" & "" & "&"
    vXmlParams = vXmlParams & "Password=" & "" & "&"
    vXmlParams = vXmlParams & "UserID=" & "" & "&"
    If Len(pPostCode) > 0 Then
      vXmlParams = vXmlParams & "Data=Address&Task=FastFind&Fields=" & vFields
      vXmlParams = vXmlParams & "&Lookup=" & pPostCode
    ElseIf Len(pKey) > 0 Then
      vXmlParams = vXmlParams & "Data=Address&Task=Retrieve&Fields=" & vFields
      vXmlParams = vXmlParams & "&Key=" & pKey
    Else
      vXmlParams = vXmlParams & "Data=Address&Task=FastFind&Fields=" & vFields
      vXmlParams = vXmlParams & "&Lookup=" & Replace(pAddress, vbCrLf, ",") & ", " & pTown
    End If
    ' Set the maximum number of records to return
    vXmlParams = vXmlParams & "&MaxQuantity=100"
    ' Load the XML from the webserver with the query string
    Try
      vXMLDoc.Load(vXmlParams)
    Catch vEx As System.Xml.XmlException
      RaiseError(DataAccessErrors.daeUniservError, "Error: " & vEx.Message)
    End Try
    ' Check if PCE returned an error and if the document is valid
    vRoot = vXMLDoc.DocumentElement
    vDataNode = vRoot.SelectSingleNode("Result")
    vItemNodes = vRoot.SelectNodes("Item")
    If vDataNode Is Nothing Or vItemNodes Is Nothing Then
      RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
    End If
    If Val(vDataNode.InnerText) < 1 Then
      vDataNode = vRoot.SelectSingleNode("ErrorText")
      If vDataNode Is Nothing Then
        RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
      Else
        RaiseError(DataAccessErrors.daeUniservError, vDataNode.InnerText) ' Show the user the error
      End If
    End If
    Return vXMLDoc
  End Function
  Public ReadOnly Property Town() As String
    Get
      Return mvTown
    End Get
  End Property
  Public ReadOnly Property County() As String
    Get
      Return mvCounty
    End Get
  End Property
  Public ReadOnly Property Postcode() As String
    Get
      Return mvPostcode
    End Get
  End Property
  Public ReadOnly Property Address() As String
    Get
      Return mvAddress
    End Get
  End Property
End Class
