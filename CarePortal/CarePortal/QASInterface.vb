Imports CarePortal.QASProWeb

Public Class QASInterface

  Dim mvQAS As CarePortal.QASProWeb.ProWeb

  Private mvAddress As New StringBuilder
  Private mvTown As String
  Private mvCounty As String
  Private mvPostcode As String

  Public Sub New()
    mvQAS = New QASProWeb.ProWeb
    mvQAS.Url = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.qas_pro_web_url)
    Dim vCanSearch As New QACanSearch
    vCanSearch.Country = "GBR"
    vCanSearch.Engine = GetEngine()
    vCanSearch.Layout = GetLayout()
    Dim vQAOK As QASearchOk = mvQAS.DoCanSearch(vCanSearch)
    If Not vQAOK.IsOk Then Throw New Exception(vQAOK.ErrorMessage)
  End Sub

  Public Function DoSearch(ByVal pPostcode As String) As ListItemCollection
    Dim vResults As New ListItemCollection
    Dim vQASearch As New QASearch
    vQASearch.Country = "GBR"
    vQASearch.Engine = GetEngine()
    vQASearch.Layout = GetLayout()
    vQASearch.Search = pPostcode
    Dim vQASearchResult As QASearchResult = mvQAS.DoSearch(vQASearch)
    If vQASearchResult.QAPicklist IsNot Nothing Then
      If vQASearchResult.QAPicklist.PicklistEntry.Length > 0 Then
        For Each vItem As PicklistEntryType In vQASearchResult.QAPicklist.PicklistEntry
          vResults.Add(New ListItem(vItem.Picklist, vItem.Moniker))
        Next
      End If
    End If
    Return vResults
  End Function

  Public Function GetAddress(ByVal pID As String) As String
    mvTown = ""
    mvCounty = ""
    mvPostcode = ""
    Dim vGetAddress As New QAGetAddress
    vGetAddress.Moniker = pID
    vGetAddress.Layout = GetLayout()
    Dim vAddress As Address = mvQAS.DoGetAddress(vGetAddress)
    If vAddress IsNot Nothing Then
      mvAddress.AppendLine(vAddress.QAAddress.AddressLine(1).Line)
      If vAddress.QAAddress.AddressLine(2).Line.Length > 0 Then
        mvAddress.AppendLine(vAddress.QAAddress.AddressLine(2).Line)
        If vAddress.QAAddress.AddressLine(3).Line.Length > 0 Then
          mvAddress.AppendLine(vAddress.QAAddress.AddressLine(3).Line)
          If vAddress.QAAddress.AddressLine(4).Line.Length > 0 Then mvAddress.AppendLine(vAddress.QAAddress.AddressLine(4).Line)
        End If
      End If
      mvTown = vAddress.QAAddress.AddressLine(5).Line
      mvCounty = vAddress.QAAddress.AddressLine(6).Line
      mvPostcode = vAddress.QAAddress.AddressLine(7).Line
      Return mvAddress.ToString
    Else
      Return ""
    End If
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

  Private Function GetEngine() As EngineType
    Dim vEngineType As New QASProWeb.EngineType
    vEngineType.Flatten = True
    vEngineType.FlattenSpecified = True
    vEngineType.Value = EngineEnumType.Singleline
    Return vEngineType
  End Function

  Private Function GetLayout() As String
    Return "CDBDefault"
  End Function

End Class
