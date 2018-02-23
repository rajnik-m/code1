Public Class StringList
  Inherits List(Of String)

  Public Sub New(ByVal pString As String)
    MyBase.New()
    BuildList(pString, ",", StringSplitOptions.RemoveEmptyEntries)
  End Sub

  Public Sub New(ByVal pString As String, ByVal pOptions As System.StringSplitOptions)
    MyBase.New()
    BuildList(pString, ",", pOptions)
  End Sub

  Public Sub New(ByVal pString As String, ByVal pSeparator As String)
    MyBase.New()
    BuildList(pString, pSeparator.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
  End Sub

  Private Sub BuildList(ByVal pString As String, ByVal pSeparator As String, ByVal pOptions As System.StringSplitOptions)
    Dim vItems() As String = pString.Split(pSeparator.ToCharArray, pOptions)
    For Each vItem As String In vItems
      Add(vItem)
    Next
  End Sub

  Public Function ContainsAnyItem(ByVal pItems As String) As Boolean
    'This function is a bit like the Exists function.
    'It determines whether the list contains any one of the list of values in pValues.  
    'It does not determine if the collection contains all of the values in pValues.
    Dim vValues() As String = pItems.Split(",".ToCharArray)
    For Each vValue As String In vValues
      If Contains(vValue) Then Return True
    Next
    Return False
  End Function

  Public Function ItemList() As String
    Return ItemList(",")
  End Function

  Public Function ItemList(ByVal pSeparator As String) As String
    Dim vString As New StringBuilder
    Dim vAddSeparator As Boolean
    For Each vItem As String In Me
      If vAddSeparator Then vString.Append(pSeparator)
      vString.Append(vItem)
      vAddSeparator = True
    Next
    Return vString.ToString
  End Function

  Public Function InList() As String
    Dim vString As New StringBuilder
    Dim vAddSeparator As Boolean
    For Each vItem As String In Me
      If vAddSeparator Then vString.Append(",")
      vString.Append("'")
      vString.Append(vItem)
      vString.Append("'")
      vAddSeparator = True
    Next
    Return vString.ToString
  End Function

End Class