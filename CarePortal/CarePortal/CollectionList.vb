Public Class CollectionList(Of ItemType)
  Inherits CollectionBase
  Private mvHashTable As Hashtable
  Private mvOffset As Integer

  Public Sub New()
    MyBase.New()
    mvHashTable = New Hashtable
  End Sub

  Public Sub New(ByVal pOffset As Integer)
    MyBase.New()
    mvHashTable = New Hashtable
    mvOffset = pOffset
  End Sub

  Public Overloads Sub Clear()
    MyBase.Clear()
    mvHashTable = New Hashtable
  End Sub

  Public Sub Add(ByVal pKey As String, ByVal pItem As ItemType)
    List.Add(pItem)
    mvHashTable.Add(pKey, pItem)
  End Sub

  Public Sub Insert(ByVal pIndex As Integer, ByVal pKey As String, ByVal pItem As ItemType)
    pIndex -= mvOffset
    List.Insert(pIndex, pItem)
    mvHashTable.Add(pKey, pItem)
  End Sub

  Public Function IndexOf(ByVal pItem As ItemType) As Integer
    Return List.IndexOf(pItem) + mvOffset
  End Function

  Public Sub Remove(ByVal pKey As String)
    If Me.ContainsKey(pKey) Then
      Dim vItem As ItemType = Me.Item(pKey)
      mvHashTable.Remove(pKey)
      List.Remove(vItem)
    Else
      Throw New ArgumentException("Key not found in collection") 'TODO Resource
    End If
  End Sub

  Public Sub Remove(ByVal pIndex As Integer)
    pIndex -= mvOffset
    If pIndex > Count - 1 Or pIndex < 0 Then
      Throw New ArgumentOutOfRangeException
    Else
      Dim vItem As ItemType = Me.Item(pIndex + mvOffset)    'Restore the offset as Me.Item will take it off again
      List.RemoveAt(pIndex)
      RemoveFromHashTable(vItem)
    End If
  End Sub

  Public Sub Remove(ByVal pItem As ItemType)
    List.Remove(pItem)
    RemoveFromHashTable(pItem)
  End Sub

  Private Sub RemoveFromHashTable(ByVal pItem As ItemType)
    For Each vItem As DictionaryEntry In mvHashTable
      If vItem.Value.Equals(pItem) Then
        mvHashTable.Remove(vItem.Key)
        Return
      End If
    Next
  End Sub

  Protected Sub ChangeKey(ByVal pItem As ItemType, ByVal pOldKey As String, ByVal pNewKey As String)
    mvHashTable.Remove(pOldKey)
    mvHashTable.Add(pNewKey, pItem)
  End Sub

  Public Function FindKey(ByVal pItem As ItemType) As String
    Dim vKey As String = ""
    For Each vItem As DictionaryEntry In mvHashTable
      If vItem.Value.Equals(pItem) Then
        vKey = vItem.Key.ToString
        Exit For
      End If
    Next
    Return vKey
  End Function

  Public Function ContainsKey(ByVal pKey As String) As Boolean
    Return mvHashTable.ContainsKey(pKey)
  End Function

  Default Public Overridable ReadOnly Property Item(ByVal pKey As String) As ItemType
    Get
      If mvHashTable.ContainsKey(pKey) Then
        Return (CType(mvHashTable.Item(pKey), ItemType))
      Else
        Throw New ArgumentException(String.Format("Key {0} not found in collection", pKey)) 'TODO Resource
      End If
    End Get
  End Property

  Default Public ReadOnly Property Item(ByVal pIndex As Integer) As ItemType
    Get
      pIndex -= mvOffset
      If pIndex > Count - 1 Or pIndex < 0 Then
        Throw New ArgumentOutOfRangeException
      Else
        Return (CType(List.Item(pIndex), ItemType))
      End If
    End Get
  End Property

  Public ReadOnly Property ItemKey(ByVal pIndex As Integer) As String
    Get
      pIndex -= mvOffset
      If pIndex > Count - 1 Or pIndex < 0 Then
        Throw New ArgumentOutOfRangeException
      Else
        Return FindKey(CType(List.Item(pIndex), ItemType))
      End If
    End Get
  End Property
End Class