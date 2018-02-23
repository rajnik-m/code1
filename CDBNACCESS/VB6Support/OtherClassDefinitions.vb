Namespace Access

  Public Class CDBCollection
    Inherits CollectionBase

    Private mvHashTable As Hashtable
    Private mvOffset As Integer = 1

    Public Sub New()
      MyBase.New()
      mvHashTable = New Hashtable
    End Sub

    Public Overloads Sub Clear()
      List.Clear()
      mvHashTable = New Hashtable
    End Sub

    Public Sub Add(ByVal pItem As Object, ByVal pKey As String)
      List.Add(pItem)
      mvHashTable.Add(pKey, pItem)
    End Sub

    Public Sub Add(ByVal pItem As Object)
      List.Add(pItem)
    End Sub

    Public Function Exists(ByVal pKey As String) As Boolean
      If mvHashTable.ContainsKey(pKey) Then Return True
    End Function

    Default Public ReadOnly Property Item(ByVal pIndex As Integer) As Object
      Get
        pIndex -= mvOffset
        If pIndex > Count - 1 Or pIndex < 0 Then
          Throw New ArgumentOutOfRangeException
        Else
          Return List.Item(pIndex)
        End If
      End Get
    End Property

    Default Public ReadOnly Property Item(ByVal pIndex As String) As Object
      Get
        If mvHashTable.ContainsKey(pIndex) Then
          Return mvHashTable.Item(pIndex)
        Else
          Throw New ArgumentOutOfRangeException
        End If
      End Get
    End Property

    Public Sub Remove(ByVal pKey As String)
      If mvHashTable.ContainsKey(pKey) Then
        Dim vItem As Object = mvHashTable.Item(pKey)
        mvHashTable.Remove(pKey)
        List.Remove(vItem)
      Else
        Throw New ArgumentException("Key not found in collection") 'TODO Resource
      End If
    End Sub

  End Class

End Namespace


