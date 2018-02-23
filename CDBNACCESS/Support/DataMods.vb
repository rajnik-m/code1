Namespace Access
  Public Class DataMods
    Private mvCol As Collection
    Public Sub Add(ByVal pDataMod As DataMod)
      mvCol.Add(pDataMod)
    End Sub
    Public ReadOnly Property Item(ByVal pIndexKey As Object) As DataMod
      Get
        Item = CType(mvCol(pIndexKey), DataMod)
      End Get
    End Property
    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count
      End Get
    End Property
    Public Sub Remove(ByVal pIndexKey As Object)
      mvCol.Remove(pIndexKey.ToString)
    End Sub
    Public Sub New()
      mvCol = New Collection
    End Sub
  End Class
End Namespace
