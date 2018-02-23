Namespace Access
  Public Class UpgradeIndices
    Public mvCol As Hashtable


    Public ReadOnly Property Exists(ByVal pIndexKey As String) As Boolean
      Get
        Return mvCol.ContainsKey(pIndexKey)
      End Get
    End Property
    Public ReadOnly Property Item(ByVal pIndexKey As Object) As UpgradeIndex
      Get
        Item = CType(mvCol(pIndexKey), UpgradeIndex)
      End Get
    End Property
    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count
      End Get
    End Property
    Public Function Add(ByVal pKey As String, ByVal pUniqueIndex As Boolean, ByVal pPrimaryValue As Integer, ByVal pOverflowValue As Integer, ByVal pAttributenames As IList(Of String)) As UpgradeIndex
      'create a new object
      Dim vNewMember As New UpgradeIndex With {.Key = pKey,
                                               .UniqueIndex = pUniqueIndex,
                                               .PrimaryValue = pPrimaryValue,
                                               .OverflowValue = pOverflowValue,
                                               .ToBeCreated = True,
                                               .ToBeDeleted = False}
      For Each vAttribute In pAttributenames
        vNewMember.Attributes.Add(vAttribute)
      Next vAttribute
      mvCol.Add(pKey, vNewMember)
      'return the object created
      Return vNewMember
    End Function
    Public Sub Remove(ByVal pIndexKey As Object)
      mvCol.Remove(pIndexKey.ToString())
    End Sub

    Public Sub New()
      mvCol = New Hashtable
    End Sub
  End Class
End Namespace
