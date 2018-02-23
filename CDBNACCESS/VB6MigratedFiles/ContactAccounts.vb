Namespace Access

  Public Class ContactAccounts
    Implements System.Collections.IEnumerable
    ' BR 11347 - Created

    Private mvCol As New Collection

    Public Function Add(ByRef pKey As String) As ContactAccount
      'create a new object
      Dim vContactAccount As ContactAccount

      vContactAccount = New ContactAccount
      mvCol.Add(vContactAccount, pKey)
      Add = vContactAccount
      vContactAccount = Nothing
    End Function

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet)
      Dim vContactAccount As ContactAccount

      While pRecordSet.Fetch() = True
        vContactAccount = New ContactAccount
        vContactAccount.InitFromRecordSet(pEnv, pRecordSet, ContactAccount.ContactAccountRecordSetTypes.cartAmendedAlias Or ContactAccount.ContactAccountRecordSetTypes.cartDetails Or ContactAccount.ContactAccountRecordSetTypes.cartNumber)
        mvCol.Add(vContactAccount, CStr(vContactAccount.ContactNumber))
      End While
      pRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property Item(ByVal pIndexKey As String) As ContactAccount
      Get
        Item = DirectCast(mvCol.Item(pIndexKey), ContactAccount)
      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public Function Exists(ByVal pIndexKey As String) As Boolean
      Return mvCol.Contains(pIndexKey)
    End Function

    Public Sub Remove(ByVal pIndexKey As String)
      mvCol.Remove(pIndexKey)
    End Sub

  End Class
End Namespace
