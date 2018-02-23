Namespace Access

  Public Class Contacts
    Implements IEnumerable

    Private mvList As CollectionList(Of Contact)
    Private mvEnv As CDBEnvironment

    Public Function Add(ByVal pKey As String) As Contact
      'create a new object
      Dim vContact As Contact

      vContact = New Contact(mvEnv)
      mvList.Add(pKey, vContact)
      Return vContact
    End Function

    Public Sub New(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      mvList = New CollectionList(Of Contact)
    End Sub

    Public Function Exists(ByVal pKey As String) As Boolean
      Return mvList.ContainsKey(pKey)
    End Function

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      Return mvList.GetEnumerator
    End Function

    Public Function Count() As Integer
      Return mvList.Count
    End Function

    Public Sub Remove(ByVal pKey As String)
      mvList.Remove(pKey)
    End Sub

    Default Public ReadOnly Property Item(ByVal pIndex As Integer) As Contact
      Get
        Return mvList(pIndex)
      End Get
    End Property

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet)
      Dim vContact As Contact

      While pRecordSet.Fetch()
        vContact = New Contact(mvEnv)
        vContact.InitFromRecordSet(pEnv, pRecordSet, Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtVAT Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtDetail)
        If Not mvList.ContainsKey(vContact.ContactNumber.ToString) Then
          'if a contact appears in the select more than once ony add the first one
          mvList.Add(vContact.ContactNumber.ToString, vContact)
        End If
      End While
      pRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromAddress(ByVal pAddressNumber As Integer, ByVal pContactNumber As Integer)
      Dim vContact As Contact

      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT contact_number FROM contacts WHERE  address_number = " & pAddressNumber & " AND contact_number <> " & pContactNumber)
      mvList = New CollectionList(Of Contact)
      While vRecordSet.Fetch()
        vContact = New Contact(mvEnv)
        vContact.Init(vRecordSet.Fields(1).IntegerValue)
        If vContact.Existing Then
          mvList.Add(vContact.ContactNumber.ToString, vContact)
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitAllFromAddress(ByVal pAddressNumber As Integer, ByVal pContactNumber As Integer)
      Dim vContact As New Contact(mvEnv)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("ca.address_number", pAddressNumber)
      If pContactNumber > 0 Then vWhereFields.Add("c.contact_number", pContactNumber, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("contacts c", "c.contact_number", "ca.contact_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vContact.GetRecordSetFields & ",historical", "contact_addresses ca", vWhereFields, "", vAnsiJoins)
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      mvList = New CollectionList(Of Contact)
      While vRecordSet.Fetch()
        vContact = New Contact(mvEnv)
        vContact.InitFromRecordSet(vRecordSet)
        If vContact.Existing Then
          vContact.AddressHistorical = vRecordSet.Fields("historical").Bool
          mvList.Add(vContact.ContactNumber.ToString, vContact)
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub

  End Class

End Namespace
