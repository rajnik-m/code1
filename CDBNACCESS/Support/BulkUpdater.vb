Public Class BulkUpdater

  Private mvEnv As CDBEnvironment
  Private mvItems As List(Of CARERecord)

  Public Sub New(ByVal pEnv As CDBEnvironment)
    Environment = pEnv
    mvItems = New List(Of CARERecord)
  End Sub

  Public Sub AddItem(ByVal pItem As CARERecord)
    If mvItems Is Nothing Then mvItems = New List(Of CARERecord)
    ItemList.Add(pItem)
  End Sub

  ''' <summary>Save the changes in bulk.</summary>
  ''' <param name="pSQLStatement">The SQL statement required by the BulkUpdate</param>
  Public Sub SaveBulkUpdate(ByVal pSQLStatement As SQLStatement)
    If ItemList.Count > 0 AndAlso pSQLStatement IsNot Nothing Then
      'This currently expects all the CareRecord classes to be for the same database table.
      Dim vTable As DataTable = CARERecord.GetBulkCopyDataTable(ItemList)
      If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
        'The BulkUpdate requires primary key fields to be defined on the DataTable
        Dim vUniqueKeyFields As CDBFields = ItemList(0).GetUniqueKeyFields()
        Dim vPrimaryKey(vUniqueKeyFields.Count) As DataColumn
        For vIndex As Integer = 1 To vUniqueKeyFields.Count
          vPrimaryKey(vIndex - 1) = vTable.Columns(vUniqueKeyFields(vIndex).Name)
        Next
        vTable.PrimaryKey = vPrimaryKey
        Me.Environment.Connection.BulkUpdate(pSQLStatement, vTable)
      End If
    End If
  End Sub

  Private Property Environment() As CDBEnvironment
    Get
      Return mvEnv
    End Get
    Set(value As CDBEnvironment)
      mvEnv = value
    End Set
  End Property

  Private ReadOnly Property ItemList() As List(Of CARERecord)
    Get
      Return mvItems
    End Get
  End Property

End Class
