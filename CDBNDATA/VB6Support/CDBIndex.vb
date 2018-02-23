Namespace Data

  Public Class CDBIndex
    Private mvUnique As Boolean
    Private mvIndexName As String
    Private mvTableName As String
    Private mvFieldNames As New StringList("")
    Private mvDropped As Boolean
    Private mvConn As CDBConnection

    Public Sub New(ByVal pConn As CDBConnection, ByVal pTableName As String, ByVal pIndexName As String, ByVal pUnique As Boolean)
      mvConn = pConn
      mvTableName = pTableName
      mvIndexName = pIndexName
      mvUnique = pUnique
      Dim vColumns As DataTable = pConn.GetIndexColumns(pTableName, pIndexName)
      For Each vRow As DataRow In vColumns.Rows
        mvFieldNames.Add(vRow("COLUMN_NAME").ToString.ToLower)
      Next
    End Sub

    Public Property TableName() As String
      Get
        TableName = mvTableName
      End Get
      Set(ByVal Value As String)
        mvTableName = Value
      End Set
    End Property

    Public Function GetFieldName(ByVal pIndex As Integer) As String
      If pIndex < mvFieldNames.Count Then
        Return mvFieldNames(pIndex)
      End If
      Return ""
    End Function

    Public ReadOnly Property Dropped() As Boolean
      Get
        Return mvDropped
      End Get
    End Property

    Public ReadOnly Property FieldNames() As String
      Get
        Return mvFieldNames.ItemList
      End Get
    End Property

    Public ReadOnly Property IndexName() As String
      Get
        Return mvIndexName
      End Get
    End Property

    Public ReadOnly Property GeneratedIndexName As String
      Get
        Return mvConn.GetIndexName(mvTableName, mvFieldNames)
      End Get
    End Property

    Public ReadOnly Property Unique() As Boolean
      Get
        Return mvUnique
      End Get
    End Property

    Public Sub Drop()
      mvConn.DropIndexByName(mvTableName, mvIndexName)
      mvDropped = True
    End Sub

    Public Sub ReCreate()
      mvConn.CreateIndex(mvUnique, mvTableName, mvFieldNames)
    End Sub

  End Class

End Namespace