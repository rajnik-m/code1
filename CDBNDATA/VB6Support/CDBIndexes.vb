Namespace Data

  Public Class CDBIndexes
    Private mvIndexes As List(Of CDBIndex) = Nothing
    Private mvTableName As String

    Public Sub Init(ByVal pConn As CDBConnection, ByVal pTableName As String)
      mvIndexes = New List(Of CDBIndex)
      mvTableName = pTableName
      Dim vTable As DataTable = pConn.GetIndexNames(pTableName)
      For Each vRow As DataRow In vTable.Rows
        Dim vIndex As New CDBIndex(pConn, pTableName, vRow("INDEX_NAME").ToString.ToLower, vRow("UNIQUE").ToString = "Y")
        mvIndexes.Add(vIndex)
      Next
    End Sub

    Public ReadOnly Property Count() As Integer
      Get
        Return mvIndexes.Count
      End Get
    End Property

    Public Sub CreateIfMissing(ByVal pConn As CDBConnection, ByVal pUnique As Boolean, pAttributes As IList(Of String))
      If Not pConn.IndexExists(mvTableName, pAttributes) Then
        pConn.CreateIndex(pUnique, mvTableName, pAttributes)
      End If
    End Sub

    Public ReadOnly Property List() As List(Of CDBIndex)
      Get
        Return mvIndexes
      End Get
    End Property

    Public Sub DropAll(ByVal pConn As CDBConnection)
      For Each vIndex As CDBIndex In mvIndexes
        vIndex.Drop()
      Next
    End Sub

    Public Sub ReCreate(ByVal pConn As CDBConnection)
      If mvIndexes IsNot Nothing Then
        For Each vIndex As CDBIndex In mvIndexes
          If Not pConn.IndexExists(vIndex.TableName, New List(Of String)(vIndex.FieldNames.AsIEnumerable)) Then
            vIndex.ReCreate()
          End If
        Next
      End If
    End Sub

  End Class

End Namespace
