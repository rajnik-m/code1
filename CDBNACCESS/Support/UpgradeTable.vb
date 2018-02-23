Namespace Access
  Public Class UpgradeTable
    Public Key As String

    Private mvUpgradeAttributes As UpgradeAttributes
    Private mvStructureModified As Boolean
    Private mvHasUnicode As Boolean
    Private mvToBeCreated As Boolean
    Private mvToBeDropped As Boolean
    Private mvUpgradeIndices As UpgradeIndices
    Private mvNumberOfRecords As Integer
    Private mvCountValid As Boolean
    Private mvConnection As CDBConnection
    Private mvDataMods As DataMods
    Private mvTableHits As Integer
    Public mvChanges As Hashtable

    Public ReadOnly Property ChangeExists(ByVal pIndexKey As String) As Boolean
      Get
        Return mvChanges.ContainsKey(pIndexKey)
      End Get
    End Property

    Public Property DataMods() As DataMods
      Get
        If mvDataMods Is Nothing Then
          mvDataMods = New DataMods
        End If
        DataMods = mvDataMods
      End Get
      Set(ByVal value As DataMods)
        mvDataMods = value
      End Set
    End Property

    Public Property Connection() As CDBConnection
      Get
        Return mvConnection
      End Get
      Set(ByVal value As CDBConnection)
        mvConnection = value
      End Set
    End Property

    Public Property RecordCount() As Integer
      Get
        If Not mvCountValid Then
          If mvToBeCreated = False Then
            mvNumberOfRecords = mvConnection.GetCount(Key, Nothing)
          End If
          mvCountValid = True
        End If
        RecordCount = mvNumberOfRecords
      End Get
      Set(ByVal value As Integer)
        mvNumberOfRecords = value
        mvCountValid = True
      End Set
    End Property
    Public Property UpgradeIndices() As UpgradeIndices
      Get
        If mvUpgradeIndices Is Nothing Then
          mvUpgradeIndices = New UpgradeIndices
        End If
        UpgradeIndices = mvUpgradeIndices
      End Get
      Set(ByVal value As UpgradeIndices)
        mvUpgradeIndices = value
      End Set
    End Property

    Public Property ToBeDropped() As Boolean
      Get
        Return mvToBeDropped
      End Get
      Set(ByVal value As Boolean)
        mvToBeDropped = value
      End Set
    End Property

    Public Property TableHits() As Integer
      Get
        Return mvTableHits
      End Get
      Set(ByVal value As Integer)
        mvTableHits = value
      End Set
    End Property
    Public Property ToBeCreated() As Boolean
      Get
        Return mvToBeCreated
      End Get
      Set(ByVal value As Boolean)
        mvToBeCreated = value
      End Set
    End Property

    Public Property StructureModified() As Boolean
      Get
        Return mvStructureModified
      End Get
      Set(ByVal value As Boolean)
        mvStructureModified = value
      End Set
    End Property
    Public Property HasUnicode() As Boolean
      Get
        Return mvHasUnicode
      End Get
      Set(ByVal value As Boolean)
        mvHasUnicode = value
      End Set
    End Property
    Public ReadOnly Property UpgradeAttributes() As UpgradeAttributes
      Get
        If mvUpgradeAttributes Is Nothing Then
          mvUpgradeAttributes = New UpgradeAttributes
          mvUpgradeAttributes.UpgradeTable = Me
        End If
        UpgradeAttributes = mvUpgradeAttributes
      End Get
    End Property

    Public Sub New()
      mvTableHits = 1
      mvChanges = New Hashtable
      mvUpgradeIndices = New UpgradeIndices()
    End Sub
    Public Function AddChange(ByVal pChangeNumber As Integer, ByVal pChangeComment As String) As UpgradeChange
      Dim vNewMember As UpgradeChange

      vNewMember = New UpgradeChange
      With vNewMember
        .ChangeNumber = pChangeNumber
        .ChangeComment = pChangeComment
        .ToBeApplied = False
      End With
      mvChanges.Add(Format(pChangeNumber), vNewMember)
      'return the object created
      AddChange = vNewMember
    End Function
    Public Function AddIndex(ByVal pConn As CDBConnection, ByVal pDB As String, ByVal pUniqueIndex As String, ByVal pPrimaryValue As Integer, ByVal pOverflowValue As Integer, pAttributeNames As IList(Of String), ByVal pToBeDropped As Boolean) As UpgradeIndex
      Dim vUpgradeIndex As UpgradeIndex

      vUpgradeIndex = New UpgradeIndex
      Dim vHashName As String = pConn.GetIndexName(Key, pAttributeNames)
      If Not mvUpgradeIndices.Exists(vHashName) Then
        vUpgradeIndex = mvUpgradeIndices.Add(vHashName, Not pUniqueIndex = "NORMAL", pPrimaryValue, pOverflowValue, pAttributeNames)
      Else
        vUpgradeIndex = mvUpgradeIndices.Item(vHashName)
        vUpgradeIndex.UniqueIndex = Not pUniqueIndex = "NORMAL"
        vUpgradeIndex.PrimaryValue = pPrimaryValue
        vUpgradeIndex.OverflowValue = pOverflowValue
      End If
      Dim vExists As Boolean = pConn.IndexExists(Key, pAttributeNames)

      If vExists Then
        If pToBeDropped Then
          vUpgradeIndex.ToBeDeleted = True
          vUpgradeIndex.ToBeCreated = False
        Else
          If pConn.IndexIsUnique(Key, pAttributeNames) = vUpgradeIndex.UniqueIndex Then
            vUpgradeIndex.ToBeDeleted = False
            vUpgradeIndex.ToBeCreated = False
          Else
            vUpgradeIndex.ToBeDeleted = True
            vUpgradeIndex.ToBeCreated = True
          End If
        End If
      Else
        If Not pToBeDropped Then
          vUpgradeIndex.ToBeDeleted = False
          vUpgradeIndex.ToBeCreated = True
        Else
          vUpgradeIndex.ToBeDeleted = False
          vUpgradeIndex.ToBeCreated = False
        End If
      End If
      Return vUpgradeIndex
    End Function
  End Class
End Namespace
