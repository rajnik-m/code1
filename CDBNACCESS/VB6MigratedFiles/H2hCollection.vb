

Namespace Access
  Public Class H2hCollection

    Public Enum H2hCollectionRecordSetTypes 'These are bit values
      hcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum H2hCollectionFields
      hcfAll = 0
      hcfCollectionNumber
      hcfStartDate
      hcfEndDate
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "h2h_collections"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("end_date", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(H2hCollectionFields.hcfCollectionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(H2hCollectionFields.hcfCollectionNumber).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As H2hCollectionFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As H2hCollectionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = H2hCollectionRecordSetTypes.hcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "hc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(H2hCollectionRecordSetTypes.hcrtAll) & " FROM h2h_collections hc WHERE collection_number = " & pCollectionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, H2hCollectionRecordSetTypes.hcrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As H2hCollectionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(H2hCollectionFields.hcfCollectionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And H2hCollectionRecordSetTypes.hcrtAll) = H2hCollectionRecordSetTypes.hcrtAll Then
          .SetItem(H2hCollectionFields.hcfStartDate, vFields)
          .SetItem(H2hCollectionFields.hcfEndDate, vFields)
        End If
      End With
    End Sub

    Friend Sub Save(Optional ByRef pAudit As Boolean = False, Optional ByRef pCollectionNumber As Integer = 0)
      SetValid(H2hCollectionFields.hcfAll)
      If pCollectionNumber > 0 Then mvClassFields(H2hCollectionFields.hcfCollectionNumber).IntegerValue = pCollectionNumber
      mvClassFields.Save(mvEnv, mvExisting, "", pAudit)
    End Sub

    Friend Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(H2hCollectionFields.hcfStartDate).Value = pParams("StartDate").Value
        .Item(H2hCollectionFields.hcfEndDate).Value = pParams("EndDate").Value
      End With
    End Sub

    Friend Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("StartDate") Then .Item(H2hCollectionFields.hcfStartDate).Value = pParams("StartDate").Value
        If pParams.Exists("EndDate") Then .Item(H2hCollectionFields.hcfEndDate).Value = pParams("EndDate").Value
      End With
    End Sub

    Friend Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcH2hCollection As H2hCollection)
      Init(pEnv)

      With mvClassFields
        .Item(H2hCollectionFields.hcfStartDate).Value = pSrcH2hCollection.StartDate
        .Item(H2hCollectionFields.hcfEndDate).Value = pSrcH2hCollection.EndDate
      End With

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        Return mvClassFields.Item(H2hCollectionFields.hcfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EndDate() As String
      Get
        EndDate = mvClassFields.Item(H2hCollectionFields.hcfEndDate).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(H2hCollectionFields.hcfStartDate).Value
      End Get
    End Property
  End Class
End Namespace
