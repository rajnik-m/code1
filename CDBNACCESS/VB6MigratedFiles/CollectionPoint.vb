

Namespace Access
  Public Class CollectionPoint

    Public Enum CollectionPointRecordSetTypes 'These are bit values
      cptrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionPointFields
      cpfAll = 0
      cpfCollectionPointNumber
      cpfCollectionRegionNumber
      cpfCollectionPointType
      cpfCollectionPoint
      cpfOrganisationNumber
      cpfAddressNumber
      cpfNoOfCollectors
      cpfNotes
      cpfAmendedBy
      cpfAmendedOn
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
          .DatabaseTableName = "collection_points"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_point_number", CDBField.FieldTypes.cftLong)
          .Add("collection_region_number", CDBField.FieldTypes.cftLong)
          .Add("collection_point_type")
          .Add("collection_point")
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("no_of_collectors", CDBField.FieldTypes.cftLong)
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(CollectionPointFields.cpfCollectionPointNumber).SetPrimaryKeyOnly()
          .Item(CollectionPointFields.cpfCollectionRegionNumber).PrefixRequired = True
          .Item(CollectionPointFields.cpfOrganisationNumber).PrefixRequired = True
          .Item(CollectionPointFields.cpfAddressNumber).PrefixRequired = True
          .Item(CollectionPointFields.cpfNotes).PrefixRequired = True

          .SetUniqueField(CollectionPointFields.cpfCollectionRegionNumber)
          .SetUniqueField(CollectionPointFields.cpfCollectionPoint)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionPointFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(CollectionPointFields.cpfCollectionPointNumber).IntegerValue = 0 Then mvClassFields.Item(CollectionPointFields.cpfCollectionPointNumber).IntegerValue = mvEnv.GetControlNumber("CN")
      mvClassFields.Item(CollectionPointFields.cpfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectionPointFields.cpfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionPointRecordSetTypes) As String
      Dim vFields As String = ""

      'Modify below to add each recordset type as required
      If pRSType = CollectionPointRecordSetTypes.cptrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cp")
        vFields = Replace(vFields, "cp.notes", "cp.notes AS cp_notes")
        vFields = Replace(vFields, "cp.organisation_number", "cp.organisation_number AS cp_organisation_number")
        vFields = Replace(vFields, "cp.address_number", "cp.address_number AS cp_address_number")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionPointNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionPointNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionPointRecordSetTypes.cptrtAll) & " FROM collection_points cp WHERE collection_point_number = " & pCollectionPointNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionPointRecordSetTypes.cptrtAll)
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

    Public Sub InitFromPoint(ByVal pEnv As CDBEnvironment, ByRef pCollectionRegionNumber As Integer, ByRef pCollectionPoint As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      With vWhereFields
        .Add("collection_region_number", CDBField.FieldTypes.cftLong, pCollectionRegionNumber)
        .Add("collection_point", CDBField.FieldTypes.cftCharacter, pCollectionPoint)
      End With
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionPointRecordSetTypes.cptrtAll) & " FROM collection_points cp WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CollectionPointRecordSetTypes.cptrtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionPointRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionPointFields.cpfCollectionPointNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionPointRecordSetTypes.cptrtAll) = CollectionPointRecordSetTypes.cptrtAll Then
          .SetItem(CollectionPointFields.cpfCollectionRegionNumber, vFields)
          .SetItem(CollectionPointFields.cpfCollectionPointType, vFields)
          .SetItem(CollectionPointFields.cpfCollectionPoint, vFields)
          .Item(CollectionPointFields.cpfOrganisationNumber).SetValue = vFields("cp_organisation_number").Value
          '      .SetItem cpfOrganisationNumber, vFields
          .Item(CollectionPointFields.cpfAddressNumber).SetValue = vFields("cp_address_number").Value
          '      .SetItem cpfAddressNumber, vFields
          .SetItem(CollectionPointFields.cpfNoOfCollectors, vFields)
          .Item(CollectionPointFields.cpfNotes).SetValue = vFields("cp_notes").Value
          '      .SetItem cpfNotes, vFields
          .SetItem(CollectionPointFields.cpfAmendedBy, vFields)
          .SetItem(CollectionPointFields.cpfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectionPointFields.cpfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(CollectionPointFields.cpfCollectionRegionNumber).Value = pParams("CollectionRegionNumber").Value
        .Item(CollectionPointFields.cpfCollectionPointType).Value = pParams("CollectionPointType").Value
        .Item(CollectionPointFields.cpfCollectionPoint).Value = pParams("CollectionPoint").Value
        If pParams.Exists("OrganisationNumber") Then .Item(CollectionPointFields.cpfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(CollectionPointFields.cpfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("NoOfCollectors") Then .Item(CollectionPointFields.cpfNoOfCollectors).Value = pParams("NoOfCollectors").Value
        If pParams.Exists("Notes") Then .Item(CollectionPointFields.cpfNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("CollectionPointType") Then .Item(CollectionPointFields.cpfCollectionPointType).Value = pParams("CollectionPointType").Value
        If pParams.Exists("CollectionPoint") Then .Item(CollectionPointFields.cpfCollectionPoint).Value = pParams("CollectionPoint").Value
        If pParams.Exists("OrganisationNumber") Then .Item(CollectionPointFields.cpfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(CollectionPointFields.cpfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("NoOfCollectors") Then .Item(CollectionPointFields.cpfNoOfCollectors).Value = pParams("NoOfCollectors").Value
        If pParams.Exists("Notes") Then .Item(CollectionPointFields.cpfNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("collection_point_number", CDBField.FieldTypes.cftLong, CollectionPointNumber)
      mvEnv.Connection.StartTransaction()
      mvEnv.Connection.DeleteRecords("collector_shifts", vWhereFields, False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      mvEnv.Connection.CommitTransaction()
    End Sub

    'Private Function DeleteAllowed() As Boolean
    '  Dim vWhereFields As New CDBFields
    '  Dim vDeleteAllowed  As Boolean
    '
    '  vDeleteAllowed = True
    '  vWhereFields.Add "collection_point_number", cftLong, CollectionPointNumber
    '  If mvEnv.Connection.GetCount("collector_shifts", vWhereFields) > 0 Then vDeleteAllowed = False
    '  DeleteAllowed = vDeleteAllowed
    'End Function

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcCollectionPoint As CollectionPoint, ByVal pCollectionRegionNumber As Integer)
      Init(pEnv)

      With mvClassFields
        .Item(CollectionPointFields.cpfCollectionRegionNumber).Value = CStr(pCollectionRegionNumber)
        .Item(CollectionPointFields.cpfCollectionPointType).Value = pSrcCollectionPoint.CollectionPointType
        .Item(CollectionPointFields.cpfCollectionPoint).Value = pSrcCollectionPoint.CollectionPointDesc
        If pSrcCollectionPoint.OrganisationNumber > 0 Then .Item(CollectionPointFields.cpfOrganisationNumber).Value = CStr(pSrcCollectionPoint.OrganisationNumber)
        If pSrcCollectionPoint.AddressNumber > 0 Then .Item(CollectionPointFields.cpfAddressNumber).Value = CStr(pSrcCollectionPoint.AddressNumber)
        If pSrcCollectionPoint.NoOfCollectors > 0 Then .Item(CollectionPointFields.cpfNoOfCollectors).Value = CStr(pSrcCollectionPoint.NoOfCollectors)
        If pSrcCollectionPoint.Notes.Length > 0 Then .Item(CollectionPointFields.cpfNotes).Value = pSrcCollectionPoint.Notes
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

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(CollectionPointFields.cpfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CollectionPointFields.cpfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionPointFields.cpfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CollectionPointDesc() As String
      Get
        CollectionPointDesc = mvClassFields.Item(CollectionPointFields.cpfCollectionPoint).Value
      End Get
    End Property

    Public ReadOnly Property CollectionPointNumber() As Integer
      Get
        CollectionPointNumber = mvClassFields.Item(CollectionPointFields.cpfCollectionPointNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPointType() As String
      Get
        CollectionPointType = mvClassFields.Item(CollectionPointFields.cpfCollectionPointType).Value
      End Get
    End Property

    Public ReadOnly Property CollectionRegionNumber() As Integer
      Get
        CollectionRegionNumber = mvClassFields.Item(CollectionPointFields.cpfCollectionRegionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property NoOfCollectors() As Integer
      Get
        NoOfCollectors = mvClassFields.Item(CollectionPointFields.cpfNoOfCollectors).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(CollectionPointFields.cpfNotes).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(CollectionPointFields.cpfOrganisationNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
