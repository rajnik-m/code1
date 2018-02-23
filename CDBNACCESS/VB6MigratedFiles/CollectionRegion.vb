

Namespace Access
  Public Class CollectionRegion

    Public Enum CollectionRegionRecordSetTypes 'These are bit values
      crertAll = &H10S
      'ADD additional recordset types here
      crertAllPlusPoints = &H100S
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionRegionFields
      crfAll = 0
      crfCollectionRegionNumber
      crfCollectionNumber
      crfGeographicalRegion
      crfAmendedBy
      crfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvCollectionPoints As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "collection_regions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_region_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("geographical_region")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Item(CollectionRegionFields.crfCollectionRegionNumber).SetPrimaryKeyOnly()
          .Item(CollectionRegionFields.crfCollectionRegionNumber).PrefixRequired = True
          .Item(CollectionRegionFields.crfCollectionNumber).PrefixRequired = True

          .SetUniqueField(CollectionRegionFields.crfCollectionNumber)
          .SetUniqueField(CollectionRegionFields.crfGeographicalRegion)
        End With
      Else
        mvClassFields.ClearItems()
        mvCollectionPoints = New Collection
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionRegionFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).IntegerValue = 0 Then mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).IntegerValue = mvEnv.GetControlNumber("CG")
      mvClassFields.Item(CollectionRegionFields.crfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectionRegionFields.crfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    Private Sub InitCollectionPoints()
      Dim vRS As CDBRecordSet
      Dim vCollPoint As New CollectionPoint
      Dim vSQL As String

      mvCollectionPoints = New Collection
      vCollPoint.Init(mvEnv)
      vSQL = "SELECT " & vCollPoint.GetRecordSetFields(CollectionPoint.CollectionPointRecordSetTypes.cptrtAll) & " FROM collection_points cp WHERE collection_region_number = " & mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).IntegerValue
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch() = True
        vCollPoint = New CollectionPoint
        vCollPoint.InitFromRecordSet(mvEnv, vRS, CollectionPoint.CollectionPointRecordSetTypes.cptrtAll)
        mvCollectionPoints.Add(vCollPoint, CStr(vCollPoint.CollectionPointNumber))
      End While
      vRS.CloseRecordSet()

    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionRegionRecordSetTypes) As String
      Dim vFields As String = ""
      Dim vCollPoint As CollectionPoint

      'Modify below to add each recordset type as required
      If (pRSType And CollectionRegionRecordSetTypes.crertAll) > 0 Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cr")
      End If
      If (pRSType And CollectionRegionRecordSetTypes.crertAllPlusPoints) > 0 Then
        vCollPoint = New CollectionPoint
        vCollPoint.Init(mvEnv)
        vFields = vFields & "," & Replace(vCollPoint.GetRecordSetFields(CollectionPoint.CollectionPointRecordSetTypes.cptrtAll), "collection_region_number,", "")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionRegionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionRegionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionRegionRecordSetTypes.crertAll) & " FROM collection_regions cr WHERE collection_region_number = " & pCollectionRegionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionRegionRecordSetTypes.crertAll)
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

    Public Sub InitFromRegion(ByVal pEnv As CDBEnvironment, ByRef pCollectionNumber As Integer, ByRef pGeographicalRegion As String)
      Dim vRecordSet As CDBRecordSet
      Dim vwherefields As New CDBFields

      mvEnv = pEnv
      vwherefields.Add("collection_number", CDBField.FieldTypes.cftLong, pCollectionNumber)
      vwherefields.Add("geographical_region", CDBField.FieldTypes.cftCharacter, pGeographicalRegion)
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionRegionRecordSetTypes.crertAll) & " FROM collection_regions cr WHERE " & mvEnv.Connection.WhereClause(vwherefields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CollectionRegionRecordSetTypes.crertAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionRegionRecordSetTypes)
      Dim vFields As CDBFields
      Dim vCollPoint As CollectionPoint

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionRegionFields.crfCollectionRegionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionRegionRecordSetTypes.crertAll) = CollectionRegionRecordSetTypes.crertAll Then
          .SetItem(CollectionRegionFields.crfCollectionNumber, vFields)
          .SetItem(CollectionRegionFields.crfGeographicalRegion, vFields)
          .SetItem(CollectionRegionFields.crfAmendedBy, vFields)
          .SetItem(CollectionRegionFields.crfAmendedOn, vFields)
        End If
      End With

      If (pRSType And CollectionRegionRecordSetTypes.crertAllPlusPoints) > 0 Then
        mvCollectionPoints = New Collection
        While ((mvClassFields.Item(CollectionRegionFields.crfCollectionNumber).IntegerValue = pRecordSet.Fields("collection_number").IntegerValue) And (mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).IntegerValue = pRecordSet.Fields("collection_region_number").IntegerValue)) And pRecordSet.Status() = True
          If pRecordSet.Fields("collection_point_number").IntegerValue > 0 Then
            vCollPoint = New CollectionPoint
            vCollPoint.InitFromRecordSet(mvEnv, pRecordSet, CollectionPoint.CollectionPointRecordSetTypes.cptrtAll)
            mvCollectionPoints.Add(vCollPoint, CStr(vCollPoint.CollectionPointNumber))
          End If
          pRecordSet.Fetch()
        End While
      End If

    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectionRegionFields.crfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(CollectionRegionFields.crfCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(CollectionRegionFields.crfGeographicalRegion).Value = pParams("GeographicalRegion").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("GeographicalRegion") Then .Item(CollectionRegionFields.crfGeographicalRegion).Value = pParams("GeographicalRegion").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vwherefields As New CDBFields

      vwherefields.Add("collection_region_number", CDBField.FieldTypes.cftLong, CollectionRegionNumber)
      mvEnv.Connection.StartTransaction()
      'delete the shifts at the points
      mvEnv.Connection.ExecuteSQL("DELETE FROM collector_shifts WHERE collection_point_number IN(SELECT collection_point_number FROM collection_points WHERE " & mvEnv.Connection.WhereClause(vwherefields) & ")")
      'delete the points for this region
      mvEnv.Connection.DeleteRecords("collection_points", vwherefields, False)
      'delete the region
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      mvEnv.Connection.CommitTransaction()
    End Sub

    'Private Function DeleteAllowed() As Boolean
    '  Dim vWhereFields As New CDBFields
    '  Dim vAppColl     As New AppealCollection(mvEnv)
    '  Dim vDeleteAllowed  As Boolean
    '
    '  vDeleteAllowed = True
    '  If CollectionPoints.Count > 0 Then
    '  RaiseError
    '  DeleteAllowed = vDeleteAllowed
    'End Function

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcCollectionRegion As CollectionRegion, ByVal pCollectionNumber As Integer)
      Dim vSrcCollPoint As CollectionPoint
      Dim vTgtCollPoint As CollectionPoint

      Init(pEnv)

      With mvClassFields
        .Item(CollectionRegionFields.crfCollectionNumber).Value = CStr(pCollectionNumber)
        .Item(CollectionRegionFields.crfGeographicalRegion).Value = pSrcCollectionRegion.GeographicalRegion
      End With

      SetValid(CollectionRegionFields.crfAll)

      mvCollectionPoints = New Collection
      For Each vSrcCollPoint In pSrcCollectionRegion.CollectionPoints
        vTgtCollPoint = New CollectionPoint
        vTgtCollPoint.Clone(pEnv, vSrcCollPoint, mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).IntegerValue)
        mvCollectionPoints.Add(vTgtCollPoint)
      Next vSrcCollPoint

    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CollectionRegionFields.crfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionRegionFields.crfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = CInt(mvClassFields.Item(CollectionRegionFields.crfCollectionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectionPoints() As Collection
      Get
        If mvCollectionPoints Is Nothing Then InitCollectionPoints()
        CollectionPoints = mvCollectionPoints
      End Get
    End Property

    Public ReadOnly Property CollectionRegionNumber() As Integer
      Get
        CollectionRegionNumber = CInt(mvClassFields.Item(CollectionRegionFields.crfCollectionRegionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property GeographicalRegion() As String
      Get
        GeographicalRegion = mvClassFields.Item(CollectionRegionFields.crfGeographicalRegion).Value
      End Get
    End Property
  End Class
End Namespace
