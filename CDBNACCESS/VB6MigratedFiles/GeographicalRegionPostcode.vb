

Namespace Access
  Public Class GeographicalRegionPostcode

    Public Enum GeographicalRegionPostcodeRecordSetTypes 'These are bit values
      grprtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GeographicalRegionPostcodeFields
      grpfAll = 0
      grpfGeographicalRegionType
      grpfGeographicalRegion
      grpfPostcode
      grpfAmendedBy
      grpfAmendedOn
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
          .DatabaseTableName = "geographical_region_postcodes"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("geographical_region_type")
          .Add("geographical_region")
          .Add("postcode")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(GeographicalRegionPostcodeFields.grpfGeographicalRegionType).SetPrimaryKeyOnly()
        mvClassFields.Item(GeographicalRegionPostcodeFields.grpfPostcode).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As GeographicalRegionPostcodeFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(GeographicalRegionPostcodeFields.grpfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(GeographicalRegionPostcodeFields.grpfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GeographicalRegionPostcodeRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GeographicalRegionPostcodeRecordSetTypes.grprtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "grp")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pGeographicalRegionType As String = "", Optional ByRef pGeographicalRegion As String = "", Optional ByRef pPostcode As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pGeographicalRegionType) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GeographicalRegionPostcodeRecordSetTypes.grprtAll) & " FROM geographical_region_postcodes grp WHERE geographical_region_type = '" & pGeographicalRegionType & "' AND geographical_region = '" & pGeographicalRegion & "' AND postcode = '" & pPostcode & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GeographicalRegionPostcodeRecordSetTypes.grprtAll)
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
    Public Sub MoveRegion(ByRef pNewRegionType As String, ByRef pNewRegion As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
 
      vUpdateFields.Add("geographical_region_type", CDBField.FieldTypes.cftCharacter, pNewRegionType)
      vUpdateFields.Add("geographical_region", CDBField.FieldTypes.cftCharacter, pNewRegion)
      vUpdateFields.AddAmendedOnBy((mvEnv.User.Logname))

      vWhereFields.Add("geographical_region_type", CDBField.FieldTypes.cftCharacter, GeographicalRegionType)
      vWhereFields.Add("geographical_region", CDBField.FieldTypes.cftCharacter, GeographicalRegion)
      vWhereFields.Add("postcode", CDBField.FieldTypes.cftCharacter, Postcode & "*", CDBField.FieldWhereOperators.fwoLike)
      mvEnv.Connection.UpdateRecords("address_geographical_regions", vUpdateFields, vWhereFields, False)

      GeographicalRegionType = pNewRegionType
      GeographicalRegion = pNewRegion
      Save()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GeographicalRegionPostcodeRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GeographicalRegionPostcodeFields.grpfGeographicalRegionType, vFields)
        .SetItem(GeographicalRegionPostcodeFields.grpfPostcode, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GeographicalRegionPostcodeRecordSetTypes.grprtAll) = GeographicalRegionPostcodeRecordSetTypes.grprtAll Then
          .SetItem(GeographicalRegionPostcodeFields.grpfGeographicalRegion, vFields)
          .SetItem(GeographicalRegionPostcodeFields.grpfAmendedBy, vFields)
          .SetItem(GeographicalRegionPostcodeFields.grpfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GeographicalRegionPostcodeFields.grpfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(GeographicalRegionPostcodeFields.grpfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(GeographicalRegionPostcodeFields.grpfAmendedOn).Value
      End Get
    End Property

    Public Property GeographicalRegion() As String
      Get
        GeographicalRegion = mvClassFields.Item(GeographicalRegionPostcodeFields.grpfGeographicalRegion).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(GeographicalRegionPostcodeFields.grpfGeographicalRegion).Value = Value
      End Set
    End Property

    Public Property GeographicalRegionType() As String
      Get
        GeographicalRegionType = mvClassFields.Item(GeographicalRegionPostcodeFields.grpfGeographicalRegionType).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(GeographicalRegionPostcodeFields.grpfGeographicalRegionType).Value = Value
      End Set
    End Property

    Public ReadOnly Property Postcode() As String
      Get
        Postcode = mvClassFields.Item(GeographicalRegionPostcodeFields.grpfPostcode).Value
      End Get
    End Property
  End Class
End Namespace
