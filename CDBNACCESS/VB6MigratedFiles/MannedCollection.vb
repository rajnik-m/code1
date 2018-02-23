

Namespace Access
  Public Class MannedCollection

    Public Enum MannedCollectionRecordSetTypes 'These are bit values
      mcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum MannedCollectionFields
      mcfAll = 0
      mcfCollectionNumber
      mcfCollectionDate
      mcfStartTime
      mcfEndTime
      mcfOrganisationNumber
      mcfMeetingPoint
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvAuthorisingOrganisation As Organisation

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "manned_collections"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("collection_date", CDBField.FieldTypes.cftDate)
          .Add("start_time", CDBField.FieldTypes.cftTime)
          .Add("end_time", CDBField.FieldTypes.cftTime)
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("meeting_point")
        End With

        mvClassFields.Item(MannedCollectionFields.mcfCollectionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(MannedCollectionFields.mcfCollectionNumber).PrefixRequired = True
        mvClassFields.Item(MannedCollectionFields.mcfOrganisationNumber).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvAuthorisingOrganisation = Nothing
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As MannedCollectionFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As MannedCollectionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = MannedCollectionRecordSetTypes.mcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "mc")
      End If
      GetRecordSetFields = vFields
    End Function

    Friend Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(MannedCollectionRecordSetTypes.mcrtAll) & " FROM manned_collections mc WHERE collection_number = " & pCollectionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, MannedCollectionRecordSetTypes.mcrtAll)
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

    Friend Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As MannedCollectionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(MannedCollectionFields.mcfCollectionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And MannedCollectionRecordSetTypes.mcrtAll) = MannedCollectionRecordSetTypes.mcrtAll Then
          .SetItem(MannedCollectionFields.mcfCollectionDate, vFields)
          .SetItem(MannedCollectionFields.mcfStartTime, vFields)
          .SetItem(MannedCollectionFields.mcfEndTime, vFields)
          .SetItem(MannedCollectionFields.mcfOrganisationNumber, vFields)
          .SetItem(MannedCollectionFields.mcfMeetingPoint, vFields)
        End If
      End With
    End Sub

    Friend Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(MannedCollectionFields.mcfCollectionDate).Value = pParams("CollectionDate").Value
        .Item(MannedCollectionFields.mcfStartTime).Value = pParams("StartTime").Value
        .Item(MannedCollectionFields.mcfEndTime).Value = pParams("EndTime").Value
        .Item(MannedCollectionFields.mcfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("MeetingPoint") Then .Item(MannedCollectionFields.mcfMeetingPoint).Value = pParams("MeetingPoint").Value
      End With
    End Sub

    Friend Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("CollectionDate") Then .Item(MannedCollectionFields.mcfCollectionDate).Value = pParams("CollectionDate").Value
        If pParams.Exists("StartTime") Then .Item(MannedCollectionFields.mcfStartTime).Value = pParams("StartTime").Value
        If pParams.Exists("EndTime") Then .Item(MannedCollectionFields.mcfEndTime).Value = pParams("EndTime").Value
        If pParams.Exists("OrganisationNumber") Then .Item(MannedCollectionFields.mcfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("MeetingPoint") Then .Item(MannedCollectionFields.mcfMeetingPoint).Value = pParams("MeetingPoint").Value
      End With
    End Sub
    Friend Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Friend Sub Save(Optional ByRef pAudit As Boolean = False, Optional ByRef pCollectionNumber As Integer = 0)
      SetValid(MannedCollectionFields.mcfAll)
      If pCollectionNumber > 0 Then mvClassFields(MannedCollectionFields.mcfCollectionNumber).IntegerValue = pCollectionNumber
      mvClassFields.Save(mvEnv, mvExisting, "", pAudit)
    End Sub

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcMannedCollection As MannedCollection)
      Init(pEnv)

      With mvClassFields
        .Item(MannedCollectionFields.mcfCollectionDate).Value = pSrcMannedCollection.CollectionDate
        .Item(MannedCollectionFields.mcfStartTime).Value = pSrcMannedCollection.StartTime
        .Item(MannedCollectionFields.mcfEndTime).Value = pSrcMannedCollection.EndTime
        .Item(MannedCollectionFields.mcfOrganisationNumber).Value = CStr(pSrcMannedCollection.OrganisationNumber)
        If pSrcMannedCollection.MeetingPoint.Length > 0 Then .Item(MannedCollectionFields.mcfMeetingPoint).Value = pSrcMannedCollection.MeetingPoint
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

    Public ReadOnly Property AuthorisingOrganisation() As Organisation
      Get
        If mvAuthorisingOrganisation Is Nothing Then
          mvAuthorisingOrganisation = New Organisation(mvEnv)
          mvAuthorisingOrganisation.Init((mvClassFields.Item(MannedCollectionFields.mcfOrganisationNumber).IntegerValue))
        End If
        AuthorisingOrganisation = mvAuthorisingOrganisation
      End Get
    End Property

    Public ReadOnly Property CollectionDate() As String
      Get
        CollectionDate = mvClassFields.Item(MannedCollectionFields.mcfCollectionDate).Value
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(MannedCollectionFields.mcfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EndTime() As String
      Get
        EndTime = mvClassFields.Item(MannedCollectionFields.mcfEndTime).Value
      End Get
    End Property

    Public ReadOnly Property MeetingPoint() As String
      Get
        MeetingPoint = mvClassFields.Item(MannedCollectionFields.mcfMeetingPoint).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(MannedCollectionFields.mcfOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StartTime() As String
      Get
        StartTime = mvClassFields.Item(MannedCollectionFields.mcfStartTime).Value
      End Get
    End Property
  End Class
End Namespace
