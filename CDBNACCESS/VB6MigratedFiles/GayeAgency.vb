

Namespace Access
  Public Class GayeAgency

    Public Enum GayeAgencyRecordSetTypes 'These are bit values
      gartAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GayeAgencyFields
      gafAll = 0
      gafOrganisationNumber
      gafAmendedBy
      gafAmendedOn
      gafPostBatchesToCb
      gafAdminFeePercentage
      gafMinAdminFee
      gafMaxAdminFee
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvOrganisation As Organisation
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "gaye_agencies"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("post_batches_to_cb")
          .Add("admin_fee_percentage")
          .Add("minimum_admin_fee")
          .Add("maximum_admin_fee")
        End With

        mvClassFields.Item(GayeAgencyFields.gafOrganisationNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(GayeAgencyFields.gafOrganisationNumber).PrefixRequired = True
        mvClassFields.Item(GayeAgencyFields.gafAdminFeePercentage).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAgencyAdminFee)
        mvClassFields.Item(GayeAgencyFields.gafMinAdminFee).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAgencyAdminFee)
        mvClassFields.Item(GayeAgencyFields.gafMaxAdminFee).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataAgencyAdminFee)
        mvClassFields.Item(GayeAgencyFields.gafAmendedBy).PrefixRequired = True
        mvClassFields.Item(GayeAgencyFields.gafAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As GayeAgencyFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(GayeAgencyFields.gafAmendedOn).Value = TodaysDate()
      mvClassFields.Item(GayeAgencyFields.gafAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GayeAgencyRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GayeAgencyRecordSetTypes.gartAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ga")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pOrganisationNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pOrganisationNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GayeAgencyRecordSetTypes.gartAll) & " FROM gaye_agencies ga WHERE organisation_number = " & pOrganisationNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GayeAgencyRecordSetTypes.gartAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GayeAgencyRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GayeAgencyFields.gafOrganisationNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GayeAgencyRecordSetTypes.gartAll) = GayeAgencyRecordSetTypes.gartAll Then
          .SetItem(GayeAgencyFields.gafAmendedBy, vFields)
          .SetItem(GayeAgencyFields.gafAmendedOn, vFields)
          .SetItem(GayeAgencyFields.gafPostBatchesToCb, vFields)
          .SetOptionalItem(GayeAgencyFields.gafAdminFeePercentage, vFields)
          .SetOptionalItem(GayeAgencyFields.gafMinAdminFee, vFields)
          .SetOptionalItem(GayeAgencyFields.gafMaxAdminFee, vFields)
        End If

        mvOrganisation = New Organisation(mvEnv)
        mvOrganisation.Init((mvClassFields.Item(GayeAgencyFields.gafOrganisationNumber).IntegerValue))
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GayeAgencyFields.gafAll)
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
        AmendedBy = mvClassFields.Item(GayeAgencyFields.gafAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(GayeAgencyFields.gafAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(GayeAgencyFields.gafOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PostBatchesToCb() As String
      Get
        PostBatchesToCb = mvClassFields.Item(GayeAgencyFields.gafPostBatchesToCb).Value
      End Get
    End Property

    Public ReadOnly Property Organisation() As Organisation
      Get
        If mvOrganisation Is Nothing Then
          mvOrganisation = New Organisation(mvEnv)
          mvOrganisation.Init((mvClassFields.Item(GayeAgencyFields.gafOrganisationNumber).IntegerValue))
        End If
        Organisation = mvOrganisation
      End Get
    End Property

    Public ReadOnly Property AdminFeePercentage() As String
      Get
        AdminFeePercentage = mvClassFields.Item(GayeAgencyFields.gafAdminFeePercentage).Value
      End Get
    End Property

    Public ReadOnly Property MinAdminFee() As String
      Get
        MinAdminFee = mvClassFields.Item(GayeAgencyFields.gafMinAdminFee).Value
      End Get
    End Property

    Public ReadOnly Property MaxAdminFee() As String
      Get
        MaxAdminFee = mvClassFields.Item(GayeAgencyFields.gafMaxAdminFee).Value
      End Get
    End Property
  End Class
End Namespace
