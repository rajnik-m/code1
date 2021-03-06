Namespace Access

  Public Class CpdCycleStatus
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum CpdCycleStatuseFields
      AllFields = 0
      CpdCycleStatus
      CpdCycleStatusDesc
      CpdCycleType
      ManualStatus
      PointsFrom
      PointsTo
      RgbValue
      CreatedBy
      CreatedOn
      WebPublish
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("cpd_cycle_status")
        .Add("cpd_cycle_status_desc")
        .Add("cpd_cycle_type")
        .Add("manual_status")
        .Add("points_from", CDBField.FieldTypes.cftInteger)
        .Add("points_to", CDBField.FieldTypes.cftInteger)
        .Add("rgb_value", CDBField.FieldTypes.cftInteger)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)
        .Add("web_publish")

        .Item(CpdCycleStatuseFields.CpdCycleStatus).PrimaryKey = True
        .Item(CpdCycleStatuseFields.CpdCycleStatus).PrefixRequired = True

        .Item(CpdCycleStatuseFields.CpdCycleType).PrefixRequired = True
        .Item(CpdCycleStatuseFields.CreatedBy).PrefixRequired = True
        .Item(CpdCycleStatuseFields.CreatedOn).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ccs"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "cpd_cycle_statuses"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property CpdCycleStatus() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.CpdCycleStatus).Value
      End Get
    End Property
    Public ReadOnly Property CpdCycleStatusDesc() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.CpdCycleStatusDesc).Value
      End Get
    End Property
    Public ReadOnly Property CpdCycleType() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.CpdCycleType).Value
      End Get
    End Property
    Public ReadOnly Property ManualStatus() As Boolean
      Get
        Return mvClassFields(CpdCycleStatuseFields.ManualStatus).Bool
      End Get
    End Property
    Public ReadOnly Property PointsFrom() As Integer
      Get
        Return mvClassFields(CpdCycleStatuseFields.PointsFrom).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PointsTo() As Integer
      Get
        Return mvClassFields(CpdCycleStatuseFields.PointsTo).IntegerValue
      End Get
    End Property
    Public ReadOnly Property RgbValue() As Integer
      Get
        Return mvClassFields(CpdCycleStatuseFields.RgbValue).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CpdCycleStatuseFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property WebPublish() As Boolean
      Get
        Return mvClassFields(CpdCycleStatuseFields.WebPublish).Bool
      End Get
    End Property
#End Region

  End Class
End Namespace
