Namespace Access

  Public Class CpdCycleType
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum CpdCycleTypeFields
      AllFields = 0
      CpdCycleType
      CpdCycleTypeDesc
      StartMonth
      EndMonth
      DefaultDuration
      CpdType
      WebPublish
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("cpd_cycle_type")
        .Add("cpd_cycle_type_desc")
        .Add("start_month", CDBField.FieldTypes.cftInteger)
        .Add("end_month", CDBField.FieldTypes.cftInteger)
        .Add("default_duration", CDBField.FieldTypes.cftInteger)
        .Add("cpd_type")
        .Add("web_publish")
        .Item(CpdCycleTypeFields.CpdCycleType).PrimaryKey = True
      End With
    End Sub
    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields.Item(CpdCycleTypeFields.WebPublish).Value = "N"
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cct"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "cpd_cycle_types"
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
    Public ReadOnly Property CpdCycleTypeCode() As String
      Get
        Return mvClassFields(CpdCycleTypeFields.CpdCycleType).Value
      End Get
    End Property
    Public ReadOnly Property CpdCycleTypeDesc() As String
      Get
        Return mvClassFields(CpdCycleTypeFields.CpdCycleTypeDesc).Value
      End Get
    End Property
    Public ReadOnly Property StartMonth() As Nullable(Of Integer)
      Get
        Dim vStartMonth As Nullable(Of Integer)
        Dim vInteger As Integer
        If Integer.TryParse(mvClassFields(CpdCycleTypeFields.StartMonth).Value, vInteger) Then
          vStartMonth = vInteger
        End If
        Return vStartMonth
      End Get
    End Property
    Public ReadOnly Property EndMonth() As Nullable(Of Integer)
      Get
        Dim vEndMonth As Nullable(Of Integer)
        Dim vInteger As Integer
        If Integer.TryParse(mvClassFields(CpdCycleTypeFields.EndMonth).Value, vInteger) Then
          vEndMonth = vInteger
        End If
        Return vEndMonth
      End Get
    End Property
    Public ReadOnly Property DefaultDuration() As Integer
      Get
        Return mvClassFields(CpdCycleTypeFields.DefaultDuration).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebPublish() As Boolean
      Get
        Return mvClassFields(CpdCycleTypeFields.WebPublish).Bool
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CpdCycleTypeFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CpdCycleTypeFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property CpdType() As String
      Get
        Return mvClassFields(CpdCycleTypeFields.CpdType).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Function EndDate(ByVal pStartDate As String) As String
      If DefaultDuration > 0 Then
        Return CDate(pStartDate).AddYears(DefaultDuration).AddDays(-1).ToString(CAREDateFormat)
      End If
      Return ""
    End Function

    ''' <summary>Is this a fixed CPD Cycle Type or a flexible CPD Cycle Type?</summary>
    ''' <returns>True if this is a fixed CPD Cycle Type, otherwise False.</returns>
    Public Function IsFixedCPDCycleType() As Boolean
      Return (StartMonth.HasValue = True AndAlso EndMonth.HasValue = True)
    End Function

#End Region
  End Class
End Namespace
