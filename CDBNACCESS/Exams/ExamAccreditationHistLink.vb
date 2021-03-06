Namespace Access

  Public Class ExamAccreditationHistLink
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ExamAccreditationHistLinkFields
      AllFields = 0
      AccreditationHistoryLinkId
      ExamUnitLinkId
      ExamCentreId
      ExamCentreUnitId
      AccreditationId
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("accreditation_history_link_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("accreditation_id", CDBField.FieldTypes.cftInteger)
        .SetControlNumberField(ExamAccreditationHistLinkFields.AccreditationHistoryLinkId, "AHL")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "eahl"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_accreditation_hist_links"
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
    Public ReadOnly Property AccreditationHistoryLinkId() As Integer
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.AccreditationHistoryLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreUnitId() As Integer
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.ExamCentreUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AccreditationId() As Integer
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.AccreditationId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamAccreditationHistLinkFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace
