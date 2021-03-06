Namespace Access

  Public Class ExamSessionCentre
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ExamSessionCentreFields
      AllFields = 0
      ExamSessionCentreId
      ExamCentreId
      ExamSessionId
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("exam_session_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_session_id", CDBField.FieldTypes.cftInteger)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamSessionCentreFields.ExamSessionCentreId).PrimaryKey = True
        .Item(ExamSessionCentreFields.ExamSessionCentreId).PrefixRequired = True
        .SetControlNumberField(ExamSessionCentreFields.ExamSessionCentreId, "XST")

        .Item(ExamSessionCentreFields.CreatedBy).PrefixRequired = True
        .Item(ExamSessionCentreFields.CreatedOn).PrefixRequired = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "esc"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_session_centres"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'IRecordCreate
'--------------------------------------------------
    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ExamSessionCentre(mvEnv)
    End Function
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ExamSessionCentreId() As Integer
      Get
        Return mvClassFields(ExamSessionCentreFields.ExamSessionCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ExamSessionCentreFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamSessionId() As Integer
      Get
        Return mvClassFields(ExamSessionCentreFields.ExamSessionId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamSessionCentreFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamSessionCentreFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamSessionCentreFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamSessionCentreFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace
