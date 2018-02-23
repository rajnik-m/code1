Namespace Access

  ''' <summary>
  ''' A record of changes to an exam centre's details.
  ''' </summary>
  Public Class ExamUnitStudyMode
    Inherits CARERecord

    ''' <summary>
    ''' The datbase table fields are identified by this enumeration.
    ''' </summary>
    Public Enum ExamUnitStudyModeColumns
      AllFields = 0
      ExamUnitStudyModeId
      ExamUnitLinkId
      StudyMode
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment, pExamUnitLinkId As Integer, pStudyMode As String) As ExamUnitStudyMode
      Dim vNewInstance As New ExamUnitStudyMode(pEnv)
      If New SQLStatement(pEnv.Connection, "Count(exam_unit_link_id)", "exam_unit_links", New CDBFields({New CDBField("exam_unit_link_id", pExamUnitLinkId)})).GetIntegerValue = 0 Then
        Throw New ArgumentException("No exam unit link exists with that ID.")
      ElseIf New SQLStatement(pEnv.Connection, "Count(study_mode)", "study_modes", New CDBFields({New CDBField("study_mode", pStudyMode)})).GetIntegerValue = 0 Then
        Throw New ArgumentException("No study mode exists with that code.")
      Else
        vNewInstance.ExamUnitLinkId = pExamUnitLinkId
        vNewInstance.StudyMode = pStudyMode
      End If
      Return vNewInstance
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pExamUnitLinkId As Integer, pStudyMode As String) As ExamUnitStudyMode
      Dim vNewInstance As New ExamUnitStudyMode(pEnv)
      vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamUnitStudyModeColumns.ExamUnitLinkId).Name,
                                                                  pExamUnitLinkId),
                                                     New CDBField(vNewInstance.mvClassFields(ExamUnitStudyModeColumns.StudyMode).Name,
                                                                  pStudyMode)}))
      If Not vNewInstance.Existing Then
        Throw New ArgumentException("Requested row does not exist.")
      End If
      Return vNewInstance
    End Function

    ''' <summary>
    ''' Creates an empty instance of the <see cref="ExamCentreHistory"/> class.  This is only used internally.  Applications 
    ''' must use the <see cref="CreateInstance" /> or <see cref="GetInstance" /> methods as appropriate.
    ''' </summary>
    ''' <param name="pEnv">The application environment.</param>
    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    ''' <summary>
    ''' Adds the fields.
    ''' </summary>
    Protected Overrides Sub AddFields()
      mvClassFields.Add("exam_unit_study_mode_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("study_mode")

      mvClassFields.Item(ExamUnitStudyModeColumns.ExamUnitStudyModeId).PrimaryKey = True
      mvClassFields.Item(ExamUnitStudyModeColumns.ExamUnitLinkId).PrefixRequired = True
      mvClassFields.Item(ExamUnitStudyModeColumns.StudyMode).PrefixRequired = True
      mvClassFields.SetControlNumberField(ExamUnitStudyModeColumns.ExamUnitStudyModeId, "XSM")
    End Sub

    ''' <summary>
    ''' Gets the database table alias.
    ''' </summary>
    ''' <value>
    ''' The database table alias.
    ''' </value>
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "xusm"
      End Get
    End Property

    ''' <summary>
    ''' Gets the name of the database table.
    ''' </summary>
    ''' <value>
    ''' The database table name.
    ''' </value>
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_unit_study_modes"
      End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether [supports amended configuration and by].
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if [supports amended configuration and by]; otherwise, <c>false</c>.
    ''' </value>
    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Public ReadOnly Property ExamUnitStudyModeId() As Integer
      Get
        Return mvClassFields(ExamUnitStudyModeColumns.ExamUnitStudyModeId).IntegerValue
      End Get
    End Property

    Public Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamUnitStudyModeColumns.ExamUnitLinkId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ExamUnitStudyModeColumns.ExamUnitLinkId).IntegerValue = value
      End Set
    End Property

    Public Property StudyMode() As String
      Get
        Return mvClassFields(ExamUnitStudyModeColumns.StudyMode).Value
      End Get
      Private Set(value As String)
        mvClassFields(ExamUnitStudyModeColumns.StudyMode).Value = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the log name of the user that amended the record.
    ''' </summary>
    ''' <value>
    ''' The amending user's log name.
    ''' </value>
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamUnitStudyModeColumns.AmendedBy).Value
      End Get
    End Property

    ''' <summary>
    ''' Gets the date that this record was last amended on.
    ''' </summary>
    ''' <value>
    ''' The last amended date.
    ''' </value>
    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ExamUnitStudyModeColumns.AmendedOn).Value)
      End Get
    End Property

  End Class

End Namespace