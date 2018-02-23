Imports System.Linq

Namespace Access

  ''' <summary>
  ''' A record of changes to an exam centre's details.
  ''' </summary>
  Public Class ExamCentreHistory
    Inherits CARERecord

    ''' <summary>
    ''' The datbase table fields are identified by this enumeration.
    ''' </summary>
    Public Enum ColumnId
      AllFields = 0
      ExamCentreHistoryId
      ExamCentreId
      ExamCentreDescriptionTimestamp
      ExamCentreDescription
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

    Private Shared mvColumnNames As New Dictionary(Of ColumnId, String)()

    ''' <summary>
    ''' Initializes the <see cref="ExamCentreHistory"/> class.
    ''' </summary>
    Shared Sub New()
      mvColumnNames.Add(ColumnId.ExamCentreHistoryId, "exam_centre_history_id")
      mvColumnNames.Add(ColumnId.ExamCentreId, "exam_centre_id")
      mvColumnNames.Add(ColumnId.ExamCentreDescriptionTimestamp, "exam_centre_desc_timestamp")
      mvColumnNames.Add(ColumnId.ExamCentreDescription, "exam_centre_description")
      mvColumnNames.Add(ColumnId.CreatedBy, "created_by")
      mvColumnNames.Add(ColumnId.CreatedOn, "created_on")
      mvColumnNames.Add(ColumnId.AmendedBy, "amended_by")
      mvColumnNames.Add(ColumnId.AmendedOn, "amended_on")
    End Sub

    ''' <summary>
    ''' Creates a new history record.
    ''' </summary>
    ''' <param name="pEnv">The application environment.</param>
    ''' <param name="pId">The exam centre unique identifier.</param>
    ''' <param name="pDescription">The exam centre description being replaced.</param>
    ''' <returns>A new <see cref="ExamCentreHistory"/> instance for the data provided.</returns>
    ''' <exception cref="System.ArgumentException">No exam centre exists with that ID.</exception>
    Public Shared Function CreateInstance(pEnv As CDBEnvironment, pId As Integer, pDescription As String) As ExamCentreHistory
      Dim vNewInstance As New ExamCentreHistory(pEnv)
      If New SQLStatement(pEnv.Connection, "Count(exam_centre_id)", "exam_centres", New CDBFields({New CDBField("exam_centre_id", pId)})).GetIntegerValue <> 0 Then
        vNewInstance.Init()
        vNewInstance.ExamCentreId = pId
        vNewInstance.Timestamp = Date.Now
        vNewInstance.Description = pDescription
      Else
        Throw New ArgumentException("No exam centre exists with that ID.")
      End If
      Return vNewInstance
    End Function

    ''' <summary>
    ''' Gets an existing exam centre history record.
    ''' </summary>
    ''' <param name="pEnv">The application environment.</param>
    ''' <param name="pExamCentreId">The application exam centre unique identifier.</param>
    ''' <param name="pExamCentreDescriptionTimestamp">The timestamp of the change required.</param>
    ''' <returns>>A <see cref="ExamCentreHistory"/> instance for the record requested.</returns>
    ''' <exception cref="System.ArgumentException">Requested row does not exist.</exception>
    Public Shared Function GetInstance(pEnv As CDBEnvironment, pExamCentreId As Integer, pExamCentreDescriptionTimestamp As Date) As ExamCentreHistory
      Dim vNewInstance As New ExamCentreHistory(pEnv)
      vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ColumnId.ExamCentreId).Name,
                                                                  pExamCentreId),
                                                     New CDBField(vNewInstance.mvClassFields(ColumnId.ExamCentreId).Name,
                                                                  CDBField.FieldTypes.cftTime,
                                                                  pExamCentreDescriptionTimestamp.ToString(CAREDateTimeFormat))}))
      If Not vNewInstance.Existing Then
        Throw New ArgumentException("Requested row does not exist.")
      End If
      Return vNewInstance
    End Function

    ''' <summary>
    ''' Get a history of the changes for an exam centre.
    ''' </summary>
    ''' <param name="pEnv">The application environemnt.</param>
    ''' <param name="pId">The exam centre unique identifier.</param>
    ''' <returns>A list of <see cref="ExamCentreHistory"/> instances containing  a history of the changes to the exam centre requested.</returns>
    Public Shared Function GetById(pEnv As CDBEnvironment, pId As Integer) As IList(Of ExamCentreHistory)
      Dim vList As New List(Of ExamCentreHistory)
      For Each pRow As DataRow In New SQLStatement(pEnv.Connection,
                                                   ColumnNameList.AsCommaSeperated,
                                                   Table,
                                                   New CDBField(ColumnName(ColumnId.ExamCentreId), pId),
                                                   ColumnName(ColumnId.ExamCentreDescriptionTimestamp)).GetDataTable.Rows
        vList.Add(New ExamCentreHistory(pEnv, pRow))
      Next pRow
      Return vList.AsReadOnly
    End Function

    ''' <summary>
    ''' Gets the name of the column.
    ''' </summary>
    ''' <value>
    ''' The name of the column.
    ''' </value>
    Public Shared ReadOnly Property ColumnName(pId As ColumnId) As String
      Get
        Return mvColumnNames(pId)
      End Get
    End Property

    ''' <summary>
    ''' Gets the aliased name of the column.
    ''' </summary>
    ''' <value>
    ''' The name of the column with the standard alias added.
    ''' </value>
    Public Shared ReadOnly Property AliasedColumnName(pId As ColumnId) As String
      Get
        Return ShortName & "." & mvColumnNames(pId)
      End Get
    End Property

    ''' <summary>
    ''' Gets the list of column names.
    ''' </summary>
    ''' <returns>A read only list containing the column names.</returns>
    Public Shared ReadOnly Property ColumnNameList As IList(Of String)
      Get
        Return (New List(Of String)(mvColumnNames.Values)).AsReadOnly
      End Get
    End Property

    ''' <summary>
    ''' Gets the list of column names aliased by the standard alias.
    ''' </summary>
    ''' <param name="pEnv">The application env.</param>
    ''' <returns></returns>
    Public Shared Function GetAliasedColumnNameList(pEnv As CDBEnvironment) As IList(Of String)
      Return (New List(Of String)(From vField As String In ColumnNameList
                                 Select ShortName & "." & vField)).AsReadOnly
    End Function

    ''' <summary>
    ''' Gets the name of the database table.
    ''' </summary>
    ''' <value>
    ''' The database table name.
    ''' </value>
    Public Shared ReadOnly Property Table As String
      Get
        Return "exam_centre_history"
      End Get
    End Property

    ''' <summary>
    ''' Gets the database table alias.
    ''' </summary>
    ''' <value>
    ''' The database table alias.
    ''' </value>
    Public Shared ReadOnly Property ShortName As String
      Get
        Return "ech"
      End Get
    End Property

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
    ''' Creates an instance of the <see cref="ExamCentreHistory"/> class containing data from the supplied <see cref="DataRow"/>.  
    ''' This is only used internally.  Applications must use the <see cref="CreateInstance" /> or <see cref="GetInstance" /> methods as appropriate.
    ''' </summary>
    ''' <param name="pEnv">The application environment.</param>
    ''' <param name="pRow">The data row.</param>
    Private Sub New(ByVal pEnv As CDBEnvironment, pRow As DataRow)
      MyBase.New(pEnv)
      Me.Init()
      Me.ExamCentreId = DirectCast(pRow(mvClassFields(ColumnId.ExamCentreId).Name), Integer)
      Me.Timestamp = DirectCast(pRow(mvClassFields(ColumnId.ExamCentreDescriptionTimestamp).Name), Date)
      Me.Description = DirectCast(pRow(mvClassFields(ColumnId.ExamCentreDescription).Name), String)
      Me.CreatedBy = DirectCast(pRow(mvClassFields(ColumnId.CreatedBy).Name), String)
      Me.CreatedOn = DirectCast(pRow(mvClassFields(ColumnId.CreatedOn).Name), Date)
      Me.AmendedBy = DirectCast(pRow(mvClassFields(ColumnId.AmendedBy).Name), String)
      Me.AmendedOn = DirectCast(pRow(mvClassFields(ColumnId.AmendedOn).Name), Date)
    End Sub

    ''' <summary>
    ''' Adds the fields.
    ''' </summary>
    Protected Overrides Sub AddFields()
      mvClassFields.Add(mvColumnNames(ColumnId.ExamCentreHistoryId), CDBField.FieldTypes.cftInteger)
      mvClassFields.Add(mvColumnNames(ColumnId.ExamCentreId), CDBField.FieldTypes.cftInteger)
      mvClassFields.Add(mvColumnNames(ColumnId.ExamCentreDescriptionTimestamp), CDBField.FieldTypes.cftTime)
      mvClassFields.Add(mvColumnNames(ColumnId.ExamCentreDescription))
      mvClassFields.Add(mvColumnNames(ColumnId.CreatedBy))
      mvClassFields.Add(mvColumnNames(ColumnId.CreatedOn), CDBField.FieldTypes.cftDate)

      mvClassFields.Item(ColumnId.ExamCentreHistoryId).PrimaryKey = True
      mvClassFields.Item(ColumnId.ExamCentreHistoryId).PrefixRequired = True
      mvClassFields.Item(ColumnId.ExamCentreId).PrefixRequired = True
      mvClassFields.Item(ColumnId.ExamCentreDescriptionTimestamp).PrefixRequired = True
      mvClassFields.Item(ColumnId.ExamCentreDescription).PrefixRequired = True
      mvClassFields.Item(ColumnId.CreatedBy).PrefixRequired = True
      mvClassFields.Item(ColumnId.CreatedOn).PrefixRequired = True
      mvClassFields.SetControlNumberField(ColumnId.ExamCentreHistoryId, "ECH")
    End Sub

    ''' <summary>
    ''' Gets the database table alias.
    ''' </summary>
    ''' <value>
    ''' The database table alias.
    ''' </value>
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return ShortName
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
        Return Table
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

    ''' <summary>
    ''' Gets the exam centre unique identifier.
    ''' </summary>
    ''' <value>
    ''' The unique identifier.
    ''' </value>
    Public Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ColumnId.ExamCentreId).IntegerValue
      End Get
      Private Set(value As Integer)
        mvClassFields(ColumnId.ExamCentreId).IntegerValue = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the timestamp of the change.
    ''' </summary>
    ''' <value>
    ''' The timestamp.
    ''' </value>
    Public Property Timestamp() As Date
      Get
        Return Date.Parse(mvClassFields(ColumnId.ExamCentreDescriptionTimestamp).Value)
      End Get
      Private Set(value As Date)
        mvClassFields(ColumnId.ExamCentreDescriptionTimestamp).Value = value.ToString(CAREDateTimeFormat)
      End Set
    End Property

    ''' <summary>
    ''' Gets the description in use up to this point in time.
    ''' </summary>
    ''' <value>
    ''' The description.
    ''' </value>
    Public Property Description() As String
      Get
        Return mvClassFields(ColumnId.ExamCentreDescription).Value
      End Get
      Private Set(value As String)
        mvClassFields(ColumnId.ExamCentreDescription).Value = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the log name of the user that created the record.
    ''' </summary>
    ''' <value>
    ''' The creating user's log name.
    ''' </value>
    Public Property CreatedBy() As String
      Get
        Return mvClassFields(ColumnId.CreatedBy).Value
      End Get
      Private Set(value As String)
        mvClassFields(ColumnId.CreatedBy).Value = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the date that this record was created on.
    ''' </summary>
    ''' <value>
    ''' The creation date.
    ''' </value>
    Public Property CreatedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ColumnId.CreatedOn).Value)
      End Get
      Private Set(value As Date)
        mvClassFields(ColumnId.CreatedOn).Value = value.ToString(CAREDateFormat)
      End Set
    End Property

    ''' <summary>
    ''' Gets the log name of the user that amended the record.
    ''' </summary>
    ''' <value>
    ''' The amending user's log name.
    ''' </value>
    Public Property AmendedBy() As String
      Get
        Return mvClassFields(ColumnId.AmendedBy).Value
      End Get
      Private Set(value As String)
        mvClassFields(ColumnId.AmendedBy).Value = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the date that this record was last amended on.
    ''' </summary>
    ''' <value>
    ''' The last amended date.
    ''' </value>
    Public Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ColumnId.AmendedOn).Value)
      End Get
      Private Set(value As Date)
        mvClassFields(ColumnId.AmendedOn).Value = value.ToString(CAREDateFormat)
      End Set
    End Property

  End Class

End Namespace