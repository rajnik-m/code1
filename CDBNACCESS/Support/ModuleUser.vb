Namespace Access

  Public Class ModuleUser
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ModuleUserFields
      AllFields = 0
      ModuleCode
      Logname
      StartTime
      BuildNumber
      NamedUser
      Active
      AccessCount
      RefusedAccess
      LastUpdatedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("module").SpecialColumn = True
        .Add("logname")
        .Add("start_time", CDBField.FieldTypes.cftTime)
        .Add("build_number", CDBField.FieldTypes.cftInteger)
        .Add("named_user")
        .Add("active")
        .Add("access_count", CDBField.FieldTypes.cftLong)
        .Add("refused_access", CDBField.FieldTypes.cftLong)
        .Add("last_updated_on", CDBField.FieldTypes.cftTime)

        .Item(ModuleUserFields.ModuleCode).PrimaryKey = True
        .Item(ModuleUserFields.ModuleCode).SpecialColumn = True
        .Item(ModuleUserFields.Logname).PrimaryKey = True
        .Item(ModuleUserFields.StartTime).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "smu"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "sys_module_users"
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
    Public ReadOnly Property ModuleCode() As String
      Get
        Return mvClassFields(ModuleUserFields.ModuleCode).Value
      End Get
    End Property
    Public ReadOnly Property Logname() As String
      Get
        Return mvClassFields(ModuleUserFields.Logname).Value
      End Get
    End Property
    Public ReadOnly Property StartTime() As String
      Get
        Return mvClassFields(ModuleUserFields.StartTime).Value
      End Get
    End Property
    Public ReadOnly Property BuildNumber() As Integer
      Get
        Return mvClassFields(ModuleUserFields.BuildNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property NamedUser() As Boolean
      Get
        Return mvClassFields(ModuleUserFields.NamedUser).Bool
      End Get
    End Property
    Public ReadOnly Property Active() As Boolean
      Get
        Return mvClassFields(ModuleUserFields.Active).Bool
      End Get
    End Property
    Public ReadOnly Property AccessCount() As Integer
      Get
        Return mvClassFields(ModuleUserFields.AccessCount).IntegerValue
      End Get
    End Property
    Public ReadOnly Property RefusedAccess() As Integer
      Get
        Return mvClassFields(ModuleUserFields.RefusedAccess).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LastUpdatedOn() As String
      Get
        Return mvClassFields(ModuleUserFields.LastUpdatedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Public Overloads Sub Init(ByVal pModule As String, ByVal pLogname As String, ByVal pStartTime As String)
      Dim vRecordSet As CDBRecordSet
      InitClassFields()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("module", pModule)
      vWhereFields.Add("logname", pLogname)
      vWhereFields.Add("start_time", CDBField.FieldTypes.cftTime, pStartTime)
      vRecordSet = New SQLStatement(mvEnv.Connection, GetRecordSetFields(), mvClassFields.TableNameAndAlias, vWhereFields).GetRecordSet
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(vRecordSet)
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields.Item(ModuleUserFields.Logname).Value = mvEnv.User.Logname
      mvClassFields.Item(ModuleUserFields.NamedUser).Bool = False
      mvClassFields.Item(ModuleUserFields.Active).Bool = True
      mvClassFields.Item(ModuleUserFields.AccessCount).IntegerValue = 1
      mvClassFields.Item(ModuleUserFields.RefusedAccess).IntegerValue = 0
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      mvClassFields.Item(ModuleUserFields.LastUpdatedOn).Value = TodaysDateAndTime()
    End Sub

    Public Overloads Sub Create(ByVal pModuleCode As String, ByVal pStartTime As Date, ByVal pBuildNumber As Integer, ByVal pRefusedAccess As Boolean)
      With mvClassFields
        .Item(ModuleUserFields.ModuleCode).Value = pModuleCode
        .Item(ModuleUserFields.StartTime).Value = pStartTime.ToString
        .Item(ModuleUserFields.BuildNumber).IntegerValue = pBuildNumber
        If pRefusedAccess Then
          .Item(ModuleUserFields.RefusedAccess).IntegerValue = 1
          .Item(ModuleUserFields.Active).Bool = False
          .Item(ModuleUserFields.AccessCount).IntegerValue = 0
        End If
      End With
    End Sub

    Public Sub SetRefusedAccess(ByVal pBuildNumber As Integer)
      With mvClassFields
        .Item(ModuleUserFields.BuildNumber).IntegerValue = pBuildNumber
        .Item(ModuleUserFields.RefusedAccess).IntegerValue = RefusedAccess + 1
      End With
    End Sub

    Public Sub SetActive(ByVal pStartTime As Date, ByVal pBuildNumber As Integer)
      With mvClassFields
        .Item(ModuleUserFields.StartTime).Value = pStartTime.ToString
        .Item(ModuleUserFields.BuildNumber).IntegerValue = pBuildNumber
        .Item(ModuleUserFields.Active).Bool = True
        .Item(ModuleUserFields.AccessCount).IntegerValue = AccessCount + 1
      End With
    End Sub

    Public Sub UpdateNamedUser(ByVal pNamedUser As Boolean)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      vUpdateFields.Add("named_user", BooleanString(pNamedUser))
      vWhereFields.Add("module", ModuleCode).SpecialColumn = True
      vWhereFields.Add("logname", Logname)
      If pNamedUser Then
        vWhereFields.Add("start_time", CDBField.FieldTypes.cftTime, StartTime)
      Else
        vWhereFields.Add("named_user", CDBField.FieldTypes.cftCharacter, "Y")
      End If
      mvEnv.Connection.UpdateRecords(mvClassFields.DatabaseTableName, vUpdateFields, vWhereFields)
    End Sub
#End Region

  End Class
End Namespace
