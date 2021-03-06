Namespace Access

  Public Class EventSessionCpd
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum SessionCpdFields
      AllFields = 0
      EventSessionCpdNumber
      EventNumber
      SessionNumber
      CpdCategoryType
      CpdCategory
      CpdYear
      CpdPoints
      CpdPoints2
      CpdItemType
      CpdOutcome
      CpdApprovalStatus
      CpdDateApproved
      CpdAwardingBody
      WebPublish
      CpdNotes
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("event_session_cpd_number", CDBField.FieldTypes.cftInteger)
        .Add("event_number", CDBField.FieldTypes.cftInteger)
        .Add("session_number", CDBField.FieldTypes.cftInteger)
        .Add("cpd_category_type")
        .Add("cpd_category")
        .Add("cpd_year", CDBField.FieldTypes.cftInteger)
        .Add("cpd_points", CDBField.FieldTypes.cftNumeric)
        .Add("cpd_points_2", CDBField.FieldTypes.cftNumeric)
        .Add("cpd_item_type")
        .Add("cpd_outcome")
        .Add("cpd_approval_status")
        .Add("cpd_date_approved", CDBField.FieldTypes.cftDate)
        .Add("cpd_awarding_body")
        .Add("web_publish")
        .Add("cpd_notes", CDBField.FieldTypes.cftMemo)

        .Item(SessionCpdFields.EventSessionCpdNumber).PrimaryKey = True
        .Item(SessionCpdFields.EventSessionCpdNumber).PrefixRequired = True

        .Item(SessionCpdFields.EventNumber).PrefixRequired = True
        .Item(SessionCpdFields.SessionNumber).PrefixRequired = True
        .Item(SessionCpdFields.CpdCategoryType).PrefixRequired = True
        .Item(SessionCpdFields.CpdCategory).PrefixRequired = True
        .Item(SessionCpdFields.CpdYear).PrefixRequired = True
        .Item(SessionCpdFields.CpdPoints).PrefixRequired = True
        .Item(SessionCpdFields.CpdPoints2).PrefixRequired = True
        .Item(SessionCpdFields.CpdItemType).PrefixRequired = True
        .Item(SessionCpdFields.CpdOutcome).PrefixRequired = True
        .Item(SessionCpdFields.CpdApprovalStatus).PrefixRequired = True
        .Item(SessionCpdFields.CpdDateApproved).PrefixRequired = True
        .Item(SessionCpdFields.CpdAwardingBody).PrefixRequired = True
        .Item(SessionCpdFields.WebPublish).PrefixRequired = True
        .Item(SessionCpdFields.CpdNotes).PrefixRequired = True

        .SetControlNumberField(SessionCpdFields.EventSessionCpdNumber, "ESC")
      End With

      AddDeleteCheckItem("contact_cpd_points", "event_session_cpd_number", "Contact CPD Points")
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "sc"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "session_cpd"
      End Get
    End Property

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields(SessionCpdFields.CpdPoints).DoubleValue = 0
      mvClassFields(SessionCpdFields.CpdPoints2).DoubleValue = 0
      mvClassFields(SessionCpdFields.WebPublish).Value = "N"
    End Sub

    Protected Overrides Sub PreValidateCreateParameters(pParameterList As CDBParameters)
      MyBase.PreValidateCreateParameters(pParameterList)
      ValidateEventAndSession(pParameterList)
      PreValidatePoints(pParameterList)
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(pParameterList As CDBParameters)
      MyBase.PreValidateUpdateParameters(pParameterList)
      PreValidatePoints(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(pParameterList As CDBParameters)
      MyBase.PostValidateCreateParameters(pParameterList)
      If pParameterList.ParameterExists("IgnoreLookupValidation").Bool = False Then ValidateData()
      If pParameterList.ContainsKey("AmendedBy") Then
        mvClassFields.Item(SessionCpdFields.AmendedBy).Value = pParameterList("AmendedBy").Value
        mvClassFields.Item(SessionCpdFields.AmendedOn).Value = TodaysDate()
        mvOverrideAmended = True
      End If
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      If mvClassFields.Item(SessionCpdFields.EventNumber).ValueChanged = True OrElse mvClassFields.Item(SessionCpdFields.SessionNumber).ValueChanged = True Then
        RaiseError(DataAccessErrors.daeCPDCannotUpdateEventOrSessionNumbers)
      End If
      ValidateData()
    End Sub

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property EventSessionCpdNumber() As Integer
      Get
        Return mvClassFields(SessionCpdFields.EventSessionCpdNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property EventNumber() As Integer
      Get
        Return mvClassFields(SessionCpdFields.EventNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property SessionNumber() As Integer
      Get
        Return mvClassFields(SessionCpdFields.SessionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CpdCategoryType() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdCategoryType).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategory() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdCategory).Value
      End Get
    End Property
    Public ReadOnly Property CpdYear() As Integer
      Get
        Return mvClassFields(SessionCpdFields.CpdYear).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CpdPoints() As Double
      Get
        Return mvClassFields(SessionCpdFields.CpdPoints).DoubleValue
      End Get
    End Property
    Public ReadOnly Property CpdPoints2() As Double
      Get
        Return mvClassFields(SessionCpdFields.CpdPoints2).DoubleValue
      End Get
    End Property
    Public ReadOnly Property CpdItemType() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdItemType).Value
      End Get
    End Property
    Public ReadOnly Property CpdOutcome() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdOutcome).Value
      End Get
    End Property
    Public ReadOnly Property CpdApprovalStatus() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdApprovalStatus).Value
      End Get
    End Property
    Public ReadOnly Property CpdDateApproved() As Nullable(Of Date)
      Get
        Dim vDateApproved As Nullable(Of Date)
        If IsDate(mvClassFields(SessionCpdFields.CpdDateApproved).Value) Then vDateApproved = Date.Parse(mvClassFields(SessionCpdFields.CpdDateApproved).Value)
        Return vDateApproved
      End Get
    End Property
    Public ReadOnly Property CpdAwardingBody() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdAwardingBody).Value
      End Get
    End Property
    Public ReadOnly Property WebPublish() As String
      Get
        Return mvClassFields(SessionCpdFields.WebPublish).Value
      End Get
    End Property
    Public ReadOnly Property CpdNotes() As String
      Get
        Return mvClassFields(SessionCpdFields.CpdNotes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(SessionCpdFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(SessionCpdFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    ''' <summary>Validate CPD Points and CPD Points 2 to ensure that they are only numeric when configuration allows.</summary>
    ''' <remarks>Uses the 'cpd_points_allow_numeric' configuration option.</remarks>
    Private Sub PreValidatePoints(ByVal pParams As CDBParameters)
      If mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
        'Make sure the values are whole numbers only
        If pParams.ParameterExists("CpdPoints").Value.Length > 0 Then pParams("CpdPoints").Value = pParams("CpdPoints").IntegerValue.ToString
        If pParams.ParameterExists("CpdPoints2").Value.Length > 0 Then pParams("CpdPoints2").Value = pParams("CpdPoints2").IntegerValue.ToString
      End If
    End Sub

    ''' <summary>Validate the data before saving.</summary>
    ''' <remarks>Records created by migrating existing data will not call this validation.</remarks>
    Private Sub ValidateData()
      If mvExisting = False OrElse _
        (mvClassFields.Item(SessionCpdFields.CpdCategoryType).ValueChanged = True OrElse mvClassFields.Item(SessionCpdFields.CpdCategory).ValueChanged = True) Then
        Dim vCPDCategoryType As New CpdCategoryType(mvEnv)
        vCPDCategoryType.Init(CpdCategoryType)
        If vCPDCategoryType.Existing = False Then
          RaiseError(DataAccessErrors.daeParameterValueInvalid, "CpdCategoryType")
        End If

        Dim vCPDCategory As New CpdCategory(mvEnv)
        vCPDCategory.Init(CpdCategory)
        If vCPDCategory.Existing = False Then
          RaiseError(DataAccessErrors.daeParameterValueInvalid, "CpdCategory")
        ElseIf vCPDCategory.CpdCategoryType.Equals(vCPDCategoryType.CpdCategoryTypeCode, StringComparison.InvariantCultureIgnoreCase) = False Then
          RaiseError(DataAccessErrors.daeCPDCategoryAndCategoryTypeInvalid)
        Else
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "start_date", "sessions", New CDBFields(New CDBField("session_number", SessionNumber)))
          Dim vStartDate As String = vSQLStatement.GetValue()
          If IsDate(vStartDate) = False Then vStartDate = TodaysDate()
          If vCPDCategory.IsValidForPointsEntry(Date.Parse(vStartDate)) = False Then
            RaiseError(DataAccessErrors.daeCPDCategoryNotApprovedOrInvalid, vCPDCategory.CpdCategoryCode)
          End If
        End If

        If mvEnv.GetConfigOption("cpd_unique_categories", True) Then
          'Data is unique for EventNumber, SessionNumber, CPDCategoryType & CPDCategory
          'Setting the unique fields also gets the Save to check this does not exist
          mvClassFields.SetUniqueField(SessionCpdFields.EventNumber)
          mvClassFields.SetUniqueField(SessionCpdFields.SessionNumber)
          mvClassFields.SetUniqueField(SessionCpdFields.CpdCategoryType)
          mvClassFields.SetUniqueField(SessionCpdFields.CpdCategory)

          Dim vWhereFields As New CDBFields(New CDBField("event_number", EventNumber))
          vWhereFields.Add("session_number", SessionNumber)
          vWhereFields.Add("cpd_category_type", CpdCategoryType)
          vWhereFields.Add("cpd_category", CpdCategory)
          If mvExisting Then vWhereFields.Add("event_session_cpd_number", CDBField.FieldTypes.cftInteger, EventSessionCpdNumber.ToString, CDBField.FieldWhereOperators.fwoNotEqual)
          If mvEnv.Connection.GetCount("session_cpd", vWhereFields) > 0 Then
            RaiseError(DataAccessErrors.daeCPDSessionCategoryTypeCategoryNotUnique)
          End If
        End If
      End If

      If mvExisting = False OrElse _
      (mvClassFields.Item(SessionCpdFields.CpdPoints).ValueChanged = True OrElse mvClassFields.Item(SessionCpdFields.CpdPoints2).ValueChanged = True) Then
        If mvClassFields.Item(SessionCpdFields.CpdPoints).DoubleValue + mvClassFields.Item(SessionCpdFields.CpdPoints2).DoubleValue = 0 Then
          RaiseError(DataAccessErrors.daeCPDPointsTotalCannotBeZero)
        End If
      End If

      If mvExisting = False OrElse mvClassFields.Item(SessionCpdFields.CpdApprovalStatus).ValueChanged = True Then
        Dim vCPDApprovalStatus As New CpdApprovalStatus(mvEnv)
        vCPDApprovalStatus.Init(CpdApprovalStatus)
        If vCPDApprovalStatus.Existing = False Then
          RaiseError(DataAccessErrors.daeParameterValueInvalid, "CpdApprovalStatus")
        Else
          If vCPDApprovalStatus.CpdApprovalDateRequired = True AndAlso CpdDateApproved.HasValue = False Then
            RaiseError(DataAccessErrors.daeCPDApprovalDateRequired)
          End If
        End If
      End If

    End Sub

    ''' <summary>When creating a new record, validate the Event and Session.</summary>
    Private Sub ValidateEventAndSession(ByVal pParams As CDBParameters)
      If mvExisting = False Then
        Dim vEvent As New CDBEvent(mvEnv)
        vEvent.Init(pParams("EventNumber").IntegerValue)
        If vEvent.Existing = False Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "EventNumber")

        Dim vSession As New EventSession()
        vSession.Init(mvEnv, pParams("SessionNumber").IntegerValue)
        If vSession.Existing = False Then
          RaiseError(DataAccessErrors.daeParameterValueInvalid, "SessionNumber")
        ElseIf vSession.EventNumber.Equals(vEvent.EventNumber) = False Then
          RaiseError(DataAccessErrors.daeParameterValueInvalid, "EventNumber")
        End If

        If vEvent.MultiSession = True Then
          'Mult-session Event so Session cannot be the base Session
          If vSession.SessionNumber.Equals(vEvent.LowestSessionNumber) = True OrElse vSession.SessionType.Equals("0") = True Then
            RaiseError(DataAccessErrors.daeCPDSessionCannotBeBaseSession)
          End If
        Else
          'Single-session Event so Session must be the base Session
          If vSession.SessionType.Equals("0") = False Then
            RaiseError(DataAccessErrors.daeCPDSessionMustBeBaseSession)
          End If
        End If
      End If
    End Sub

#End Region

  End Class
End Namespace
