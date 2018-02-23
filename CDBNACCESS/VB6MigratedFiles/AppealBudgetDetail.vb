Namespace Access
  Public Class AppealBudgetDetail

    Public Enum AppealBudgetDetailRecordSetTypes 'These are bit values
      abdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum AppealBudgetDetailFields
      abdfAll = 0
      abdfAppealBudgetDetailsNumber
      abdfAppealBudgetNumber
      abdfSegment
      abdfReasonForDespatch
      abdfForecastUnits
      abdfBudgetedCosts
      abdfBudgetedIncome
      abdfAmendedBy
      abdfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "appeal_budget_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("appeal_budget_details_number", CDBField.FieldTypes.cftLong)
          .Add("appeal_budget_number", CDBField.FieldTypes.cftLong)
          .Add("segment")
          .Add("reason_for_despatch")
          .Add("forecast_units", CDBField.FieldTypes.cftLong)
          .Add("budgeted_costs", CDBField.FieldTypes.cftNumeric)
          .Add("budgeted_income", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields.Item(AppealBudgetDetailFields.abdfSegment).PrefixRequired = True
        mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedBy).PrefixRequired = True
        mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedOn).PrefixRequired = True
        mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As AppealBudgetDetailFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedBy).Value = mvEnv.User.Logname
      If (pField = AppealBudgetDetailFields.abdfAll Or pField = AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber) And mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber).IntegerValue = 0 Then
        mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber).Value = CStr(mvEnv.GetControlNumber("AD"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As AppealBudgetDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AppealBudgetDetailRecordSetTypes.abdrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "abd")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAppealBudgetDetailsNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pAppealBudgetDetailsNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AppealBudgetDetailRecordSetTypes.abdrtAll) & " FROM appeal_budget_details abd WHERE appeal_budget_details_number = " & pAppealBudgetDetailsNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, AppealBudgetDetailRecordSetTypes.abdrtAll)
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

    Public Sub InitFromDetails(ByVal pEnv As CDBEnvironment, ByRef pAppealBudgetNumber As Integer, ByRef pSegment As String, ByRef pReasonForDespatch As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      vWhereFields.Add("appeal_budget_number", CDBField.FieldTypes.cftLong, pAppealBudgetNumber)
      vWhereFields.Add("segment", CDBField.FieldTypes.cftCharacter, pSegment)
      vWhereFields.Add("reason_for_despatch", CDBField.FieldTypes.cftCharacter, pReasonForDespatch)
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AppealBudgetDetailRecordSetTypes.abdrtAll) & " FROM appeal_budget_details abd WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, AppealBudgetDetailRecordSetTypes.abdrtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AppealBudgetDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AppealBudgetDetailRecordSetTypes.abdrtAll) = AppealBudgetDetailRecordSetTypes.abdrtAll Then
          .SetItem(AppealBudgetDetailFields.abdfAppealBudgetNumber, vFields)
          .SetItem(AppealBudgetDetailFields.abdfSegment, vFields)
          .SetItem(AppealBudgetDetailFields.abdfReasonForDespatch, vFields)
          .SetItem(AppealBudgetDetailFields.abdfForecastUnits, vFields)
          .SetItem(AppealBudgetDetailFields.abdfBudgetedCosts, vFields)
          .SetItem(AppealBudgetDetailFields.abdfBudgetedIncome, vFields)
          .SetItem(AppealBudgetDetailFields.abdfAmendedBy, vFields)
          .SetItem(AppealBudgetDetailFields.abdfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(AppealBudgetDetailFields.abdfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByRef pAppealBudgetNumber As Integer, ByRef pSegment As String, ByRef pReasonForDespatch As String, ByRef pForecastUnits As String, ByRef pBudgetedCost As String, ByRef pBudgetedIncome As String)
      mvEnv = pEnv
      InitClassFields()
      SetValid(AppealBudgetDetailFields.abdfAll)
      mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetNumber).Value = CStr(pAppealBudgetNumber)
      mvClassFields.Item(AppealBudgetDetailFields.abdfSegment).Value = pSegment
      mvClassFields.Item(AppealBudgetDetailFields.abdfReasonForDespatch).Value = pReasonForDespatch
      If Len(pForecastUnits) > 0 Then mvClassFields.Item(AppealBudgetDetailFields.abdfForecastUnits).Value = pForecastUnits
      If Len(pBudgetedCost) > 0 Then mvClassFields.Item(AppealBudgetDetailFields.abdfBudgetedCosts).Value = pBudgetedCost
      If Len(pBudgetedIncome) > 0 Then mvClassFields.Item(AppealBudgetDetailFields.abdfBudgetedIncome).Value = pBudgetedIncome
      Save()
    End Sub

    Public Sub Update(ByRef pReasonForDespatch As String, ByRef pForecastUnits As String, ByRef pBudgetedCosts As String, ByRef pBudgetedIncome As String)
      With mvClassFields
        .Item(AppealBudgetDetailFields.abdfReasonForDespatch).Value = pReasonForDespatch
        .Item(AppealBudgetDetailFields.abdfForecastUnits).Value = pForecastUnits
        .Item(AppealBudgetDetailFields.abdfBudgetedCosts).Value = pBudgetedCosts
        .Item(AppealBudgetDetailFields.abdfBudgetedIncome).Value = pBudgetedIncome
      End With
    End Sub

    Public Sub Delete()
      If mvExisting Then mvEnv.Connection.DeleteRecords("appeal_budget_details", mvClassFields.WhereFields)
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
        AmendedBy = mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(AppealBudgetDetailFields.abdfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AppealBudgetDetailsNumber() As Integer
      Get
        AppealBudgetDetailsNumber = mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetDetailsNumber).IntegerValue
      End Get
    End Property

    Public Property AppealBudgetNumber() As Integer
      Get
        AppealBudgetNumber = mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(AppealBudgetDetailFields.abdfAppealBudgetNumber).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property BudgetedCosts() As String
      Get
        BudgetedCosts = mvClassFields.Item(AppealBudgetDetailFields.abdfBudgetedCosts).Value
      End Get
    End Property

    Public ReadOnly Property BudgetedIncome() As String
      Get
        BudgetedIncome = mvClassFields.Item(AppealBudgetDetailFields.abdfBudgetedIncome).Value
      End Get
    End Property

    Public ReadOnly Property ForecastUnits() As String
      Get
        ForecastUnits = mvClassFields.Item(AppealBudgetDetailFields.abdfForecastUnits).Value
      End Get
    End Property

    Public ReadOnly Property ReasonForDespatch() As String
      Get
        ReasonForDespatch = mvClassFields.Item(AppealBudgetDetailFields.abdfReasonForDespatch).Value
      End Get
    End Property

    Public ReadOnly Property Segment() As String
      Get
        Segment = mvClassFields.Item(AppealBudgetDetailFields.abdfSegment).Value
      End Get
    End Property

  End Class
End Namespace
