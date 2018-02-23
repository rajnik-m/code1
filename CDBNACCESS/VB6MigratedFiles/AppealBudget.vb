Namespace Access
  Public Class AppealBudget

    Public Enum AppealBudgetRecordSetTypes 'These are bit values
      abrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum AppealBudgetFields
      abfAll = 0
      abfAppealBudgetNumber
      abfCampaign
      abfAppeal
      abfBudgetPeriod
      abfPeriodStartDate
      abfPeriodEndDate
      abfAmendedBy
      abfAmendedOn
      abfPeriodPercentage
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvBudgetDetails As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "appeal_budgets"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("appeal_budget_number", CDBField.FieldTypes.cftLong)
          .Add("campaign")
          .Add("appeal")
          .Add("budget_period", CDBField.FieldTypes.cftInteger)
          .Add("period_start_date", CDBField.FieldTypes.cftDate)
          .Add("period_end_date", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("period_percentage", CDBField.FieldTypes.cftNumeric)
        End With
        mvClassFields.Item(AppealBudgetFields.abfAppealBudgetNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As AppealBudgetFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(AppealBudgetFields.abfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(AppealBudgetFields.abfAmendedBy).Value = mvEnv.User.Logname
      If (pField = AppealBudgetFields.abfAll Or pField = AppealBudgetFields.abfAppealBudgetNumber) And mvClassFields.Item(AppealBudgetFields.abfAppealBudgetNumber).IntegerValue = 0 Then
        mvClassFields.Item(AppealBudgetFields.abfAppealBudgetNumber).Value = CStr(mvEnv.GetControlNumber("AB"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As AppealBudgetRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AppealBudgetRecordSetTypes.abrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ab")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAppealBudgetNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pAppealBudgetNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AppealBudgetRecordSetTypes.abrtAll) & " FROM appeal_budgets ab WHERE appeal_budget_number = " & pAppealBudgetNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, AppealBudgetRecordSetTypes.abrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AppealBudgetRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AppealBudgetFields.abfAppealBudgetNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AppealBudgetRecordSetTypes.abrtAll) = AppealBudgetRecordSetTypes.abrtAll Then
          .SetItem(AppealBudgetFields.abfCampaign, vFields)
          .SetItem(AppealBudgetFields.abfAppeal, vFields)
          .SetItem(AppealBudgetFields.abfBudgetPeriod, vFields)
          .SetItem(AppealBudgetFields.abfPeriodStartDate, vFields)
          .SetItem(AppealBudgetFields.abfPeriodEndDate, vFields)
          .SetItem(AppealBudgetFields.abfAmendedBy, vFields)
          .SetItem(AppealBudgetFields.abfAmendedOn, vFields)
          .SetOptionalItem(AppealBudgetFields.abfPeriodPercentage, vFields)
        End If
      End With
      InitAppealBudgetDetails()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(AppealBudgetFields.abfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pCampaign As String, ByVal pAppeal As String, ByVal pBudgetPeriod As Integer, ByVal pPeriodStartDate As String, ByVal pPeriodEndDate As String, Optional ByVal pPeriodPercentage As String = "")
      With mvClassFields
        .Item(AppealBudgetFields.abfCampaign).Value = pCampaign
        .Item(AppealBudgetFields.abfAppeal).Value = pAppeal
        .Item(AppealBudgetFields.abfBudgetPeriod).IntegerValue = pBudgetPeriod
        .Item(AppealBudgetFields.abfPeriodStartDate).Value = pPeriodStartDate
        .Item(AppealBudgetFields.abfPeriodEndDate).Value = pPeriodEndDate
        If Len(pPeriodPercentage) > 0 Then .Item(AppealBudgetFields.abfPeriodPercentage).DoubleValue = Val(pPeriodPercentage)
      End With
    End Sub

    Public Sub Update(ByVal pBudgetPeriod As Integer, ByVal pPeriodStartDate As String, ByVal pPeriodEndDate As String, ByVal pPeriodPercentage As Double)
      With mvClassFields
        .Item(AppealBudgetFields.abfBudgetPeriod).IntegerValue = pBudgetPeriod
        .Item(AppealBudgetFields.abfPeriodStartDate).Value = pPeriodStartDate
        .Item(AppealBudgetFields.abfPeriodEndDate).Value = pPeriodEndDate
        .Item(AppealBudgetFields.abfPeriodPercentage).DoubleValue = pPeriodPercentage
      End With
    End Sub

    Public Sub Delete()
      Dim vABD As AppealBudgetDetail

      For Each vABD In mvBudgetDetails
        vABD.Delete()
      Next vABD
      If mvExisting Then mvEnv.Connection.DeleteRecords("appeal_budgets", mvClassFields.WhereFields)
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
        AmendedBy = mvClassFields.Item(AppealBudgetFields.abfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(AppealBudgetFields.abfAmendedOn).Value
      End Get
    End Property

    Public Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(AppealBudgetFields.abfAppeal).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(AppealBudgetFields.abfAppeal).Value = Value
      End Set
    End Property

    Public ReadOnly Property AppealBudgetNumber() As Integer
      Get
        AppealBudgetNumber = mvClassFields.Item(AppealBudgetFields.abfAppealBudgetNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BudgetPeriod() As Integer
      Get
        BudgetPeriod = mvClassFields.Item(AppealBudgetFields.abfBudgetPeriod).IntegerValue
      End Get
    End Property

    Public Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(AppealBudgetFields.abfCampaign).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(AppealBudgetFields.abfCampaign).Value = Value
      End Set
    End Property

    Public Property PeriodEndDate() As String
      Get
        PeriodEndDate = mvClassFields.Item(AppealBudgetFields.abfPeriodEndDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(AppealBudgetFields.abfPeriodEndDate).Value = Value
      End Set
    End Property

    Public Property PeriodStartDate() As String
      Get
        PeriodStartDate = mvClassFields.Item(AppealBudgetFields.abfPeriodStartDate).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(AppealBudgetFields.abfPeriodStartDate).Value = Value
      End Set
    End Property

    Public Property PeriodPercentage() As Double
      Get
        PeriodPercentage = mvClassFields.Item(AppealBudgetFields.abfPeriodPercentage).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(AppealBudgetFields.abfPeriodPercentage).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property AppealBudgetDetails() As Collection
      Get
        AppealBudgetDetails = mvBudgetDetails
      End Get
    End Property

    Public Sub InitAppealBudgetDetails()
      Dim vRecordSet As CDBRecordSet
      Dim vABD As New AppealBudgetDetail

      mvBudgetDetails = New Collection
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vABD.GetRecordSetFields(AppealBudgetDetail.AppealBudgetDetailRecordSetTypes.abdrtAll) & ", s.segment_sequence FROM appeal_budget_details abd, segments s where appeal_budget_number = " & AppealBudgetNumber & " AND abd.segment = s.segment AND s.campaign = '" & Campaign & "' AND s.appeal = '" & Appeal & "' ORDER BY s.segment_sequence, abd.segment, reason_for_despatch")
      While vRecordSet.Fetch() = True
        vABD = New AppealBudgetDetail
        vABD.InitFromRecordSet(mvEnv, vRecordSet, AppealBudgetDetail.AppealBudgetDetailRecordSetTypes.abdrtAll)
        mvBudgetDetails.Add(vABD, CStr(vABD.AppealBudgetDetailsNumber))
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Function AddAppealBudgetDetail(ByRef pSegment As String, ByRef pReasonForDespatch As String, ByRef pForecastUnits As String, ByRef pBudgetedCost As String, ByRef pBudgetedIncome As String) As AppealBudgetDetail
      Dim vABD As New AppealBudgetDetail

      vABD.Create(mvEnv, AppealBudgetNumber, pSegment, pReasonForDespatch, pForecastUnits, pBudgetedCost, pBudgetedIncome)
      'mvBudgetDetails.Add vABD, CStr(vABD.AppealBudgetDetailsNumber)
      Return vABD
    End Function

  End Class
End Namespace
