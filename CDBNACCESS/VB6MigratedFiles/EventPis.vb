

Namespace Access
  Public Class EventPis

    Public Enum EventPisRecordSetTypes 'These are bit values
      episrtAll = &HFFFFS
      episrtNumber = 1
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventPisFields
      episfAll = 0
      episfEventPisNumber
      episfEventNumber
      episfPisNumber
      episfEventDelegateNumber
      episfIssueDate
      episfAmount
      episfBankedBy
      episfBankedOn
      episfReconciledOn
      episfReconciledStatus
      episfAmendedBy
      episfAmendedOn
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
          .DatabaseTableName = "event_pis"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_pis_number", CDBField.FieldTypes.cftLong)
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("pis_number")
          .Add("event_delegate_number", CDBField.FieldTypes.cftLong)
          .Add("issue_date", CDBField.FieldTypes.cftDate)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("banked_by", CDBField.FieldTypes.cftLong)
          .Add("banked_on", CDBField.FieldTypes.cftDate)
          .Add("reconciled_on", CDBField.FieldTypes.cftDate)
          .Add("reconciled_status")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(EventPisFields.episfEventPisNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As EventPisFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields(EventPisFields.episfEventPisNumber).IntegerValue = 0 Then mvClassFields.Item(EventPisFields.episfEventPisNumber).IntegerValue = mvEnv.GetControlNumber("EN")
      mvClassFields.Item(EventPisFields.episfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventPisFields.episfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventPisRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventPisRecordSetTypes.episrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ep")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventPisNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pEventPisNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventPisRecordSetTypes.episrtAll) & " FROM event_pis ep WHERE event_pis_number = " & pEventPisNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventPisRecordSetTypes.episrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventPisRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventPisFields.episfEventPisNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventPisRecordSetTypes.episrtAll) = EventPisRecordSetTypes.episrtAll Then
          .SetItem(EventPisFields.episfEventNumber, vFields)
          .SetItem(EventPisFields.episfPisNumber, vFields)
          .SetItem(EventPisFields.episfEventDelegateNumber, vFields)
          .SetItem(EventPisFields.episfIssueDate, vFields)
          .SetItem(EventPisFields.episfAmount, vFields)
          .SetItem(EventPisFields.episfBankedBy, vFields)
          .SetItem(EventPisFields.episfBankedOn, vFields)
          .SetItem(EventPisFields.episfReconciledOn, vFields)
          .SetItem(EventPisFields.episfReconciledStatus, vFields)
          .SetItem(EventPisFields.episfAmendedBy, vFields)
          .SetItem(EventPisFields.episfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub InitFromPISNumber(ByVal pEnv As CDBEnvironment, Optional ByRef pPisNumber As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pPisNumber) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventPisRecordSetTypes.episrtAll) & " FROM event_pis ep WHERE pis_number = '" & pPisNumber & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventPisRecordSetTypes.episrtAll)
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

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vTransaction As Boolean
      Dim vDeleteLookup As Boolean
      Dim vWhereFields As New CDBFields

      SetValid(EventPisFields.episfAll)

      If Not mvExisting Then vDeleteLookup = True
      vWhereFields.Add("pis_number", CDBField.FieldTypes.cftCharacter, PisNumber)
      vWhereFields.Add("bank_account", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEventpisBankAccount))
      If Not mvEnv.Connection.InTransaction And vDeleteLookup Then
        mvEnv.Connection.StartTransaction()
        vTransaction = True
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      If vDeleteLookup Then
        mvEnv.Connection.DeleteRecords("pis_numbers", vWhereFields, True)
        vWhereFields.Clear()
        vWhereFields.Add("lookup_item", CDBField.FieldTypes.cftCharacter, PisNumber)
        vWhereFields.Add("lookup_group", CDBField.FieldTypes.cftCharacter, "SELECT lookup_group FROM lookup_groups WHERE table_name = 'pis_numbers'", CDBField.FieldWhereOperators.fwoIn)
        mvEnv.Connection.DeleteRecords("lookup_group_details", vWhereFields, False)
        If vTransaction Then mvEnv.Connection.CommitTransaction()
      End If
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
        AmendedBy = mvClassFields.Item(EventPisFields.episfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventPisFields.episfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(EventPisFields.episfAmount).Value
      End Get
    End Property

    Public ReadOnly Property BankedBy() As Integer
      Get
        BankedBy = mvClassFields.Item(EventPisFields.episfBankedBy).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BankedOn() As String
      Get
        BankedOn = mvClassFields.Item(EventPisFields.episfBankedOn).Value
      End Get
    End Property

    Public ReadOnly Property EventDelegateNumber() As Integer
      Get
        EventDelegateNumber = mvClassFields.Item(EventPisFields.episfEventDelegateNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventPisFields.episfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventPisNumber() As Integer
      Get
        EventPisNumber = mvClassFields.Item(EventPisFields.episfEventPisNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property IssueDate() As String
      Get
        IssueDate = mvClassFields.Item(EventPisFields.episfIssueDate).Value
      End Get
    End Property

    Public ReadOnly Property PisNumber() As String
      Get
        PisNumber = mvClassFields.Item(EventPisFields.episfPisNumber).Value
      End Get
    End Property

    Public ReadOnly Property ReconciledOn() As String
      Get
        ReconciledOn = mvClassFields.Item(EventPisFields.episfReconciledOn).Value
      End Get
    End Property

    Public ReadOnly Property ReconciledStatus() As String
      Get
        ReconciledStatus = mvClassFields.Item(EventPisFields.episfReconciledStatus).Value
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(EventPisFields.episfEventNumber).Value = pParams("EventNumber").Value
        .Item(EventPisFields.episfPisNumber).Value = pParams("PisNumber").Value
        Update(pParams)

      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("EventDelegateNumber") Then .Item(EventPisFields.episfEventDelegateNumber).Value = pParams("EventDelegateNumber").Value
        If pParams.Exists("IssueDate") Then .Item(EventPisFields.episfIssueDate).Value = pParams("IssueDate").Value
        If pParams.Exists("Amount") Then .Item(EventPisFields.episfAmount).Value = pParams("Amount").Value
        If pParams.Exists("BankedBy") Then .Item(EventPisFields.episfBankedBy).Value = pParams("BankedBy").Value
        If pParams.Exists("BankedOn") Then .Item(EventPisFields.episfBankedOn).Value = pParams("BankedOn").Value
        CheckValidity()
      End With
    End Sub

    Private Sub CheckValidity()
      Dim vWhereFields As New CDBFields

      With mvClassFields
        If mvExisting And .Item(EventPisFields.episfEventDelegateNumber).ValueChanged And Len(.Item(EventPisFields.episfEventDelegateNumber).SetValue) > 0 Then
          RaiseError(DataAccessErrors.daeCannotUpdateDelegateOnEventPIS)
        End If
        If mvClassFields(EventPisFields.episfPisNumber).ValueChanged Then
          vWhereFields.Add("bank_account", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEventpisBankAccount))
          vWhereFields.Add("pis_number", CDBField.FieldTypes.cftCharacter, PisNumber)
          If mvEnv.Connection.GetCount("pis_numbers", vWhereFields) = 0 Then RaiseError(DataAccessErrors.daeInvalidCode, "PIS Number,Bank Account")

          'check that this pis has not been used for another collection ever
          vWhereFields.Remove(("bank_account"))
          If mvEnv.Connection.GetCount("event_pis", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "Bank Account,PIS Number")
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      If DeleteAllowed() Then
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      End If
    End Sub

    Private Function DeleteAllowed() As Boolean
      Dim vWhereFields As New CDBFields
      Dim vEvent As New CDBEvent(mvEnv)
      Dim vDeleteAllowed As Boolean

      vDeleteAllowed = True
      'cannot delete if PIS have been assigned to delegates
      If EventDelegateNumber > 0 Then
        RaiseError(DataAccessErrors.daeCannotDeleteEventPISAsDelegates)
      End If
      vEvent.Init(EventNumber)
      vEvent.CalculateSponsorshipIncome()
      If Val(vEvent.SponsorshipIncome) > 0 Then
        RaiseError(DataAccessErrors.daeCannotDeleteEventPISAsPayments)
      End If
      DeleteAllowed = vDeleteAllowed
    End Function

    Public Sub MarkReconciled()
      mvClassFields(EventPisFields.episfReconciledStatus).Value = "F"
      mvClassFields(EventPisFields.episfReconciledOn).Value = TodaysDate()
    End Sub
  End Class
End Namespace
