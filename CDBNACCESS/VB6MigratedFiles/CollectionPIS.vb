

Namespace Access
  Public Class CollectionPIS

    Public Enum CollectionPISRecordSetTypes 'These are bit values
      cpisrtAll = &HFFFFS
      cpisrtCollectionNumber = 1
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionPISFields
      cpisfAll = 0
      cpisfCollectionPISNumber
      cpisfCollectionNumber
      cpisfPisNumber
      cpisfCollectorNumber
      cpisfIssueDate
      cpisfAmount
      cpisfBankedBy
      cpisfBankedOn
      cpisfReconciledOn
      cpisfReconciledStatus
      cpisfAmendedBy
      cpisfAmendedOn
    End Enum

    Public Enum CollectionPISReconciledStatus
      cpisrsUnReconciled = 0
      cpisrsFullyReconciled
      cpisrsTrader
      cpisrsReversed
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvMannedCB As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "collection_pis"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_pis_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("pis_number", CDBField.FieldTypes.cftCharacter)
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("issue_date", CDBField.FieldTypes.cftDate)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("banked_by", CDBField.FieldTypes.cftLong)
          .Add("banked_on", CDBField.FieldTypes.cftDate)
          .Add("reconciled_on", CDBField.FieldTypes.cftDate)
          .Add("reconciled_status")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(CollectionPISFields.cpisfCollectionPISNumber).SetPrimaryKeyOnly()
          .SetUniqueField(CollectionPISFields.cpisfCollectionNumber)
          .SetUniqueField(CollectionPISFields.cpisfPisNumber)
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionPISFields)
      'Add code here to ensure all values are valid before saving
      Dim vAppealColl As New AppealCollection(mvEnv)
      Dim vWhereFields As New CDBFields

      With mvClassFields
        If mvClassFields(CollectionPISFields.cpisfCollectionPISNumber).IntegerValue = 0 Then .Item(CollectionPISFields.cpisfCollectionPISNumber).IntegerValue = mvEnv.GetControlNumber("CF")
        If mvExisting And (.Item(CollectionPISFields.cpisfCollectorNumber).ValueChanged Or .Item(CollectionPISFields.cpisfAmount).ValueChanged) Then
          If HasPayments Then RaiseError(DataAccessErrors.daeCannotChangeCollPISAsPayments)
        End If
        If mvClassFields(CollectionPISFields.cpisfPisNumber).ValueChanged Then
          vAppealColl.Init(CollectionNumber)
          vWhereFields.Add("bank_account", CDBField.FieldTypes.cftCharacter, vAppealColl.BankAccount)
          vWhereFields.Add("pis_number", CDBField.FieldTypes.cftCharacter, PisNumber)
          If mvEnv.Connection.GetCount("pis_numbers", vWhereFields) = 0 Then RaiseError(DataAccessErrors.daeInvalidCode, "PIS Number,Bank Account")

          'check that this pis has not been used for another collection ever
          vWhereFields.Remove(("bank_account"))
          vWhereFields.Add("collection_number", CDBField.FieldTypes.cftLong, "SELECT collection_number FROM appeal_collections WHERE bank_account = '" & vAppealColl.BankAccount & "'", CDBField.FieldWhereOperators.fwoIn)
          If mvEnv.Connection.GetCount("collection_pis", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "Bank Account,PIS Number")
        End If
        .Item(CollectionPISFields.cpisfAmendedOn).Value = TodaysDate()
        .Item(CollectionPISFields.cpisfAmendedBy).Value = mvEnv.User.Logname
      End With
    End Sub

    Private Function GetReconciledStatusFromCode(ByVal pReconciledStatusCode As String) As CollectionPISReconciledStatus
      Select Case pReconciledStatusCode
        Case "F"
          GetReconciledStatusFromCode = CollectionPISReconciledStatus.cpisrsFullyReconciled
        Case "R"
          GetReconciledStatusFromCode = CollectionPISReconciledStatus.cpisrsReversed
        Case "T"
          GetReconciledStatusFromCode = CollectionPISReconciledStatus.cpisrsTrader
        Case Else
          GetReconciledStatusFromCode = CollectionPISReconciledStatus.cpisrsUnReconciled
      End Select
    End Function

    Private Function GetReconciledStatusCode(ByVal pReconciledStatus As CollectionPISReconciledStatus) As String
      Select Case pReconciledStatus
        Case CollectionPISReconciledStatus.cpisrsFullyReconciled
          GetReconciledStatusCode = "F"
        Case CollectionPISReconciledStatus.cpisrsTrader
          GetReconciledStatusCode = "T"
        Case CollectionPISReconciledStatus.cpisrsReversed
          GetReconciledStatusCode = "R"
        Case Else
          GetReconciledStatusCode = ""
      End Select
    End Function

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionPISRecordSetTypes) As String
      Dim vFields As String = ""

      'Modify below to add each recordset type as required
      If pRSType = CollectionPISRecordSetTypes.cpisrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cp")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionPISNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionPISNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionPISRecordSetTypes.cpisrtAll) & " FROM collection_pis cp WHERE collection_pis_number = " & pCollectionPISNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionPISRecordSetTypes.cpisrtAll)
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

    Public Sub InitFromPISNumber(ByVal pEnv As CDBEnvironment, ByRef pCollectionNumber As Integer, ByRef pPisNumber As Integer)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      vWhereFields.Add("collection_number", CDBField.FieldTypes.cftLong, pCollectionNumber)
      vWhereFields.Add("pis_number", CDBField.FieldTypes.cftCharacter, pPisNumber)
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionPISRecordSetTypes.cpisrtAll) & " FROM collection_pis cp WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CollectionPISRecordSetTypes.cpisrtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionPISRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionPISFields.cpisfCollectionPISNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionPISRecordSetTypes.cpisrtAll) = CollectionPISRecordSetTypes.cpisrtAll Then
          .SetItem(CollectionPISFields.cpisfCollectionNumber, vFields)
          .SetItem(CollectionPISFields.cpisfPisNumber, vFields)
          .SetItem(CollectionPISFields.cpisfCollectorNumber, vFields)
          .SetItem(CollectionPISFields.cpisfIssueDate, vFields)
          .SetItem(CollectionPISFields.cpisfAmount, vFields)
          .SetItem(CollectionPISFields.cpisfBankedBy, vFields)
          .SetItem(CollectionPISFields.cpisfBankedOn, vFields)
          .SetItem(CollectionPISFields.cpisfReconciledOn, vFields)
          .SetItem(CollectionPISFields.cpisfReconciledStatus, vFields)
          .SetItem(CollectionPISFields.cpisfAmendedBy, vFields)
          .SetItem(CollectionPISFields.cpisfAmendedOn, vFields)
        End If

        If (pRSType And CollectionPISRecordSetTypes.cpisrtCollectionNumber) = CollectionPISRecordSetTypes.cpisrtCollectionNumber Then
          .SetItem(CollectionPISFields.cpisfCollectionNumber, vFields)
        End If

      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields
      Dim vDeleteLookup As Boolean

      SetValid(CollectionPISFields.cpisfAll)
      If Not mvExisting Then vDeleteLookup = True
      vWhereFields.Add("pis_number", CDBField.FieldTypes.cftCharacter, PisNumber)
      If vDeleteLookup Then mvEnv.Connection.StartTransaction()
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      If vDeleteLookup Then
        mvEnv.Connection.DeleteRecords("pis_numbers", vWhereFields, True)
        vWhereFields.Clear()
        vWhereFields.Add("lookup_item", CDBField.FieldTypes.cftCharacter, PisNumber)
        vWhereFields.Add("lookup_group", CDBField.FieldTypes.cftCharacter, "SELECT lookup_group FROM lookup_groups WHERE table_name = 'pis_numbers'", CDBField.FieldWhereOperators.fwoIn)
        mvEnv.Connection.DeleteRecords("lookup_group_details", vWhereFields, False)
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    'Public Sub Create(ByVal pCollectionNumber As Long, ByVal pCollectionPISNumber As Long, ByVal pIssueDate As String, Optional ByVal pCollectorNumber As Long = 0, Optional ByVal pAmount As Double = 0, Optional ByVal pBankedBy As String = "", Optional ByVal pBankedOn As String = "", Optional ByVal pReconciledOn As String = "")
    '  With mvClassFields
    '    .Item(cpfCollectionNumber).IntegerValue = pCollectionNumber
    '    .Item(cpfCollectionPISNumber).IntegerValue = pCollectionPISNumber
    '    .Item(cpfIssueDate).Value = pIssueDate
    '    If pCollectorNumber > 0 Then .Item(cpfCollectorNumber).IntegerValue = pCollectorNumber
    '    If pAmount > 0 Then .Item(cpfAmount).DoubleValue = pAmount
    '    If Len(pBankedBy) > 0 Then .Item(cpfBankedBy).Value = pBankedBy
    '    If Len(pBankedOn) > 0 Then .Item(cpfBankedOn).Value = pBankedOn
    '    If Len(pReconciledOn) > 0 Then .Item(cpfReconciledOn).Value = pReconciledOn
    '  End With
    'End Sub
    '
    'Public Sub Update(ByVal pCollectorNumber As Long, ByVal pIssueDate As String, Optional ByVal pAmount As Double = 0, Optional ByVal pBankedBy As String = "", Optional ByVal pBankedOn As String = "", Optional ByVal pReconciledOn As String = "")
    '  With mvClassFields
    '    .Item(cpfCollectorNumber).IntegerValue = pCollectorNumber
    '    .Item(cpfIssueDate).Value = pIssueDate
    '    .Item(cpfAmount).DoubleValue = pAmount
    '    .Item(cpfBankedBy).Value = pBankedBy
    '    .Item(cpfBankedOn).Value = pBankedOn
    '    .Item(cpfReconciledOn).Value = pReconciledOn
    '  End With
    'End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(CollectionPISFields.cpisfCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(CollectionPISFields.cpisfPisNumber).Value = pParams("PisNumber").Value
        .Item(CollectionPISFields.cpisfIssueDate).Value = pParams("IssueDate").Value
        If pParams.Exists("CollectorNumber") Then .Item(CollectionPISFields.cpisfCollectorNumber).Value = pParams("CollectorNumber").Value
        If pParams.Exists("Amount") Then .Item(CollectionPISFields.cpisfAmount).Value = pParams("Amount").Value
        If pParams.Exists("BankedBy") Then .Item(CollectionPISFields.cpisfBankedBy).Value = pParams("BankedBy").Value
        If pParams.Exists("BankedOn") Then .Item(CollectionPISFields.cpisfBankedOn).Value = pParams("BankedOn").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("PisNumber") Then .Item(CollectionPISFields.cpisfPisNumber).Value = pParams("PisNumber").Value
        If pParams.Exists("IssueDate") Then .Item(CollectionPISFields.cpisfIssueDate).Value = pParams("IssueDate").Value
        If pParams.Exists("CollectorNumber") Then .Item(CollectionPISFields.cpisfCollectorNumber).Value = pParams("CollectorNumber").Value
        If pParams.Exists("Amount") Then .Item(CollectionPISFields.cpisfAmount).Value = pParams("Amount").Value
        If pParams.Exists("BankedBy") Then .Item(CollectionPISFields.cpisfBankedBy).Value = pParams("BankedBy").Value
        If pParams.Exists("BankedOn") Then .Item(CollectionPISFields.cpisfBankedOn).Value = pParams("BankedOn").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      If DeleteAllowed() Then
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      End If
    End Sub

    Public Sub Reconcile(ByVal pReconciledStatus As CollectionPISReconciledStatus)
      If pReconciledStatus <> CollectionPISReconciledStatus.cpisrsUnReconciled Then
        mvClassFields.Item(CollectionPISFields.cpisfReconciledStatus).Value = GetReconciledStatusCode(pReconciledStatus)
        mvClassFields.Item(CollectionPISFields.cpisfReconciledOn).Value = TodaysDate()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Function DeleteAllowed() As Boolean
      Dim vWhereFields As New CDBFields
      Dim vAppColl As New AppealCollection(mvEnv)
      Dim vDeleteAllowed As Boolean

      vDeleteAllowed = True
      vAppColl.Init(CollectionNumber)
      If vAppColl.CollectionType = AppealCollection.AppealCollectionType.actHouseToHouse Then
        'cannot delete if collectors have been assigned this PIS
        If CollectorNumber > 0 Then
          RaiseError(DataAccessErrors.daeCannotDeleteCollPISAsCollectors)
        End If
      Else
        vWhereFields.Add("collection_number", CDBField.FieldTypes.cftLong, CollectionNumber)
        vWhereFields.Add("collection_pis_number", CDBField.FieldTypes.cftLong, CollectionPisNumber)
        If mvEnv.Connection.GetCount("collection_boxes", vWhereFields) > 0 Then
          RaiseError(DataAccessErrors.daeCannotDeleteCollPISAsBoxes)
        End If
      End If
      If vDeleteAllowed Then
        If HasPayments Then RaiseError(DataAccessErrors.daeCannotDeleteCollPISAsPayments)
      End If
      DeleteAllowed = vDeleteAllowed
    End Function

    Private ReadOnly Property HasPayments() As Boolean
      Get
        Dim vWhereFields As New CDBFields

        vWhereFields.Add("collection_pis_number", CDBField.FieldTypes.cftLong, CollectionPisNumber)
        If mvEnv.Connection.GetCount("collection_payments", vWhereFields) > 0 Then
          HasPayments = True
        Else
          HasPayments = False
        End If
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CollectionPISFields.cpisfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionPISFields.cpisfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(CollectionPISFields.cpisfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property BankedBy() As Integer
      Get
        BankedBy = mvClassFields.Item(CollectionPISFields.cpisfBankedBy).IntegerValue
      End Get
    End Property

    Public ReadOnly Property BankedOn() As String
      Get
        BankedOn = mvClassFields.Item(CollectionPISFields.cpisfBankedOn).Value
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(CollectionPISFields.cpisfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPisNumber() As Integer
      Get
        CollectionPisNumber = mvClassFields.Item(CollectionPISFields.cpisfCollectionPISNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        CollectorNumber = mvClassFields.Item(CollectionPISFields.cpisfCollectorNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property IssueDate() As String
      Get
        IssueDate = mvClassFields.Item(CollectionPISFields.cpisfIssueDate).Value
      End Get
    End Property

    Public ReadOnly Property PisNumber() As String
      Get
        PisNumber = mvClassFields.Item(CollectionPISFields.cpisfPisNumber).Value
      End Get
    End Property

    Public ReadOnly Property ReconciledOn() As String
      Get
        ReconciledOn = mvClassFields.Item(CollectionPISFields.cpisfReconciledOn).Value
      End Get
    End Property

    Public ReadOnly Property ReconciledStatus() As CollectionPISReconciledStatus
      Get
        ReconciledStatus = GetReconciledStatusFromCode(mvClassFields.Item(CollectionPISFields.cpisfReconciledStatus).Value)
      End Get
    End Property

    Public ReadOnly Property ReconciledStatusCode() As String
      Get
        ReconciledStatusCode = mvClassFields.Item(CollectionPISFields.cpisfReconciledStatus).Value
      End Get
    End Property

    Public ReadOnly Property MannedCollectionBoxes() As Collection
      Get
        InitMannedCollectionBoxes()
        MannedCollectionBoxes = mvMannedCB
      End Get
    End Property

    Public Sub MarkReconciled()
      mvClassFields(CollectionPISFields.cpisfReconciledStatus).Value = "F"
      mvClassFields(CollectionPISFields.cpisfReconciledOn).Value = TodaysDate()
    End Sub

    Private Sub InitMannedCollectionBoxes()
      Dim vRS As CDBRecordSet
      Dim vCB As New CollectionBox
      Dim vSQL As String

      If mvMannedCB Is Nothing Then
        mvMannedCB = New Collection
        vCB.Init(mvEnv)
        vSQL = "SELECT " & vCB.GetRecordSetFields(CollectionBox.CollectionBoxRecordSetTypes.cbrtAll) & ", c.contact_number, c.address_number "
        vSQL = vSQL & " FROM collection_boxes cb"
        vSQL = vSQL & " LEFT OUTER JOIN manned_collectors mc ON cb.collector_number = mc.collector_number"
        vSQL = vSQL & " LEFT OUTER JOIN contacts c ON mc.contact_number = c.contact_number"
        vSQL = vSQL & " WHERE cb.collection_pis_number = " & CollectionPisNumber
        vRS = mvEnv.Connection.GetRecordSet(mvEnv.Connection.ProcessAnsiJoins(vSQL))
        While vRS.Fetch() = True
          vCB = New CollectionBox
          vCB.InitFromRecordSet(mvEnv, vRS, CollectionBox.CollectionBoxRecordSetTypes.cbrtAll)
          mvMannedCB.Add(vCB, CStr(vCB.CollectionBoxNumber))
        End While
        vRS.CloseRecordSet()
      End If

    End Sub
  End Class
End Namespace
