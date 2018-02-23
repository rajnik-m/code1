Namespace Access
  Public Class BacsOperation

    Public Enum BacsOperationRecordSetTypes 'These are bit values
      bortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BacsOperationFields
      bofAll = 0
      bofBacsAdviceReason
      bofBacsSource
      bofBacsRecordType
      bofBacsCancellationReason
      bofReversePayment
      bofRefundPayment
      bofNewDirectDebit
      bofUpdateDirectDebit
      bofCancelDirectDebit
      bofUpdatePaymentPlan
      bofCancelPaymentPlan
      bofUpdateContactAccount
      bofNewContactAccount
      bofSkipRejectedPayment
      bofRejectedPaymentsCancelCount
      bofRejectedPaymentsCancelReason
      bofDDCancellationRange
      bofAmendedBy
      bofAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvCancellationStatus As String = ""
    Private mvCancellationReasonDesc As String = ""
    Private mvRejectedPaymentCancelStatus As String = ""
    Private mvRejectedPaymentCancelReasonDesc As String = ""
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "bacs_operations"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("bacs_advice_reason")
          .Add("bacs_source")
          .Add("bacs_record_type")
          .Add("bacs_cancellation_reason")
          .Add("reverse_payment")
          .Add("refund_payment")
          .Add("new_direct_debit")
          .Add("update_direct_debit")
          .Add("cancel_direct_debit")
          .Add("update_payment_plan")
          .Add("cancel_payment_plan")
          .Add("update_contact_account")
          .Add("new_contact_account")
          .Add("skip_rejected_payment")
          .Add("rejected_pmnts_cancel_count", CDBField.FieldTypes.cftLong)
          .Add("rejected_pmnts_cancel_reason")
          .Add("dd_cancellation_range", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(BacsOperationFields.bofBacsAdviceReason).SetPrimaryKeyOnly()
        mvClassFields.Item(BacsOperationFields.bofBacsSource).SetPrimaryKeyOnly()

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBacsSkipRejectedPayment) = False Then
          With mvClassFields
            .Item(BacsOperationFields.bofSkipRejectedPayment).InDatabase = False
            .Item(BacsOperationFields.bofRejectedPaymentsCancelCount).InDatabase = False
            .Item(BacsOperationFields.bofRejectedPaymentsCancelReason).InDatabase = False
          End With
        End If
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As BacsOperationFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(BacsOperationFields.bofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(BacsOperationFields.bofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BacsOperationRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BacsOperationRecordSetTypes.bortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "bo")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBacsAdviceReason As String = "", Optional ByRef pBacsSource As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If Len(pBacsAdviceReason) > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(BacsOperationRecordSetTypes.bortAll)
        vSQL = vSQL & ", x.cancellation_reason_desc AS bacs_cancel_reason_desc, x.status AS bacs_cancel_reason_status, x.reason_required AS bacs_cancel_reason_required"
        vSQL = vSQL & ", y.cancellation_reason_desc AS rej_pmnts_cancel_reason_desc, y.status AS rej_pmnts_cancel_reason_status, y.reason_required AS rej_pmnts_cancel_reason_reqd"
        vSQL = vSQL & " FROM bacs_operations bo"
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT cancellation_reason,cancellation_reason_desc, cr.status, reason_required FROM cancellation_reasons cr, statuses s WHERE cr.status = s.status) x ON bo.bacs_cancellation_reason = x.cancellation_reason"
        vSQL = vSQL & " LEFT OUTER JOIN (SELECT cancellation_reason,cancellation_reason_desc, cr.status, reason_required FROM cancellation_reasons cr, statuses s WHERE cr.status = s.status) y ON bo.rejected_pmnts_cancel_reason = y.cancellation_reason"
        vSQL = vSQL & " WHERE bacs_advice_reason = '" & pBacsAdviceReason & "' AND bacs_source = '" & pBacsSource & "'"
        vRecordSet = pEnv.Connection.GetRecordSetAnsiJoins(vSQL)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BacsOperationRecordSetTypes.bortAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BacsOperationRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BacsOperationFields.bofBacsAdviceReason, vFields)
        .SetItem(BacsOperationFields.bofBacsSource, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BacsOperationRecordSetTypes.bortAll) = BacsOperationRecordSetTypes.bortAll Then
          .SetItem(BacsOperationFields.bofBacsRecordType, vFields)
          .SetItem(BacsOperationFields.bofBacsCancellationReason, vFields)
          .SetItem(BacsOperationFields.bofReversePayment, vFields)
          .SetItem(BacsOperationFields.bofRefundPayment, vFields)
          .SetItem(BacsOperationFields.bofNewDirectDebit, vFields)
          .SetItem(BacsOperationFields.bofUpdateDirectDebit, vFields)
          .SetItem(BacsOperationFields.bofCancelDirectDebit, vFields)
          .SetItem(BacsOperationFields.bofUpdatePaymentPlan, vFields)
          .SetItem(BacsOperationFields.bofCancelPaymentPlan, vFields)
          .SetItem(BacsOperationFields.bofUpdateContactAccount, vFields)
          .SetItem(BacsOperationFields.bofNewContactAccount, vFields)
          .SetOptionalItem(BacsOperationFields.bofSkipRejectedPayment, vFields)
          .SetOptionalItem(BacsOperationFields.bofRejectedPaymentsCancelCount, vFields)
          .SetOptionalItem(BacsOperationFields.bofRejectedPaymentsCancelReason, vFields)
          .SetOptionalItem(BacsOperationFields.bofDDCancellationRange, vFields)
          .SetItem(BacsOperationFields.bofAmendedBy, vFields)
          .SetItem(BacsOperationFields.bofAmendedOn, vFields)
        End If
        'If Cancel PP or Cancel DD options set then expose any associated cancellation status and description
        If mvClassFields.Item(BacsOperationFields.bofCancelDirectDebit).Bool Or mvClassFields.Item(BacsOperationFields.bofCancelPaymentPlan).Bool Then
          mvCancellationStatus = pRecordSet.Fields.Item("bacs_cancel_reason_status").Value
          If pRecordSet.Fields.Item("bacs_cancel_reason_required").Bool Then mvCancellationReasonDesc = pRecordSet.Fields.Item("bacs_cancel_reason_desc").Value
          mvRejectedPaymentCancelStatus = pRecordSet.Fields.Item("rej_pmnts_cancel_reason_status").Value
          If pRecordSet.Fields.Item("rej_pmnts_cancel_reason_reqd").Bool Then mvRejectedPaymentCancelReasonDesc = pRecordSet.Fields.Item("rej_pmnts_cancel_reason_desc").Value
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(BacsOperationFields.bofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(BacsOperationFields.bofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(BacsOperationFields.bofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property BacsAdviceReason() As String
      Get
        BacsAdviceReason = mvClassFields.Item(BacsOperationFields.bofBacsAdviceReason).Value
      End Get
    End Property

    Public ReadOnly Property BacsCancellationReason() As String
      Get
        BacsCancellationReason = mvClassFields.Item(BacsOperationFields.bofBacsCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property BacsRecordType() As String
      Get
        BacsRecordType = mvClassFields.Item(BacsOperationFields.bofBacsRecordType).Value
      End Get
    End Property

    Public ReadOnly Property BacsSource() As String
      Get
        BacsSource = mvClassFields.Item(BacsOperationFields.bofBacsSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelDirectDebit() As Boolean
      Get
        CancelDirectDebit = mvClassFields.Item(BacsOperationFields.bofCancelDirectDebit).Bool
      End Get
    End Property

    Public ReadOnly Property CancelPaymentPlan() As Boolean
      Get
        CancelPaymentPlan = mvClassFields.Item(BacsOperationFields.bofCancelPaymentPlan).Bool
      End Get
    End Property

    Public ReadOnly Property NewContactAccount() As Boolean
      Get
        NewContactAccount = mvClassFields.Item(BacsOperationFields.bofNewContactAccount).Bool
      End Get
    End Property

    Public ReadOnly Property NewDirectDebit() As Boolean
      Get
        NewDirectDebit = mvClassFields.Item(BacsOperationFields.bofNewDirectDebit).Bool
      End Get
    End Property

    Public ReadOnly Property RefundPayment() As Boolean
      Get
        RefundPayment = mvClassFields.Item(BacsOperationFields.bofRefundPayment).Bool
      End Get
    End Property

    Public ReadOnly Property ReversePayment() As Boolean
      Get
        ReversePayment = mvClassFields.Item(BacsOperationFields.bofReversePayment).Bool
      End Get
    End Property

    Public ReadOnly Property UpdateContactAccount() As Boolean
      Get
        UpdateContactAccount = mvClassFields.Item(BacsOperationFields.bofUpdateContactAccount).Bool
      End Get
    End Property

    Public ReadOnly Property UpdateDirectDebit() As Boolean
      Get
        UpdateDirectDebit = mvClassFields.Item(BacsOperationFields.bofUpdateDirectDebit).Bool
      End Get
    End Property

    Public ReadOnly Property UpdatePaymentPlan() As Boolean
      Get
        UpdatePaymentPlan = mvClassFields.Item(BacsOperationFields.bofUpdatePaymentPlan).Bool
      End Get
    End Property

    Public ReadOnly Property SkipRejectedPayment() As Boolean
      Get
        SkipRejectedPayment = mvClassFields.Item(BacsOperationFields.bofSkipRejectedPayment).Bool
      End Get
    End Property

    Public ReadOnly Property RejectedPaymentsCancelCount() As Integer
      Get
        RejectedPaymentsCancelCount = mvClassFields.Item(BacsOperationFields.bofRejectedPaymentsCancelCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RejectedPaymentsCancelReason() As String
      Get
        RejectedPaymentsCancelReason = mvClassFields.Item(BacsOperationFields.bofRejectedPaymentsCancelReason).Value
      End Get
    End Property

    Public ReadOnly Property BacsCancellationStatus() As String
      Get
        BacsCancellationStatus = mvCancellationStatus
      End Get
    End Property

    Public ReadOnly Property BacsCancellationReasonDesc() As String
      Get
        BacsCancellationReasonDesc = mvCancellationReasonDesc
      End Get
    End Property

    Public ReadOnly Property RejectedPaymentCancelStatus() As String
      Get
        RejectedPaymentCancelStatus = mvRejectedPaymentCancelStatus
      End Get
    End Property

    Public ReadOnly Property RejectedPaymentCancelReasonDesc() As String
      Get
        RejectedPaymentCancelReasonDesc = mvRejectedPaymentCancelReasonDesc
      End Get
    End Property

    Public ReadOnly Property DDCancellationRange() As Integer
      Get
        DDCancellationRange = mvClassFields.Item(BacsOperationFields.bofDDCancellationRange).IntegerValue
      End Get
    End Property
  End Class
End Namespace
