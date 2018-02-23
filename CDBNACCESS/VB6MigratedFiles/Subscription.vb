

Namespace Access
  Public Class Subscription

    Public Enum SubscriptionRecordSetTypes 'These are bit values
      subrstAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SubscriptionFields
      sfAll = 0
      sfSubscriptionNumber
      sfOrderNumber
      sfContactNumber
      sfAddressNumber
      sfProduct
      sfQuantity
      sfValidFrom
      sfValidTo
      sfCancellationReason
      sfCancelledOn
      sfCancelledBy
      sfReasonForDespatch
      sfDespatchMethod
      sfAmendedBy
      sfAmendedOn
      sfCancellationSource
      sfCommunicationNumber
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAmendedValid As Boolean
    Private mvPPD As PaymentPlanDetail

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "subscriptions"
          .Add("subscription_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("product")
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("valid_from", CDBField.FieldTypes.cftDate)
          .Add("valid_to", CDBField.FieldTypes.cftDate)
          .Add("cancellation_reason")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("cancelled_by")
          .Add("reason_for_despatch")
          .Add("despatch_method")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("cancellation_source")
          .Add("communication_number", CDBField.FieldTypes.cftLong)
        End With
        mvClassFields.Item(SubscriptionFields.sfSubscriptionNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(SubscriptionFields.sfCommunicationNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber)
      Else
        mvClassFields.ClearItems()
      End If
      mvAmendedValid = False
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As SubscriptionFields)
      'Add code here to ensure all values are valid before saving
      If Not mvAmendedValid Then
        mvClassFields.Item(SubscriptionFields.sfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(SubscriptionFields.sfAmendedBy).Value = mvEnv.User.UserID
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SubscriptionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SubscriptionRecordSetTypes.subrstAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "su")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSubscriptionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSubscriptionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SubscriptionRecordSetTypes.subrstAll) & " FROM subscriptions WHERE subscription_number = " & pSubscriptionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SubscriptionRecordSetTypes.subrstAll)
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

    Public Sub InitNewFromPayPlanDetail(ByVal pEnv As CDBEnvironment, ByRef pDetail As PaymentPlanDetail, ByRef pReasonForDespatch As String, ByVal pRSType As SubscriptionRecordSetTypes)

      mvEnv = pEnv
      InitClassFields()

      'Always include the primary key attributes
      'sfSubscriptionNumber set on .Save

      'Modify below to handle each recordset type as required
      If (pRSType And SubscriptionRecordSetTypes.subrstAll) = SubscriptionRecordSetTypes.subrstAll Then
        With pDetail
          mvClassFields.Item(SubscriptionFields.sfOrderNumber).Value = CStr(.PlanNumber)
          mvClassFields.Item(SubscriptionFields.sfContactNumber).Value = CStr(.ContactNumber)
          mvClassFields.Item(SubscriptionFields.sfAddressNumber).Value = CStr(.AddressNumber)
          mvClassFields.Item(SubscriptionFields.sfProduct).Value = .ProductCode
          mvClassFields.Item(SubscriptionFields.sfQuantity).Value = CStr(.Quantity)
          mvClassFields.Item(SubscriptionFields.sfValidFrom).Value = .SubscriptionValidFrom
          mvClassFields.Item(SubscriptionFields.sfValidTo).Value = .SubscriptionValidTo
          If Len(pReasonForDespatch) = 0 Then
            mvClassFields.Item(SubscriptionFields.sfReasonForDespatch).Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReasonForDespatch)
          Else
            mvClassFields.Item(SubscriptionFields.sfReasonForDespatch).Value = pReasonForDespatch
          End If
          mvClassFields.Item(SubscriptionFields.sfDespatchMethod).Value = .DespatchMethod
          mvClassFields.Item(SubscriptionFields.sfCommunicationNumber).Value = .CommunicationNumber
        End With
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SubscriptionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SubscriptionFields.sfSubscriptionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SubscriptionRecordSetTypes.subrstAll) = SubscriptionRecordSetTypes.subrstAll Then
          .SetItem(SubscriptionFields.sfOrderNumber, vFields)
          .SetItem(SubscriptionFields.sfContactNumber, vFields)
          .SetItem(SubscriptionFields.sfAddressNumber, vFields)
          .SetItem(SubscriptionFields.sfProduct, vFields)
          .SetItem(SubscriptionFields.sfQuantity, vFields)
          .SetItem(SubscriptionFields.sfValidFrom, vFields)
          .SetItem(SubscriptionFields.sfValidTo, vFields)
          .SetItem(SubscriptionFields.sfCancellationReason, vFields)
          .SetItem(SubscriptionFields.sfCancelledOn, vFields)
          .SetItem(SubscriptionFields.sfCancelledBy, vFields)
          .SetItem(SubscriptionFields.sfReasonForDespatch, vFields)
          .SetItem(SubscriptionFields.sfDespatchMethod, vFields)
          .SetItem(SubscriptionFields.sfAmendedBy, vFields)
          .SetItem(SubscriptionFields.sfAmendedOn, vFields)
          .SetOptionalItem(SubscriptionFields.sfCancellationSource, vFields)
          .SetOptionalItem(SubscriptionFields.sfCommunicationNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SubscriptionFields.sfAll)
      If Not mvExisting Then
        If Len(mvClassFields.Item(SubscriptionFields.sfSubscriptionNumber).Value) = 0 Then mvClassFields.Item(SubscriptionFields.sfSubscriptionNumber).Value = CStr(mvEnv.GetControlNumber("S"))
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Cancel(ByRef pCancellationReason As String, Optional ByRef pCancelledBy As String = "", Optional ByVal pCancellationSource As String = "")

      If Len(pCancellationReason) > 0 Then
        If Len(pCancelledBy) = 0 Then pCancelledBy = mvEnv.User.UserID
        mvClassFields.Item(SubscriptionFields.sfValidTo).Value = TodaysDate()
        If CDate(ValidFrom) > CDate(ValidTo) Then
          mvClassFields.Item(SubscriptionFields.sfValidFrom).Value = TodaysDate()
        End If
        mvClassFields.Item(SubscriptionFields.sfCancellationReason).Value = pCancellationReason
        mvClassFields.Item(SubscriptionFields.sfCancellationSource).Value = pCancellationSource
        mvClassFields.Item(SubscriptionFields.sfCancelledBy).Value = pCancelledBy
        mvClassFields.Item(SubscriptionFields.sfCancelledOn).Value = TodaysDate()

      End If
    End Sub

    Public Sub CancelAnyExisting(ByRef pPPNumber As Integer, ByRef pContactNumber As Integer, ByRef pProductCode As String, ByRef pCancellationReason As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, pPPNumber)
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      vWhereFields.Add("product", CDBField.FieldTypes.cftCharacter, pProductCode)
      vWhereFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter)

      vUpdateFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter, pCancellationReason)
      vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
      vUpdateFields.Add("cancelled_by", CDBField.FieldTypes.cftCharacter, "automatic")
      vUpdateFields.Add("cancelled_on", CDBField.FieldTypes.cftDate, TodaysDate)
      vUpdateFields.Add("amended_by", CDBField.FieldTypes.cftCharacter, "automatic")
      vUpdateFields.Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      mvEnv.Connection.UpdateRecords("subscriptions", vUpdateFields, vWhereFields, False)
    End Sub

    Public Sub CreateForFutureChange(ByRef pPP As PaymentPlan, ByRef pPPD As PaymentPlanDetail, ByRef pMember As Member, ByRef pChangeDate As String)
      Dim vExpiry As Date

      If mvEnv.GetConfigOption("subscription_extension") Then
        vExpiry = CDate(pPP.StartDate).AddYears(99)
      Else
        If pMember.MembershipType.Annual = "Y" Then
          vExpiry = CDate(pPP.RenewalDate).AddMonths(pMember.MembershipType.SuspensionGrace).AddDays(-1)
        Else
          vExpiry = CDate(pMember.Joined).AddYears(99).AddMonths(pMember.MembershipType.SuspensionGrace).AddDays(-1)
        End If
      End If

      PaymentPlanNumber = pPP.PlanNumber
      ContactNumber = pPPD.ContactNumber
      AddressNumber = pPPD.AddressNumber
      Product = pPPD.ProductCode
      Quantity = CInt(pPPD.Quantity)
      ValidFrom = pChangeDate
      mvClassFields.Item(SubscriptionFields.sfValidTo).Value = CStr(vExpiry)
      If pPPD.DespatchMethod.Length > 0 Then
        DespatchMethod = pPPD.DespatchMethod
      Else
        DespatchMethod = "POST"
      End If
      ReasonForDespatch = pPP.ReasonForDespatch
      Save("automatic")
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(SubscriptionFields.sfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SubscriptionFields.sfAddressNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(SubscriptionFields.sfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SubscriptionFields.sfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(SubscriptionFields.sfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(SubscriptionFields.sfCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(SubscriptionFields.sfCancelledBy).Value
      End Get
    End Property
    'Friend Property Let CancelledBy(pNewValue As String)
    '  mvClassFields.Item(sfCancelledBy).Value = pNewValue
    'End Property

    Public ReadOnly Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(SubscriptionFields.sfCancelledOn).Value
      End Get
    End Property
    'Friend Property Let CancelledOn(pNewValue As String)
    '  mvClassFields.Item(sfCancelledOn).Value = pNewValue
    'End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(SubscriptionFields.sfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SubscriptionFields.sfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public Property DespatchMethod() As String
      Get
        DespatchMethod = mvClassFields.Item(SubscriptionFields.sfDespatchMethod).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SubscriptionFields.sfDespatchMethod).Value = Value
      End Set
    End Property

    Public Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(SubscriptionFields.sfOrderNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SubscriptionFields.sfOrderNumber).IntegerValue = Value
      End Set
    End Property

    Public Property Product() As String
      Get
        Product = mvClassFields.Item(SubscriptionFields.sfProduct).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SubscriptionFields.sfProduct).Value = Value
      End Set
    End Property

    Public Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(SubscriptionFields.sfQuantity).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SubscriptionFields.sfQuantity).IntegerValue = Value
      End Set
    End Property

    Public Property ReasonForDespatch() As String
      Get
        ReasonForDespatch = mvClassFields.Item(SubscriptionFields.sfReasonForDespatch).Value
      End Get
      Set(ByVal Value As String)
        If Len(Value) = 0 Then Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlReasonForDespatch)
        mvClassFields.Item(SubscriptionFields.sfReasonForDespatch).Value = Value
      End Set
    End Property

    Public Property SubscriptionNumber() As Integer
      Get
        SubscriptionNumber = mvClassFields.Item(SubscriptionFields.sfSubscriptionNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SubscriptionFields.sfSubscriptionNumber).IntegerValue = Value
      End Set
    End Property

    Public Property ValidFrom() As String
      Get
        ValidFrom = mvClassFields.Item(SubscriptionFields.sfValidFrom).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SubscriptionFields.sfValidFrom).Value = Value
      End Set
    End Property

    Public ReadOnly Property ValidTo() As String
      Get
        ValidTo = mvClassFields.Item(SubscriptionFields.sfValidTo).Value
      End Get
    End Property

    Public Property CommunicationNumber() As String
      Get
        CommunicationNumber = mvClassFields(SubscriptionFields.sfCommunicationNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SubscriptionFields.sfCommunicationNumber).Value = Value
      End Set
    End Property
    Public ReadOnly Property PaymentPlanDetail() As PaymentPlanDetail
      Get
        Dim vRS As CDBRecordSet
        Dim vSQL As String

        If mvPPD Is Nothing Then
          mvPPD = New PaymentPlanDetail
          mvPPD.Init(mvEnv)
          vSQL = "SELECT " & mvPPD.GetRecordSetFields(PaymentPlanDetail.PaymentPlanDetailRecordSetTypes.odrtAll) & " FROM order_details od, products p, rates r WHERE order_number = " & PaymentPlanNumber & " AND contact_number = " & ContactNumber & " AND address_number = " & AddressNumber & " AND od.product = '" & Product & "' AND od.despatch_method = '" & DespatchMethod & "' AND quantity = " & Quantity
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationNumber) = True And Val(CommunicationNumber) > 0 Then
            vSQL = vSQL & " AND communication_number = " & CommunicationNumber
          End If
          vSQL = vSQL & " AND od.product = p.product AND p.product = r.product and od.rate = r.rate"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          If vRS.Fetch() = True Then mvPPD.InitFromRecordSet(mvEnv, vRS, PaymentPlanDetail.PaymentPlanDetailRecordSetTypes.odrtAll)
          vRS.CloseRecordSet()
        End If
        PaymentPlanDetail = mvPPD
      End Get
    End Property
    'Friend Property Let ValidTo(pNewValue As String)
    '  mvClassFields.Item(sfValidTo).Value = pNewValue
    'End Property

    Public Sub SetValidTo(ByRef pStartDate As String, ByRef pExpiry As String)
      If mvEnv.GetConfigOption("subscription_extension") Then
        mvClassFields.Item(SubscriptionFields.sfValidTo).Value = CDate(pStartDate).AddYears(99).ToString(CAREDateFormat)
      Else
        mvClassFields.Item(SubscriptionFields.sfValidTo).Value = pExpiry
      End If
    End Sub

    Friend Sub UnCancel()
      mvClassFields.Item(SubscriptionFields.sfCancellationReason).Value = ""
      mvClassFields.Item(SubscriptionFields.sfCancelledBy).Value = ""
      mvClassFields.Item(SubscriptionFields.sfCancelledOn).Value = ""
      mvClassFields.Item(SubscriptionFields.sfCancellationSource).Value = ""
    End Sub

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(SubscriptionFields.sfAmendedOn).Value = pAmendedOn
      mvClassFields.Item(SubscriptionFields.sfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub
  End Class
End Namespace
