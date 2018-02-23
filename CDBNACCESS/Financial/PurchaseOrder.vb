Namespace Access

  Public Class PurchaseOrder
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum PurchaseOrderFields
      AllFields = 0
      PurchaseOrderNumber
      ContactNumber
      AddressNumber
      Amount
      OutputGroup
      PrintedOn
      CancellationReason
      CancelledOn
      CancelledBy
      Source
      PurchaseOrderType
      PurchaseOrderDesc
      PayeeContactNumber
      PayeeAddressNumber
      StartDate
      NumberOfPayments
      DistributionMethod
      PaymentAsPercentage
      Balance
      CancellationSource
      Campaign
      Appeal
      Segment
      PaymentFrequency
      PoAuthorisationLevel
      AuthorisedBy
      AuthorisedOn
      CurrencyCode
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("purchase_order_number", CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("amount", CDBField.FieldTypes.cftNumeric)
        .Add("output_group")
        .Add("printed_on", CDBField.FieldTypes.cftDate)
        .Add("cancellation_reason")
        .Add("cancelled_on", CDBField.FieldTypes.cftDate)
        .Add("cancelled_by")
        .Add("source")
        .Add("purchase_order_type")
        .Add("purchase_order_desc")
        .Add("payee_contact_number", CDBField.FieldTypes.cftLong)
        .Add("payee_address_number", CDBField.FieldTypes.cftLong)
        .Add("start_date", CDBField.FieldTypes.cftDate)
        .Add("number_of_payments", CDBField.FieldTypes.cftInteger)
        .Add("distribution_method")
        .Add("payment_as_percentage")
        .Add("balance", CDBField.FieldTypes.cftNumeric)
        .Add("cancellation_source")
        .Add("campaign")
        .Add("appeal")
        .Add("segment")
        .Add("payment_frequency")
        .Add("po_authorisation_level")
        .Add("authorised_by")
        .Add("authorised_on", CDBField.FieldTypes.cftTime)
        .Add("currency_code")

        .Item(PurchaseOrderFields.PurchaseOrderNumber).PrimaryKey = True
        .SetControlNumberField(PurchaseOrderFields.PurchaseOrderNumber, "PO")

        .Item(PurchaseOrderFields.PurchaseOrderNumber).PrefixRequired = True
        .Item(PurchaseOrderFields.Amount).PrefixRequired = True
        .Item(PurchaseOrderFields.AuthorisedBy).PrefixRequired = True
        .Item(PurchaseOrderFields.AuthorisedOn).PrefixRequired = True
        .Item(PurchaseOrderFields.CancellationReason).PrefixRequired = True
        .Item(PurchaseOrderFields.CancellationSource).PrefixRequired = True
        .Item(PurchaseOrderFields.CancelledBy).PrefixRequired = True
        .Item(PurchaseOrderFields.CancelledOn).PrefixRequired = True
        .Item(PurchaseOrderFields.PayeeAddressNumber).PrefixRequired = True
        .Item(PurchaseOrderFields.PayeeContactNumber).PrefixRequired = True
        '
        Dim vAuthorisation As Boolean = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation)
        .Item(PurchaseOrderFields.PoAuthorisationLevel).InDatabase = vAuthorisation
        .Item(PurchaseOrderFields.AuthorisedBy).InDatabase = vAuthorisation
        .Item(PurchaseOrderFields.AuthorisedOn).InDatabase = vAuthorisation
        .Item(PurchaseOrderFields.CurrencyCode).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderCurrencyCode)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "po"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "purchase_orders"
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
    Public ReadOnly Property PurchaseOrderNumber() As Integer
      Get
        Return mvClassFields(PurchaseOrderFields.PurchaseOrderNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(PurchaseOrderFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        Return mvClassFields(PurchaseOrderFields.AddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Amount() As Double
      Get
        Return mvClassFields(PurchaseOrderFields.Amount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property OutputGroup() As String
      Get
        Return mvClassFields(PurchaseOrderFields.OutputGroup).Value
      End Get
    End Property
    Public ReadOnly Property PrintedOn() As String
      Get
        Return mvClassFields(PurchaseOrderFields.PrintedOn).Value
      End Get
    End Property
    Public ReadOnly Property CancellationReason() As String
      Get
        Return mvClassFields(PurchaseOrderFields.CancellationReason).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn() As String
      Get
        Return mvClassFields(PurchaseOrderFields.CancelledOn).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        Return mvClassFields(PurchaseOrderFields.CancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(PurchaseOrderFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(PurchaseOrderFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Return mvClassFields(PurchaseOrderFields.Source).Value
      End Get
    End Property
    Public ReadOnly Property PurchaseOrderTypeCode() As String
      Get
        Return mvClassFields(PurchaseOrderFields.PurchaseOrderType).Value
      End Get
    End Property
    Public ReadOnly Property PurchaseOrderDesc() As String
      Get
        Return mvClassFields(PurchaseOrderFields.PurchaseOrderDesc).Value
      End Get
    End Property
    Public ReadOnly Property PayeeContactNumber() As Integer
      Get
        Return mvClassFields(PurchaseOrderFields.PayeeContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PayeeAddressNumber() As Integer
      Get
        Return mvClassFields(PurchaseOrderFields.PayeeAddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property StartDate() As String
      Get
        Return mvClassFields(PurchaseOrderFields.StartDate).Value
      End Get
    End Property
    Public ReadOnly Property NumberOfPayments() As String
      Get
        Return mvClassFields(PurchaseOrderFields.NumberOfPayments).Value
      End Get
    End Property
    Public ReadOnly Property DistributionMethod() As PurchaseOrder.PODistributionMethods
      Get
        Select Case mvClassFields.Item(PurchaseOrderFields.DistributionMethod).Value
          Case "S"
            Return PODistributionMethods.podmSequential
          Case "P"
            Return PODistributionMethods.podmProportional
          Case Else
            Return PODistributionMethods.podmNone
        End Select
      End Get
    End Property
    Public ReadOnly Property PaymentAsPercentage() As Boolean
      Get
        Return mvClassFields(PurchaseOrderFields.PaymentAsPercentage).Bool
      End Get
    End Property
    Public Property Balance() As Double
      Get
        Return mvClassFields(PurchaseOrderFields.Balance).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(PurchaseOrderFields.Balance).DoubleValue = Value
      End Set
    End Property
    Public ReadOnly Property CancellationSource() As String
      Get
        Return mvClassFields(PurchaseOrderFields.CancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property Campaign() As String
      Get
        Return mvClassFields(PurchaseOrderFields.Campaign).Value
      End Get
    End Property
    Public ReadOnly Property Appeal() As String
      Get
        Return mvClassFields(PurchaseOrderFields.Appeal).Value
      End Get
    End Property
    Public ReadOnly Property Segment() As String
      Get
        Return mvClassFields(PurchaseOrderFields.Segment).Value
      End Get
    End Property
    Public ReadOnly Property PaymentFrequency() As String
      Get
        Return mvClassFields(PurchaseOrderFields.PaymentFrequency).Value
      End Get
    End Property
    Public ReadOnly Property PoAuthorisationLevel() As String
      Get
        Return mvClassFields(PurchaseOrderFields.PoAuthorisationLevel).Value
      End Get
    End Property
    Public ReadOnly Property AuthorisedBy() As String
      Get
        Return mvClassFields(PurchaseOrderFields.AuthorisedBy).Value
      End Get
    End Property
    Public ReadOnly Property AuthorisedOn() As String
      Get
        Return mvClassFields(PurchaseOrderFields.AuthorisedOn).Value
      End Get
    End Property
    Public ReadOnly Property CurrencyCode() As String
      Get
        Return mvClassFields(PurchaseOrderFields.CurrencyCode).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenrated Code"

    Public Enum PODistributionMethods
      podmNone
      podmProportional
      podmSequential
    End Enum

    Private mvDetails As CollectionList(Of PurchaseOrderDetail)
    Private mvPayments As CollectionList(Of PurchaseOrderPayment)
    Private mvDoNotChangeAuthLevel As Boolean = False 'Only used when authorising a regular payment
    Private mvOldPORegularPaymentAmount As Nullable(Of Double)  'Only used when a regular payment amount is changed

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvDetails = Nothing
      mvPayments = Nothing
      mvDoNotChangeAuthLevel = False
      mvOldPORegularPaymentAmount = Nothing
    End Sub

    Public ReadOnly Property Details() As CollectionList(Of PurchaseOrderDetail)
      Get
        If mvDetails Is Nothing Then mvDetails = New CollectionList(Of PurchaseOrderDetail)
        Return mvDetails
      End Get
    End Property

    Public ReadOnly Property Payments() As CollectionList(Of PurchaseOrderPayment)
      Get
        If mvPayments Is Nothing Then mvPayments = New CollectionList(Of PurchaseOrderPayment)
        Return mvPayments
      End Get
    End Property

    Public Sub InitDetails()
      Dim vPOD As New PurchaseOrderDetail
      Dim vRecordSet As CDBRecordSet

      vPOD.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vPOD.GetRecordSetFields(PurchaseOrderDetail.PurchaseOrderDetailRecordSetTypes.podrtAll) & " FROM purchase_order_details pod WHERE purchase_order_number = " & PurchaseOrderNumber & " ORDER BY line_number")
      While vRecordSet.Fetch() = True
        vPOD = New PurchaseOrderDetail
        vPOD.InitFromRecordSet(mvEnv, vRecordSet, PurchaseOrderDetail.PurchaseOrderDetailRecordSetTypes.podrtAll)
        Details.Add(vPOD.LineNumber.ToString, vPOD)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitPayments()
      Dim vPOP As New PurchaseOrderPayment(mvEnv)
      Dim vRecordSet As CDBRecordSet

      vPOP.Init()
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vPOP.GetRecordSetFields() & " FROM purchase_order_payments pop WHERE purchase_order_number = " & PurchaseOrderNumber & " ORDER BY payment_number")
      While vRecordSet.Fetch() = True
        vPOP = New PurchaseOrderPayment(mvEnv)
        vPOP.InitFromRecordSet(vRecordSet)
        Payments.Add(vPOP.PaymentNumber.ToString, vPOP)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Cancel(ByVal pCancellationReason As String, ByVal pCancellationSource As String, Optional ByVal pCancelledOn As String = "")
      Dim vPOD As New PurchaseOrderDetail
      Dim vTransactionStarted As Boolean

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        InitDetails()
        If mvDetails.Count() > 0 Then
          If Not mvEnv.Connection.InTransaction Then
            mvEnv.Connection.StartTransaction()
            vTransactionStarted = True
          End If
          For Each vPOD In mvDetails
            If vPOD.ProductCode.Length > 0 And vPOD.WarehouseCode.Length > 0 Then
              With vPOD.ProductWarehouse
                If .Existing Then
                  If vPOD.Quantity > .QuantityOnOrder Then
                    .QuantityOnOrder = 0
                  Else
                    .QuantityOnOrder = .QuantityOnOrder - vPOD.Quantity
                  End If
                  .Save(mvEnv.User.UserID, True)
                End If
              End With
            End If
          Next vPOD
        End If
      End If
      mvClassFields(PurchaseOrderFields.CancellationReason).Value = pCancellationReason
      mvClassFields(PurchaseOrderFields.CancelledOn).Value = If(IsDate(pCancelledOn), pCancelledOn, TodaysDate)
      mvClassFields(PurchaseOrderFields.CancelledBy).Value = mvEnv.User.UserID
      If Len(pCancellationSource) > 0 Then mvClassFields(PurchaseOrderFields.CancellationSource).Value = pCancellationSource
      Save(mvEnv.User.UserID, True)
      If vTransactionStarted Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Sub Reinstate()
      Dim vPOD As New PurchaseOrderDetail
      Dim vTransactionStarted As Boolean

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
        InitDetails()
        If mvDetails.Count() > 0 Then
          If Not mvEnv.Connection.InTransaction Then
            mvEnv.Connection.StartTransaction()
            vTransactionStarted = True
          End If
          For Each vPOD In mvDetails
            If Len(vPOD.ProductCode) > 0 And Len(vPOD.WarehouseCode) > 0 Then
              With vPOD.ProductWarehouse
                If .Existing Then
                  .QuantityOnOrder = .QuantityOnOrder + vPOD.Quantity
                  .Save(mvEnv.User.UserID, True)
                End If
              End With
            End If
          Next vPOD
        End If
      End If
      mvClassFields(PurchaseOrderFields.CancellationReason).Value = ""
      mvClassFields(PurchaseOrderFields.CancelledOn).Value = ""
      mvClassFields(PurchaseOrderFields.CancelledBy).Value = ""
      If Len(mvClassFields(PurchaseOrderFields.CancellationSource).Value) > 0 Then mvClassFields(PurchaseOrderFields.CancellationSource).Value = ""
      Save(mvEnv.User.UserID, True)
      If vTransactionStarted Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Sub SaveWithDetails(ByVal pDeleteFirst As Boolean, Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vPOD As PurchaseOrderDetail
      Dim vPOP As PurchaseOrderPayment
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet

      If pDeleteFirst Then 'Delete records first
        vWhereFields.Add((mvClassFields(PurchaseOrderFields.PurchaseOrderNumber).Name), CDBField.FieldTypes.cftLong, PurchaseOrderNumber)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderLink) Then
          'TA BR 11326: Select & use class to Delete so that if necessary, qty on order is reduced on product_warehouses.
          vPOD = New PurchaseOrderDetail
          vPOD.Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vPOD.GetRecordSetFields(PurchaseOrderDetail.PurchaseOrderDetailRecordSetTypes.podrtAll) & " FROM purchase_order_details pod WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          While vRecordSet.Fetch() = True
            vPOD.Init(mvEnv)
            vPOD.InitFromRecordSet(mvEnv, vRecordSet, PurchaseOrderDetail.PurchaseOrderDetailRecordSetTypes.podrtAll)
            vPOD.Delete()
          End While
          vRecordSet.CloseRecordSet()
        Else
          mvEnv.Connection.DeleteRecords("purchase_order_details", vWhereFields)
        End If
        If Not PurchaseOrderType.AdHocPayments Then
          mvEnv.Connection.DeleteRecords("purchase_order_payments", vWhereFields, False)
        End If
        If Not mvExisting Then mvEnv.Connection.DeleteRecords((mvClassFields.DatabaseTableName), vWhereFields)
      End If
      For Each vPOD In Details
        vPOD.Save(pAmendedBy, pAudit)
      Next vPOD
      For Each vPOP In Payments
        If vPOP.PayeeContactNumber = 0 Then
          vPOP.SetPayeeDetails(PayeeContactNumber, PayeeAddressNumber)
        End If
        vPOP.Save(pAmendedBy, pAudit)
      Next vPOP
      MyBase.Save(pAmendedBy, pAudit)
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vTransactionStarted As Boolean
      If Not mvEnv.Connection.InTransaction Then
        mvEnv.Connection.StartTransaction()
        vTransactionStarted = True
      End If

      If mvClassFields(PurchaseOrderFields.Amount).ValueChanged AndAlso mvDoNotChangeAuthLevel = False Then
        'Changing amount may need to change the authorisation level
        SetAuthorisationLevel()
        'Increasing the amount requires re-authorising the purchase order
        If Amount > DoubleValue(mvClassFields(PurchaseOrderFields.Amount).SetValue) Then
          mvClassFields(PurchaseOrderFields.AuthorisedOn).Value = ""
          mvClassFields(PurchaseOrderFields.AuthorisedBy).Value = ""
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderHistory) Then
          Dim vPOH As New PurchaseOrderHistory(mvEnv)
          If mvOldPORegularPaymentAmount.HasValue Then
            vPOH.Create(PurchaseOrderNumber, mvPayments(mvPayments.Count - 1).Amount, mvOldPORegularPaymentAmount.Value, mvClassFields(PurchaseOrderFields.PoAuthorisationLevel).SetValue, mvClassFields(PurchaseOrderFields.AuthorisedBy).SetValue, mvClassFields(PurchaseOrderFields.AuthorisedOn).SetValue)
          Else
            vPOH.Create(PurchaseOrderNumber, Amount, DoubleValue(mvClassFields(PurchaseOrderFields.Amount).SetValue), mvClassFields(PurchaseOrderFields.PoAuthorisationLevel).SetValue, mvClassFields(PurchaseOrderFields.AuthorisedBy).SetValue, mvClassFields(PurchaseOrderFields.AuthorisedOn).SetValue)
          End If
          vPOH.Save(pAmendedBy, pAudit, pJournalNumber)
        End If
      End If
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      If vTransactionStarted Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Function PurchaseOrderType() As PurchaseOrderType
      Dim vPOT As New PurchaseOrderType(mvEnv)
      If PurchaseOrderTypeCode.Length > 0 Then
        vPOT.Init(PurchaseOrderTypeCode)
      Else
        vPOT.Init()
      End If
      Return vPOT
    End Function

    Private Sub SetAuthorisationLevel()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPurchaseOrderAuthorisation) Then
        If PurchaseOrderType.RequiresAuthorisation Then
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("purchase_order_type", PurchaseOrderTypeCode)
          Dim vAmount As String = Amount.ToString
          If mvOldPORegularPaymentAmount.HasValue Then vAmount = mvPayments(mvPayments.Count - 1).Amount.ToString
          vWhereFields.Add("lower_limit", CDBField.FieldTypes.cftNumeric, vAmount, CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("upper_limit", CDBField.FieldTypes.cftNumeric, vAmount, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          Dim vSQL As New SQLStatement(mvEnv.Connection, "po_authorisation_level", "po_authorisation_levels", vWhereFields)
          Dim vDataTable As DataTable = vSQL.GetDataTable
          If vDataTable.Rows.Count <> 1 Then
            RaiseError(DataAccessErrors.daeNoAuthorisationLevel)
          Else
            mvClassFields(PurchaseOrderFields.PoAuthorisationLevel).Value = vDataTable.Rows(0)("po_authorisation_level").ToString
          End If
        End If
      End If
    End Sub

    Public Sub Authorise(ByVal pAuthoriseOn As String, ByVal pAuthorisedBy As String)
      'First check if the user can authorised the purchase order
      If PoAuthorisationLevel.Length > 0 Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("po_authorisation_level", PoAuthorisationLevel)
        vWhereFields.Add("logname", pAuthorisedBy)
        If mvEnv.Connection.GetCount("po_authorisation_users", vWhereFields) > 0 Then
          mvClassFields(PurchaseOrderFields.AuthorisedOn).Value = pAuthoriseOn
          mvClassFields(PurchaseOrderFields.AuthorisedBy).Value = pAuthorisedBy
          Save()
        Else
          RaiseError(DataAccessErrors.daeAccessLevel)
        End If
      End If
    End Sub

    Public Sub AddDetailLine(ByVal pLineItem As String, ByVal pLinePrice As Double, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pBalance As Double, ByVal pAccount As String, ByVal pDistributionCode As String, Optional ByVal pProductCode As String = "", Optional ByVal pWarehouseCode As String = "")
      Dim vPOD As New PurchaseOrderDetail

      If mvDetails Is Nothing Then mvDetails = New CollectionList(Of PurchaseOrderDetail)
      SetValid()
      vPOD.InitFromPO(mvEnv, Me, pLineItem, pLinePrice, pQuantity, pAmount, pBalance, pAccount, pDistributionCode, pProductCode, pWarehouseCode)
      mvDetails.Add(vPOD.LineNumber.ToString, vPOD)
    End Sub

    Public Sub AddDetail(ByVal pPOD As PurchaseOrderDetail)

      If mvDetails Is Nothing Then mvDetails = New CollectionList(Of PurchaseOrderDetail)
      SetValid()
      Dim vParams As CDBParameters = pPOD.GetDataAsParameters
      If vParams.ParameterExists("PurchaseOrderNumber").IntegerValue = 0 Then vParams("PurchaseOrderNumber").Value = PurchaseOrderNumber.ToString
      pPOD.Create(mvEnv, vParams)
      mvDetails.Add(pPOD.LineNumber.ToString, pPOD)
    End Sub

    Public Sub AddPayment(ByVal pPPA As PurchaseOrderPayment)

      If mvPayments Is Nothing Then mvPayments = New CollectionList(Of PurchaseOrderPayment)
      SetValid()
      Dim vParams As CDBParameters = pPPA.GetDataAsParameters
      If vParams.ParameterExists("PurchaseOrderNumber").IntegerValue = 0 Then vParams("PurchaseOrderNumber").Value = PurchaseOrderNumber.ToString

      If vParams.ParameterExists("PaymentNumber").IntegerValue = 0 Then
        If vParams.Exists("PaymentNumber") = False Then vParams.Add("PaymentNumber", CDBField.FieldTypes.cftInteger)
        vParams("PaymentNumber").Value = pPPA.GetNextPaymentNumber.ToString
      End If
      pPPA.Create(vParams)
      mvPayments.Add(pPPA.PaymentNumber.ToString, pPPA)
    End Sub

    Public Overloads Sub Create(ByVal pContactNo As Integer, ByVal pAddressNo As Integer, ByVal pAmount As Double, ByVal pOutputGroup As String)
      InitClassFields()
      SetValid()
      mvClassFields(PurchaseOrderFields.ContactNumber).IntegerValue = pContactNo
      mvClassFields(PurchaseOrderFields.AddressNumber).IntegerValue = pAddressNo
      mvClassFields(PurchaseOrderFields.Amount).DoubleValue = pAmount
      mvClassFields(PurchaseOrderFields.OutputGroup).Value = pOutputGroup
      mvClassFields(PurchaseOrderFields.PayeeContactNumber).IntegerValue = pContactNo
      mvClassFields(PurchaseOrderFields.PayeeAddressNumber).IntegerValue = pAddressNo
    End Sub

    Public Sub CreateFromTrader(ByVal pParams As CDBParameters)
      'NOTE: Dont use init as this class has already been initialised
      mvClassFields(PurchaseOrderFields.ContactNumber).IntegerValue = pParams("POD_ContactNumber").IntegerValue
      mvClassFields(PurchaseOrderFields.AddressNumber).IntegerValue = pParams("POD_AddressNumber").IntegerValue
      If mvClassFields(PurchaseOrderFields.Amount).DoubleValue > 0 Then
        mvClassFields(PurchaseOrderFields.Amount).DoubleValue = pParams("POD_Amount").DoubleValue
      Else
        mvClassFields(PurchaseOrderFields.Amount).DoubleValue = pParams("PPBalance").DoubleValue
      End If
      mvClassFields(PurchaseOrderFields.Balance).DoubleValue = pParams("PPBalance").DoubleValue
      mvClassFields(PurchaseOrderFields.OutputGroup).Value = pParams("POD_OutputGroup").Value
      mvClassFields(PurchaseOrderFields.PurchaseOrderType).Value = pParams("POD_PurchaseOrderType").Value
      mvClassFields(PurchaseOrderFields.PurchaseOrderDesc).Value = pParams.ParameterExists("POD_PurchaseOrderDesc").Value
      mvClassFields(PurchaseOrderFields.PayeeContactNumber).IntegerValue = pParams("POD_PayeeContactNumber").IntegerValue
      mvClassFields(PurchaseOrderFields.PayeeAddressNumber).IntegerValue = pParams("POD_PayeeAddressNumber").IntegerValue
      mvClassFields(PurchaseOrderFields.StartDate).Value = pParams("POD_StartDate").Value
      mvClassFields(PurchaseOrderFields.NumberOfPayments).Value = pParams.ParameterExists("POD_NumberOfPayments").Value
      mvClassFields(PurchaseOrderFields.DistributionMethod).Value = pParams("POD_DistributionMethod").Value
      mvClassFields(PurchaseOrderFields.PaymentAsPercentage).Value = pParams.ParameterExists("POD_PaymentAsPercentage").Value
      mvClassFields(PurchaseOrderFields.Campaign).Value = pParams.ParameterExists("POD_Campaign").Value
      mvClassFields(PurchaseOrderFields.Appeal).Value = pParams.ParameterExists("POD_Appeal").Value
      mvClassFields(PurchaseOrderFields.Segment).Value = pParams.ParameterExists("POD_Segment").Value
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataPurchaseOrderManagement) Then
        mvClassFields(PurchaseOrderFields.PaymentFrequency).Value = pParams.ParameterExists("POD_PaymentFrequency").Value
      End If
      If pParams.ContainsKey("OldPORegularPaymentAmount") Then mvOldPORegularPaymentAmount = pParams("OldPORegularPaymentAmount").DoubleValue
      mvClassFields(PurchaseOrderFields.CurrencyCode).Value = pParams.ParameterExists("POD_CurrencyCode").Value
    End Sub

    Friend Sub UpdateForRegularPayments(ByVal pAmount As Double)
      'Used when authorising a regular payment
      mvDoNotChangeAuthLevel = True
      mvClassFields(PurchaseOrderFields.Amount).DoubleValue += pAmount
      mvClassFields(PurchaseOrderFields.Balance).DoubleValue = pAmount
      Details(0).UpdateForRegularPayments(pAmount)
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If CurrencyCode.Length = 0 Then mvClassFields(PurchaseOrderFields.CurrencyCode).Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCurrencyCode)
    End Sub
#End Region

  End Class
End Namespace