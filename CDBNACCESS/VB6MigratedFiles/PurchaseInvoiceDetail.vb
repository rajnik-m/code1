

Namespace Access
  Public Class PurchaseInvoiceDetail

    Public Enum PurchaseInvoiceDetailRecordSetTypes 'These are bit values
      pidrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PurchaseInvoiceDetailFields
      pidfAll = 0
      pidfPurchaseInvoiceNumber
      pidfLineNumber
      pidfLineItem
      pidfLinePrice
      pidfQuantity
      pidfAmount
      pidfAmendedBy
      pidfAmendedOn
      pidfNominalAccount
      pidfDistributionCode
      pidfAdjustmentStatus
      pidfCancellationReason
      pidfCancellationSource
      pidfCancelledBy
      pidfCancelledOn
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
          .DatabaseTableName = "purchase_invoice_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("purchase_invoice_number", CDBField.FieldTypes.cftLong)
          .Add("line_number", CDBField.FieldTypes.cftInteger)
          .Add("line_item")
          .Add("line_price", CDBField.FieldTypes.cftNumeric)
          .Add("quantity", CDBField.FieldTypes.cftInteger)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("nominal_account")
          .Add("distribution_code")
          .Add("adjustment_status")
          .Add("cancellation_reason")
          .Add("cancellation_source")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)

          .Item(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).SetPrimaryKeyOnly()
          .Item(PurchaseInvoiceDetailFields.pidfLineNumber).SetPrimaryKeyOnly()

          .Item(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).PrefixRequired = True
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As PurchaseInvoiceDetailFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(PurchaseInvoiceDetailFields.pidfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PurchaseInvoiceDetailFields.pidfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PurchaseInvoiceDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PurchaseInvoiceDetailRecordSetTypes.pidrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pid")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pPurchaseInvoiceNumber As Integer = 0, Optional ByRef pLineNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pPurchaseInvoiceNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PurchaseInvoiceDetailRecordSetTypes.pidrtAll) & " FROM purchase_invoice_details pid WHERE purchase_invoice_number = " & pPurchaseInvoiceNumber & " AND line_number = " & pLineNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PurchaseInvoiceDetailRecordSetTypes.pidrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PurchaseInvoiceDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber, vFields)
        .SetItem(PurchaseInvoiceDetailFields.pidfLineNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PurchaseInvoiceDetailRecordSetTypes.pidrtAll) = PurchaseInvoiceDetailRecordSetTypes.pidrtAll Then
          .SetItem(PurchaseInvoiceDetailFields.pidfLineItem, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfLinePrice, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfQuantity, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfAmount, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfAmendedBy, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfAmendedOn, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfNominalAccount, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfDistributionCode, vFields)
          'BR17340
          .SetItem(PurchaseInvoiceDetailFields.pidfAdjustmentStatus, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfCancellationReason, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfCancellationSource, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfCancelledBy, vFields)
          .SetItem(PurchaseInvoiceDetailFields.pidfCancelledOn, vFields)

        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PurchaseInvoiceDetailFields.pidfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub
    Public Overridable Sub Update(ByVal pParameterList As CDBParameters)
      Update(pParameterList, True)
    End Sub

    Private Sub Update(ByVal pParameterList As CDBParameters, ByVal pValidate As Boolean)
      'If pValidate Then PreValidateUpdateParameters(pParameterList)
      For Each vClassField As ClassField In mvClassFields
        If vClassField.PrimaryKey = False AndAlso (vClassField.NonUpdatable = False OrElse pValidate = False) Then
          If pParameterList.ContainsKey(vClassField.ParameterName) Then vClassField.Value = pParameterList(vClassField.ParameterName).Value
        End If
      Next
      '  If pValidate Then PostValidateUpdateParameters(pParameterList)
    End Sub

    Public Sub InitFromPI(ByVal pEnv As CDBEnvironment, ByRef pPI As PurchaseInvoice, ByRef pLineItem As String, ByRef pLinePrice As Double, ByRef pQuantity As Integer, ByRef pAmount As Double, ByRef pAccount As String, ByRef pDistributionCode As String)
      Init(pEnv)
      mvClassFields(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).Value = CStr(pPI.PurchaseInvoiceNumber)
      mvClassFields(PurchaseInvoiceDetailFields.pidfLineNumber).Value = CStr(pPI.Details.Count() + 1)
      mvClassFields(PurchaseInvoiceDetailFields.pidfLineItem).Value = pLineItem
      mvClassFields(PurchaseInvoiceDetailFields.pidfLinePrice).Value = CStr(pLinePrice)
      mvClassFields(PurchaseInvoiceDetailFields.pidfQuantity).Value = CStr(pQuantity)
      mvClassFields(PurchaseInvoiceDetailFields.pidfAmount).Value = CStr(pAmount)
      mvClassFields(PurchaseInvoiceDetailFields.pidfNominalAccount).Value = pAccount
      mvClassFields(PurchaseInvoiceDetailFields.pidfDistributionCode).Value = pDistributionCode
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pParams As CDBParameters)
      Init(pEnv)
      mvClassFields(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).IntegerValue = pParams.ParameterExists("PurchaseInvoiceNumber").IntegerValue
      mvClassFields(PurchaseInvoiceDetailFields.pidfLineNumber).IntegerValue = pParams("LineNumber").IntegerValue
      mvClassFields(PurchaseInvoiceDetailFields.pidfLineItem).Value = pParams("LineItem").Value
      mvClassFields(PurchaseInvoiceDetailFields.pidfLinePrice).DoubleValue = pParams("LinePrice").DoubleValue
      mvClassFields(PurchaseInvoiceDetailFields.pidfQuantity).IntegerValue = pParams("Quantity").IntegerValue
      mvClassFields(PurchaseInvoiceDetailFields.pidfAmount).DoubleValue = pParams("Amount").DoubleValue
      mvClassFields(PurchaseInvoiceDetailFields.pidfNominalAccount).Value = pParams.ParameterExists("NominalAccount").Value
      mvClassFields(PurchaseInvoiceDetailFields.pidfDistributionCode).Value = pParams.ParameterExists("DistributionCode").Value
    End Sub

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      LineDataType = mvClassFields.ItemDataType(pAttributeName)
    End Function

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        mvClassFields.ItemValue(pAttributeName) = Value
      End Set
    End Property

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
        AmendedBy = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As Double
      Get
        Amount = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfDistributionCode).Value
      End Get
    End Property

    Public ReadOnly Property LineItem() As String
      Get
        LineItem = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfLineItem).Value
      End Get
    End Property

    Public ReadOnly Property LineNumber() As Integer
      Get
        LineNumber = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfLineNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LinePrice() As Double
      Get
        LinePrice = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfLinePrice).DoubleValue
      End Get
    End Property

    Public ReadOnly Property NominalAccount() As String
      Get
        NominalAccount = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfNominalAccount).Value
      End Get
    End Property

    Public ReadOnly Property PurchaseInvoiceNumber() As Integer
      Get
        PurchaseInvoiceNumber = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(PurchaseInvoiceDetailFields.pidfQuantity).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AdjustmentStatus() As String
      Get
        Return mvClassFields(PurchaseInvoiceDetailFields.pidfAdjustmentStatus).Value
      End Get
    End Property
    Public ReadOnly Property CancellationReason() As String
      Get
        Return mvClassFields(PurchaseInvoiceDetailFields.pidfCancellationReason).Value
      End Get
    End Property
    Public ReadOnly Property CancellationSource() As String
      Get
        Return mvClassFields(PurchaseInvoiceDetailFields.pidfCancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        Return mvClassFields(PurchaseInvoiceDetailFields.pidfCancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn() As String
      Get
        Return mvClassFields(PurchaseInvoiceDetailFields.pidfCancelledOn).Value
      End Get
    End Property

    Public Function GetDataAsParameters() As CDBParameters
      Dim vParams As New CDBParameters
      Dim vField As ClassField

      For Each vField In mvClassFields
        If vField.Name <> "amended_by" And vField.Name <> "amended_on" Then vParams.Add(ProperName((vField.Name)), (vField.FieldType), If(vField.FieldType = CDBField.FieldTypes.cftNumeric, FixedFormat(vField.DoubleValue), vField.Value))
      Next vField
      GetDataAsParameters = vParams
    End Function

    ''' <summary>Cancel the Purchase Invoice Detail.</summary>
    ''' <param name="pCancelReason">Reason for the cancellation.</param>
    ''' <param name="pCancelBy">User performing the cancellation.</param>
    ''' <param name="pCancelOn">Date cancellation takes place.</param>
    ''' <param name="pCancelSource">The source of the cancellation.</param>
    ''' <param name="pAdjustmentStatus">The adjustment status that should be applied.</param>
    Friend Sub Cancel(ByVal pCancelReason As String, ByVal pCancelBy As String, ByVal pCancelOn As String, ByVal pCancelSource As String, ByVal pAdjustmentStatus As String)
      With mvClassFields
        .Item(PurchaseInvoiceDetailFields.pidfCancellationReason).Value = pCancelReason
        .Item(PurchaseInvoiceDetailFields.pidfCancelledOn).Value = pCancelOn
        .Item(PurchaseInvoiceDetailFields.pidfCancelledBy).Value = pCancelBy
        .Item(PurchaseInvoiceDetailFields.pidfCancellationSource).Value = pCancelSource
        .Item(PurchaseInvoiceDetailFields.pidfAdjustmentStatus).Value = pAdjustmentStatus
      End With
    End Sub

    Friend Sub Clone(ByVal pRecord As PurchaseInvoiceDetail, ByVal pPrimaryKeyValue As Integer)
      With mvClassFields
        .Item(PurchaseInvoiceDetailFields.pidfPurchaseInvoiceNumber).IntegerValue = pPrimaryKeyValue
        .Item(PurchaseInvoiceDetailFields.pidfLineNumber).IntegerValue = pRecord.LineNumber
        .Item(PurchaseInvoiceDetailFields.pidfLineItem).Value = pRecord.LineItem
        .Item(PurchaseInvoiceDetailFields.pidfLinePrice).Value = pRecord.LinePrice.ToString("F")
        .Item(PurchaseInvoiceDetailFields.pidfQuantity).IntegerValue = pRecord.Quantity
        .Item(PurchaseInvoiceDetailFields.pidfAmount).Value = pRecord.Amount.ToString("F")
        .Item(PurchaseInvoiceDetailFields.pidfNominalAccount).Value = pRecord.NominalAccount
        .Item(PurchaseInvoiceDetailFields.pidfDistributionCode).Value = pRecord.DistributionCode
        .Item(PurchaseInvoiceDetailFields.pidfAdjustmentStatus).Value = String.Empty
        .Item(PurchaseInvoiceDetailFields.pidfCancellationReason).Value = String.Empty
        .Item(PurchaseInvoiceDetailFields.pidfCancelledBy).Value = String.Empty
        .Item(PurchaseInvoiceDetailFields.pidfCancelledOn).Value = String.Empty
      End With
    End Sub
  End Class
End Namespace
