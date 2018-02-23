

Namespace Access
  Public Class CompanyControl

    Public Enum CompanyControlRecordSetTypes 'These are bit values
      cocrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CompanyControlFields
      ccfAll = 0
      ccfCompany
      ccfOverPaymentProduct
      ccfOverPaymentRate
      ccfLockedProduct
      ccfLockedRate
      ccfCancelledProduct
      ccfCancelledRate
      ccfInAdvanceProduct
      ccfInAdvanceRate
      ccfDetailsProduct
      ccfDetailsRate
      ccfAwaitingDespatchProduct
      ccfAwaitingDespatchRate
      ccfCommissionProduct
      ccfCommissionRate
      ccfSundryCreditProduct
      ccfSundryCreditRate
      ccfAmendedOn
      ccfAmendedBy
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvAwaitingDespatchProduct As Product
    Private mvCancelledProduct As Product
    Private mvCommissionProduct As Product
    Private mvDetailsProduct As Product
    Private mvInAdvanceProduct As Product
    Private mvLockedProduct As Product
    Private mvOverPaymentProduct As Product
    Private mvSundryCreditProduct As Product

    'Credit Controls
    Private mvTermsNumber As Integer
    Private mvTermsPeriod As String
    Private mvTermsFrom As String
    Private mvDepositPercentage As Double

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "company_controls"
          .Add("company")
          .Add("over_payment_product")
          .Add("over_payment_rate")
          .Add("locked_product")
          .Add("locked_rate")
          .Add("cancelled_product")
          .Add("cancelled_rate")
          .Add("in_advance_product")
          .Add("in_advance_rate")
          .Add("details_product")
          .Add("details_rate")
          .Add("awaiting_despatch_product")
          .Add("awaiting_despatch_rate")
          .Add("commission_product")
          .Add("commission_rate")
          .Add("sundry_credit_product")
          .Add("sundry_credit_rate")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
        End With

        mvClassFields.Item(CompanyControlFields.ccfCompany).SetPrimaryKeyOnly()
        mvClassFields.Item(CompanyControlFields.ccfCompany).PrefixRequired = True
        mvClassFields.Item(CompanyControlFields.ccfAmendedBy).PrefixRequired = True
        mvClassFields.Item(CompanyControlFields.ccfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As CompanyControlFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CompanyControlFields.ccfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CompanyControlFields.ccfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CompanyControlRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CompanyControlRecordSetTypes.cocrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cc")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCompany As String = "", Optional ByVal pInitCreditControls As Boolean = False)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCompany) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CompanyControlRecordSetTypes.cocrtAll) & " FROM company_controls cc WHERE company = '" & pCompany & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CompanyControlRecordSetTypes.cocrtAll)
          InitProducts()
          If pInitCreditControls Then InitCreditControls()
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

    Public Sub InitFromBankAccount(ByVal pEnv As CDBEnvironment, ByRef pBankAccount As String, Optional ByVal pInitCreditControls As Boolean = False)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CompanyControlRecordSetTypes.cocrtAll) & " FROM bank_accounts ba,company_controls cc WHERE bank_account = '" & pBankAccount & "' AND ba.company = cc.company")
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CompanyControlRecordSetTypes.cocrtAll)
      Else
        vRecordSet.CloseRecordSet()
        RaiseError(DataAccessErrors.daeCompanyInvalid, pBankAccount)
      End If
      vRecordSet.CloseRecordSet()
      InitProducts()
      If pInitCreditControls Then InitCreditControls()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CompanyControlRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CompanyControlFields.ccfCompany, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CompanyControlRecordSetTypes.cocrtAll) = CompanyControlRecordSetTypes.cocrtAll Then
          .SetItem(CompanyControlFields.ccfOverPaymentProduct, vFields)
          .SetItem(CompanyControlFields.ccfOverPaymentRate, vFields)
          .SetItem(CompanyControlFields.ccfLockedProduct, vFields)
          .SetItem(CompanyControlFields.ccfLockedRate, vFields)
          .SetItem(CompanyControlFields.ccfCancelledProduct, vFields)
          .SetItem(CompanyControlFields.ccfCancelledRate, vFields)
          .SetItem(CompanyControlFields.ccfInAdvanceProduct, vFields)
          .SetItem(CompanyControlFields.ccfInAdvanceRate, vFields)
          .SetItem(CompanyControlFields.ccfDetailsProduct, vFields)
          .SetItem(CompanyControlFields.ccfDetailsRate, vFields)
          .SetItem(CompanyControlFields.ccfAwaitingDespatchProduct, vFields)
          .SetItem(CompanyControlFields.ccfAwaitingDespatchRate, vFields)
          .SetItem(CompanyControlFields.ccfCommissionProduct, vFields)
          .SetItem(CompanyControlFields.ccfCommissionRate, vFields)
          .SetItem(CompanyControlFields.ccfSundryCreditProduct, vFields)
          .SetItem(CompanyControlFields.ccfSundryCreditRate, vFields)
          .SetItem(CompanyControlFields.ccfAmendedOn, vFields)
          .SetItem(CompanyControlFields.ccfAmendedBy, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(CompanyControlFields.ccfAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Private Sub InitProducts()
      Dim vParams As New CDBParameters
      Dim vRecordSet As CDBRecordSet
      Dim vProduct As New Product(mvEnv)
      Dim vCode As String

      vProduct.Init()
      If Not vParams.Exists(AwaitingDespatchProductCode) Then vParams.Add(AwaitingDespatchProductCode)
      If Not vParams.Exists(CancelledProductCode) Then vParams.Add(CancelledProductCode)
      If Not vParams.Exists(OverPaymentProductCode) Then vParams.Add(OverPaymentProductCode)
      If Not vParams.Exists(LockedProductCode) Then vParams.Add(LockedProductCode)
      If Not vParams.Exists(DetailsProductCode) Then vParams.Add(DetailsProductCode)
      If CommissionProductCode.Length > 0 Then
        If Not vParams.Exists(CommissionProductCode) Then vParams.Add(CommissionProductCode)
      End If
      If InAdvanceProductCode.Length > 0 Then
        If Not vParams.Exists(InAdvanceProductCode) Then vParams.Add(InAdvanceProductCode)
      End If
      If SundryCreditProductCode.Length > 0 Then
        If Not vParams.Exists(SundryCreditProductCode) Then vParams.Add(SundryCreditProductCode)
      End If
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vProduct.GetRecordSetFields() & " FROM products p WHERE product IN(" & vParams.InList & ")")
      While vRecordSet.Fetch() = True
        vCode = vRecordSet.Fields("product").Value
        vProduct = New Product(mvEnv)
        vProduct.InitFromRecordSet(vRecordSet)
        If vCode = AwaitingDespatchProductCode Then mvAwaitingDespatchProduct = vProduct
        If vCode = CancelledProductCode Then mvCancelledProduct = vProduct
        If vCode = OverPaymentProductCode Then mvOverPaymentProduct = vProduct
        If vCode = LockedProductCode Then mvLockedProduct = vProduct
        If vCode = DetailsProductCode Then mvDetailsProduct = vProduct
        If vCode = CommissionProductCode Then mvCommissionProduct = vProduct
        If vCode = InAdvanceProductCode Then mvInAdvanceProduct = vProduct
        If vCode = SundryCreditProductCode Then mvSundryCreditProduct = vProduct
      End While
      vRecordSet.CloseRecordSet()
    End Sub
    Private Sub InitCreditControls()
      Dim vRS As CDBRecordSet
      Dim vSQL As String

      vSQL = "SELECT terms_number, terms_period, terms_from"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets) Then vSQL = vSQL & ", deposit_percentage"
      vSQL = vSQL & " FROM company_credit_controls WHERE company = '" & mvClassFields.Item(CompanyControlFields.ccfCompany).Value & "'"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      With vRS
        If .Fetch() = True Then
          mvTermsNumber = .Fields.Item(1).IntegerValue
          mvTermsPeriod = .Fields.Item(2).Value
          mvTermsFrom = .Fields.Item(3).Value
          mvDepositPercentage = .Fields.FieldExists("deposit_percentage").DoubleValue
        End If
        .CloseRecordSet()
      End With
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
        AmendedBy = mvClassFields.Item(CompanyControlFields.ccfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CompanyControlFields.ccfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AwaitingDespatchProductCode() As String
      Get
        AwaitingDespatchProductCode = mvClassFields.Item(CompanyControlFields.ccfAwaitingDespatchProduct).Value
      End Get
    End Property

    Public ReadOnly Property AwaitingDespatchRate() As String
      Get
        AwaitingDespatchRate = mvClassFields.Item(CompanyControlFields.ccfAwaitingDespatchRate).Value
      End Get
    End Property

    Public ReadOnly Property CancelledProductCode() As String
      Get
        CancelledProductCode = mvClassFields.Item(CompanyControlFields.ccfCancelledProduct).Value
      End Get
    End Property

    Public ReadOnly Property CancelledRate() As String
      Get
        CancelledRate = mvClassFields.Item(CompanyControlFields.ccfCancelledRate).Value
      End Get
    End Property

    Public ReadOnly Property CommissionProductCode() As String
      Get
        CommissionProductCode = mvClassFields.Item(CompanyControlFields.ccfCommissionProduct).Value
      End Get
    End Property

    Public ReadOnly Property CommissionRate() As String
      Get
        CommissionRate = mvClassFields.Item(CompanyControlFields.ccfCommissionRate).Value
      End Get
    End Property

    Public ReadOnly Property Company() As String
      Get
        Company = mvClassFields.Item(CompanyControlFields.ccfCompany).Value
      End Get
    End Property

    Public ReadOnly Property DetailsProductCode() As String
      Get
        DetailsProductCode = mvClassFields.Item(CompanyControlFields.ccfDetailsProduct).Value
      End Get
    End Property

    Public ReadOnly Property DetailsRate() As String
      Get
        DetailsRate = mvClassFields.Item(CompanyControlFields.ccfDetailsRate).Value
      End Get
    End Property

    Public ReadOnly Property InAdvanceProductCode() As String
      Get
        InAdvanceProductCode = mvClassFields.Item(CompanyControlFields.ccfInAdvanceProduct).Value
      End Get
    End Property

    Public ReadOnly Property InAdvanceRate() As String
      Get
        InAdvanceRate = mvClassFields.Item(CompanyControlFields.ccfInAdvanceRate).Value
      End Get
    End Property

    Public ReadOnly Property LockedProductCode() As String
      Get
        LockedProductCode = mvClassFields.Item(CompanyControlFields.ccfLockedProduct).Value
      End Get
    End Property

    Public ReadOnly Property LockedRate() As String
      Get
        LockedRate = mvClassFields.Item(CompanyControlFields.ccfLockedRate).Value
      End Get
    End Property

    Public ReadOnly Property OverPaymentProductCode() As String
      Get
        OverPaymentProductCode = mvClassFields.Item(CompanyControlFields.ccfOverPaymentProduct).Value
      End Get
    End Property

    Public ReadOnly Property OverPaymentRate() As String
      Get
        OverPaymentRate = mvClassFields.Item(CompanyControlFields.ccfOverPaymentRate).Value
      End Get
    End Property

    Public ReadOnly Property SundryCreditProductCode() As String
      Get
        SundryCreditProductCode = mvClassFields.Item(CompanyControlFields.ccfSundryCreditProduct).Value
      End Get
    End Property

    Public ReadOnly Property SundryCreditRate() As String
      Get
        SundryCreditRate = mvClassFields.Item(CompanyControlFields.ccfSundryCreditRate).Value
      End Get
    End Property

    Public ReadOnly Property AwaitingDespatchProduct() As Product
      Get
        AwaitingDespatchProduct = mvAwaitingDespatchProduct
      End Get
    End Property
    Public ReadOnly Property CancelledProduct() As Product
      Get
        CancelledProduct = mvCancelledProduct
      End Get
    End Property
    Public ReadOnly Property CommissionProduct() As Product
      Get
        CommissionProduct = mvCommissionProduct
      End Get
    End Property
    Public ReadOnly Property DetailsProduct() As Product
      Get
        DetailsProduct = mvDetailsProduct
      End Get
    End Property
    Public ReadOnly Property InAdvanceProduct() As Product
      Get
        InAdvanceProduct = mvInAdvanceProduct
      End Get
    End Property
    Public ReadOnly Property LockedProduct() As Product
      Get
        LockedProduct = mvLockedProduct
      End Get
    End Property
    Public ReadOnly Property OverPaymentProduct() As Product
      Get
        OverPaymentProduct = mvOverPaymentProduct
      End Get
    End Property
    Public ReadOnly Property SundryCreditProduct() As Product
      Get
        SundryCreditProduct = mvSundryCreditProduct
      End Get
    End Property
    Public ReadOnly Property CSTermsNumber() As Integer
      Get
        CSTermsNumber = mvTermsNumber
      End Get
    End Property
    Public ReadOnly Property CSTermsPeriod() As String
      Get
        CSTermsPeriod = mvTermsPeriod
      End Get
    End Property
    Public ReadOnly Property CSTermsFrom() As String
      Get
        CSTermsFrom = mvTermsFrom
      End Get
    End Property
    Public ReadOnly Property CSDepositPercentage() As Double
      Get
        CSDepositPercentage = mvDepositPercentage
      End Get
    End Property
  End Class
End Namespace
