

Namespace Access
  Public Class SundryCost

    Public Enum SundryCostRecordSetTypes 'These are bit values
      scrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    Public Enum SundryCostTypes
      sctEvent
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SundryCostFields
      scfAll = 0
      scfSundryCostNumber
      scfRecordType
      scfUniqueId
      scfSundryCostType
      scfAmount 'Balance value
      scfPaymentDate 'Balance date
      scfDepositAmount
      scfDepositDate
      scfFullAmount
      scfFullPaymentDate
      scfNotes
      scfSponsorshipValue
      scfContactNumber
      scfAddressNumber
      scfItemReceived
      scfReserveAmount
      scfSoldAmount
      scfSupplierContactNumber
      scfSupplierAddressNumber
      scfAmendedBy
      scfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      Dim vFinancialAnalysis As Boolean

      vFinancialAnalysis = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventFinancialAnalysis)

      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "sundry_costs"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("sundry_cost_number", CDBField.FieldTypes.cftLong)
          .Add("record_type")
          .Add("unique_id", CDBField.FieldTypes.cftLong)
          .Add("sundry_cost_type")
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("payment_date", CDBField.FieldTypes.cftDate)
          .Add("deposit_amount", CDBField.FieldTypes.cftNumeric)
          .Add("deposit_date", CDBField.FieldTypes.cftDate)
          .Add("full_amount", CDBField.FieldTypes.cftNumeric)
          .Add("full_payment_date", CDBField.FieldTypes.cftDate)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("sponsorship_value", CDBField.FieldTypes.cftNumeric)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("item_received", CDBField.FieldTypes.cftCharacter)
          .Add("reserve_amount", CDBField.FieldTypes.cftNumeric)
          .Add("sold_amount", CDBField.FieldTypes.cftNumeric)
          .Add("supplier_contact_number", CDBField.FieldTypes.cftLong)
          .Add("supplier_address_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Item(SundryCostFields.scfNotes).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfSponsorshipValue).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfContactNumber).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfAddressNumber).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfItemReceived).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfReserveAmount).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfSoldAmount).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfSupplierContactNumber).InDatabase = vFinancialAnalysis
          .Item(SundryCostFields.scfSupplierAddressNumber).InDatabase = vFinancialAnalysis
        End With
        mvClassFields(SundryCostFields.scfSundryCostNumber).SetPrimaryKeyOnly()

        mvClassFields(SundryCostFields.scfAmendedBy).PrefixRequired = True
        mvClassFields(SundryCostFields.scfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SundryCostFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SundryCostFields.scfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SundryCostFields.scfAmendedBy).Value = mvEnv.User.Logname
      If mvClassFields.Item(SundryCostFields.scfSundryCostNumber).IntegerValue = 0 Then
        mvClassFields.Item(SundryCostFields.scfSundryCostNumber).Value = CStr(mvEnv.GetControlNumber("SC"))
      End If
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(SundryCostFields.scfUniqueId).Value = pParams("EventNumber").Value
        .Item(SundryCostFields.scfRecordType).Value = GetSundryCostTypeCode(SundryCostTypes.sctEvent)
        Update(pParams)
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("SundryCostType") Then .Item(SundryCostFields.scfSundryCostType).Value = pParams("SundryCostType").Value
        If pParams.Exists("TotalAmount") Then .Item(SundryCostFields.scfFullAmount).Value = pParams("TotalAmount").Value
        If pParams.Exists("DueDate") Then .Item(SundryCostFields.scfFullPaymentDate).Value = pParams("DueDate").Value
        If pParams.Exists("Deposit") Then .Item(SundryCostFields.scfDepositAmount).Value = pParams("Deposit").Value
        If pParams.Exists("DepositPaidDate") Then .Item(SundryCostFields.scfDepositDate).Value = pParams("DepositPaidDate").Value
        If pParams.Exists("Balance") Then .Item(SundryCostFields.scfAmount).Value = pParams("Balance").Value
        If pParams.Exists("BalancePaidDate") Then .Item(SundryCostFields.scfPaymentDate).Value = pParams("BalancePaidDate").Value
        If pParams.Exists("Notes") Then .Item(SundryCostFields.scfNotes).Value = pParams("Notes").Value
        If pParams.Exists("SponsorshipValue") Then .Item(SundryCostFields.scfSponsorshipValue).Value = pParams("SponsorshipValue").Value
        If pParams.Exists("ContactNumber") Then .Item(SundryCostFields.scfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(SundryCostFields.scfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("SupplierContactNumber") Then .Item(SundryCostFields.scfSupplierContactNumber).Value = pParams("SupplierContactNumber").Value
        If pParams.Exists("SupplierAddressNumber") Then .Item(SundryCostFields.scfSupplierAddressNumber).Value = pParams("SupplierAddressNumber").Value
        If pParams.Exists("ItemReceived") Then .Item(SundryCostFields.scfItemReceived).Value = pParams("ItemReceived").Value
        If pParams.Exists("ReserveAmount") Then .Item(SundryCostFields.scfReserveAmount).Value = pParams("ReserveAmount").Value
        If pParams.Exists("SoldAmount") Then .Item(SundryCostFields.scfSoldAmount).Value = pParams("SoldAmount").Value
      End With
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SundryCostRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SundryCostRecordSetTypes.scrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSundryCostNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSundryCostNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SundryCostRecordSetTypes.scrtAll) & " FROM sundry_costs sc WHERE sundry_cost_number = " & pSundryCostNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SundryCostRecordSetTypes.scrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SundryCostRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And SundryCostRecordSetTypes.scrtAll) = SundryCostRecordSetTypes.scrtAll Then
          .SetItem(SundryCostFields.scfSundryCostNumber, vFields)
          .SetItem(SundryCostFields.scfRecordType, vFields)
          .SetItem(SundryCostFields.scfUniqueId, vFields)
          .SetItem(SundryCostFields.scfSundryCostType, vFields)
          .SetItem(SundryCostFields.scfAmount, vFields)
          .SetItem(SundryCostFields.scfPaymentDate, vFields)
          .SetItem(SundryCostFields.scfDepositAmount, vFields)
          .SetItem(SundryCostFields.scfDepositDate, vFields)
          .SetItem(SundryCostFields.scfFullAmount, vFields)
          .SetItem(SundryCostFields.scfFullPaymentDate, vFields)
          .SetItem(SundryCostFields.scfAmendedBy, vFields)
          .SetItem(SundryCostFields.scfAmendedOn, vFields)
          .SetOptionalItem(SundryCostFields.scfNotes, vFields)
          .SetOptionalItem(SundryCostFields.scfSponsorshipValue, vFields)
          .SetOptionalItem(SundryCostFields.scfContactNumber, vFields)
          .SetOptionalItem(SundryCostFields.scfAddressNumber, vFields)
          .SetOptionalItem(SundryCostFields.scfItemReceived, vFields)
          .SetOptionalItem(SundryCostFields.scfReserveAmount, vFields)
          .SetOptionalItem(SundryCostFields.scfSoldAmount, vFields)
          .SetOptionalItem(SundryCostFields.scfSupplierContactNumber, vFields)
          .SetOptionalItem(SundryCostFields.scfSupplierAddressNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(SundryCostFields.scfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
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
        AmendedBy = mvClassFields.Item(SundryCostFields.scfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SundryCostFields.scfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(SundryCostFields.scfAmount).Value
      End Get
    End Property

    Public ReadOnly Property DepositAmount() As String
      Get
        DepositAmount = mvClassFields.Item(SundryCostFields.scfDepositAmount).Value
      End Get
    End Property

    Public ReadOnly Property DepositDate() As String
      Get
        DepositDate = mvClassFields.Item(SundryCostFields.scfDepositDate).Value
      End Get
    End Property

    Public ReadOnly Property FullAmount() As String
      Get
        FullAmount = mvClassFields.Item(SundryCostFields.scfFullAmount).Value
      End Get
    End Property

    Public ReadOnly Property FullPaymentDate() As String
      Get
        FullPaymentDate = mvClassFields.Item(SundryCostFields.scfFullPaymentDate).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(SundryCostFields.scfNotes).Value
      End Get
    End Property
    Public ReadOnly Property SponsorshipValue() As String
      Get
        SponsorshipValue = mvClassFields.Item(SundryCostFields.scfSponsorshipValue).Value
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(SundryCostFields.scfContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(SundryCostFields.scfAddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property SupplierContactNumber() As Integer
      Get
        SupplierContactNumber = mvClassFields.Item(SundryCostFields.scfSupplierContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property SupplierAddressNumber() As Integer
      Get
        SupplierAddressNumber = mvClassFields.Item(SundryCostFields.scfSupplierAddressNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ItemReceived() As Boolean
      Get
        ItemReceived = mvClassFields.Item(SundryCostFields.scfItemReceived).Bool
      End Get
    End Property
    Public ReadOnly Property ReserveAmount() As String
      Get
        ReserveAmount = mvClassFields.Item(SundryCostFields.scfReserveAmount).Value
      End Get
    End Property
    Public ReadOnly Property SoldAmount() As String
      Get
        SoldAmount = mvClassFields.Item(SundryCostFields.scfSoldAmount).Value
      End Get
    End Property
    Public ReadOnly Property PaymentDate() As String
      Get
        PaymentDate = mvClassFields.Item(SundryCostFields.scfPaymentDate).Value
      End Get
    End Property

    Public ReadOnly Property RecordType() As String
      Get
        RecordType = mvClassFields.Item(SundryCostFields.scfRecordType).Value
      End Get
    End Property

    Public ReadOnly Property SundryCostNumber() As Integer
      Get
        SundryCostNumber = mvClassFields.Item(SundryCostFields.scfSundryCostNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SundryCostType() As String
      Get
        SundryCostType = mvClassFields.Item(SundryCostFields.scfSundryCostType).Value
      End Get
    End Property

    Public ReadOnly Property UniqueId() As Integer
      Get
        UniqueId = mvClassFields.Item(SundryCostFields.scfUniqueId).IntegerValue
      End Get
    End Property

    Public WriteOnly Property EventNumber() As Integer
      Set(ByVal Value As Integer)
        mvClassFields.Item(SundryCostFields.scfUniqueId).IntegerValue = Value
        mvClassFields.Item(SundryCostFields.scfRecordType).Value = "E"
      End Set
    End Property

     Public Function GetSundryCostTypeCode(ByRef pSundryCostType As SundryCostTypes) As String
      Select Case pSundryCostType
        Case SundryCostTypes.sctEvent
          Return "E"
        Case Else
          Return ""       'Added fix for compiler warning
      End Select
    End Function
  End Class
End Namespace
