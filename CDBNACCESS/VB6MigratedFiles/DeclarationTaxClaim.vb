

Namespace Access
  Public Class DeclarationTaxClaim

    Public Enum DeclarationTaxClaimRecordSetTypes 'These are bit values
      dtcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum DeclarationTaxClaimFields
      dtcfAll = 0
      dtcfClaimNumber
      dtcfClaimGeneratedDate
      dtcfAmountClaimed
      dtcfClaimTaxYearStart
      dtcfCalculatedTaxAmount
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvTotalNetAmount As Double

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "declaration_tax_claims"
          .Add("claim_number", CDBField.FieldTypes.cftLong)
          .Add("claim_generated_date", CDBField.FieldTypes.cftDate)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
          .Add("claim_tax_year_start", CDBField.FieldTypes.cftDate)
          .Add("calculated_tax_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults(ByVal pClaimNumber As Integer)
      'Add code here to initialise the class with default values for a new record
      'mvClassFields.Item(dtcfClaimNumber).Value = mvEnv.GetControlNumber("TC")
      mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimNumber).Value = CStr(pClaimNumber)
      mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimGeneratedDate).Value = TodaysDate()
      mvClassFields.Item(DeclarationTaxClaimFields.dtcfAmountClaimed).Value = CStr(0)
      mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimTaxYearStart).Value = CStr(CDate("06/04/" & Year(CDate(TodaysDate()))))
      mvClassFields.Item(DeclarationTaxClaimFields.dtcfCalculatedTaxAmount).Value = CStr(0)
    End Sub

    Private Sub SetValid(ByRef pField As DeclarationTaxClaimFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As DeclarationTaxClaimRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = DeclarationTaxClaimRecordSetTypes.dtcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "dtc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pClaimNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pClaimNumber > 0 Then
        InitClassFields()
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(DeclarationTaxClaimRecordSetTypes.dtcrtAll) & " FROM declaration_tax_claims WHERE claim_number = " & pClaimNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, DeclarationTaxClaimRecordSetTypes.dtcrtAll)
        Else
          InitClassFields()
          SetDefaults(pClaimNumber)
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults(0)
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As DeclarationTaxClaimRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(DeclarationTaxClaimFields.dtcfClaimNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And DeclarationTaxClaimRecordSetTypes.dtcrtAll) = DeclarationTaxClaimRecordSetTypes.dtcrtAll Then
          .SetItem(DeclarationTaxClaimFields.dtcfClaimGeneratedDate, vFields)
          .SetItem(DeclarationTaxClaimFields.dtcfAmountClaimed, vFields)
          .SetItem(DeclarationTaxClaimFields.dtcfClaimTaxYearStart, vFields)
          .SetItem(DeclarationTaxClaimFields.dtcfCalculatedTaxAmount, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(DeclarationTaxClaimFields.dtcfAll)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Public Sub CalculateTaxToReclaim(ByVal pTaxPercent As Integer)
      AmountClaimed = FixTwoPlaces(mvTotalNetAmount * (pTaxPercent / (100 - pTaxPercent)))
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public Property AmountClaimed() As Double
      Get
        AmountClaimed = mvClassFields.Item(DeclarationTaxClaimFields.dtcfAmountClaimed).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationTaxClaimFields.dtcfAmountClaimed).DoubleValue = Value
      End Set
    End Property

    Public ReadOnly Property ClaimGeneratedDate() As String
      Get
        ClaimGeneratedDate = mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimGeneratedDate).Value
      End Get
    End Property

    Public Property ClaimNumber() As Integer
      Get
        ClaimNumber = mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimNumber).IntegerValue = Value
      End Set
    End Property

    Public Property ClaimTaxYearStart() As Date
      Get
        ClaimTaxYearStart = CDate(mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimTaxYearStart).Value)
      End Get
      Set(ByVal Value As Date)
        mvClassFields.Item(DeclarationTaxClaimFields.dtcfClaimTaxYearStart).Value = CStr(Value)
      End Set
    End Property

    Public Property CalculatedTaxAmount() As Double
      Get
        CalculatedTaxAmount = mvClassFields.Item(DeclarationTaxClaimFields.dtcfCalculatedTaxAmount).DoubleValue
      End Get
      Set(ByVal Value As Double)
        mvClassFields.Item(DeclarationTaxClaimFields.dtcfCalculatedTaxAmount).DoubleValue = Value
      End Set
    End Property

    Public Property TotalNetAmount() As Double
      Get
        TotalNetAmount = mvTotalNetAmount
      End Get
      Set(ByVal Value As Double)
        mvTotalNetAmount = Value
      End Set
    End Property
  End Class
End Namespace
