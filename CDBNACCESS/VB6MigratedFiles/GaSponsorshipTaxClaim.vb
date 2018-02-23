Namespace Access
  Public Class GaSponsorshipTaxClaim

    Public Enum GaSponsorshipTaxClaimRecordSetTypes 'These are bit values
      gstcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GaSponsorshipTaxClaimFields
      gstcfAll = 0
      gstcfClaimNumber
      gstcfClaimGeneratedDate
      gstcfClaimTaxYearStart
      gstcfClaimedProduct
      gstcfUnclaimedProduct
      gstcfAmountClaimed
      gstcfCalculatedTaxAmount
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvTaxPercent As Integer
    Private mvTotalNetAmount As Double

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "ga_sponsorship_tax_claims"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("claim_number", CDBField.FieldTypes.cftLong)
          .Add("claim_generated_date", CDBField.FieldTypes.cftDate)
          .Add("claim_tax_year_start", CDBField.FieldTypes.cftDate)
          .Add("claimed_product")
          .Add("unclaimed_product")
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
          .Add("calculated_tax_amount", CDBField.FieldTypes.cftNumeric)
        End With

        mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults(Optional ByVal pClaimNumber As Integer = 0)
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimNumber).Value = CStr(pClaimNumber)
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimGeneratedDate).Value = TodaysDate()
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfAmountClaimed).Value = CStr(0)
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfCalculatedTaxAmount).Value = CStr(0)
    End Sub

    Private Sub SetValid(ByVal pField As GaSponsorshipTaxClaimFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GaSponsorshipTaxClaimRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GaSponsorshipTaxClaimRecordSetTypes.gstcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "gstc")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pClaimNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pClaimNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GaSponsorshipTaxClaimRecordSetTypes.gstcrtAll) & " FROM ga_sponsorship_tax_claims gstc WHERE claim_number = " & pClaimNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GaSponsorshipTaxClaimRecordSetTypes.gstcrtAll)
        Else
          InitClassFields()
          SetDefaults(pClaimNumber)
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GaSponsorshipTaxClaimRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(GaSponsorshipTaxClaimFields.gstcfClaimNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GaSponsorshipTaxClaimRecordSetTypes.gstcrtAll) = GaSponsorshipTaxClaimRecordSetTypes.gstcrtAll Then
          .SetItem(GaSponsorshipTaxClaimFields.gstcfClaimGeneratedDate, vFields)
          .SetItem(GaSponsorshipTaxClaimFields.gstcfClaimTaxYearStart, vFields)
          .SetItem(GaSponsorshipTaxClaimFields.gstcfClaimedProduct, vFields)
          .SetItem(GaSponsorshipTaxClaimFields.gstcfUnclaimedProduct, vFields)
          .SetItem(GaSponsorshipTaxClaimFields.gstcfAmountClaimed, vFields)
          .SetItem(GaSponsorshipTaxClaimFields.gstcfCalculatedTaxAmount, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GaSponsorshipTaxClaimFields.gstcfAll)
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

    Public ReadOnly Property AmountClaimed() As Double
      Get
        AmountClaimed = CDbl(mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfAmountClaimed).Value)
      End Get
    End Property

    Public ReadOnly Property CalculatedTaxAmount() As Double
      Get
        CalculatedTaxAmount = CDbl(mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfCalculatedTaxAmount).Value)
      End Get
    End Property

    Public ReadOnly Property ClaimGeneratedDate() As String
      Get
        ClaimGeneratedDate = mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimGeneratedDate).Value
      End Get
    End Property

    Public ReadOnly Property ClaimNumber() As Integer
      Get
        ClaimNumber = CInt(mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimNumber).Value)
      End Get
    End Property

    Public ReadOnly Property ClaimTaxYearStart() As String
      Get
        ClaimTaxYearStart = mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimTaxYearStart).Value
      End Get
    End Property

    Public ReadOnly Property ClaimedProduct() As String
      Get
        ClaimedProduct = mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimedProduct).Value
      End Get
    End Property

    Public ReadOnly Property UnclaimedProduct() As String
      Get
        UnclaimedProduct = mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfUnclaimedProduct).Value
      End Get
    End Property

    Public ReadOnly Property TotalNetAmount() As Double
      Get
        TotalNetAmount = mvTotalNetAmount
      End Get
    End Property

    Public Sub SetClaimDetails(ByVal pClaimTaxYearStart As String, ByVal pClaimedProduct As String, ByVal pUnclaimedProduct As String)
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimTaxYearStart).Value = pClaimTaxYearStart
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfClaimedProduct).Value = pClaimedProduct
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfUnclaimedProduct).Value = pUnclaimedProduct
    End Sub

    Friend Sub UpdateClaimAmounts(ByVal pNetAmount As Double, ByVal pAmountClaimed As Double, ByVal pTaxPercent As Integer)
      'CalculatedTaxAmount = tax amount claimed on each line
      'AmountClaimed = tax% on total net amounts of all lines (mvTotalNetAmount)
      mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfCalculatedTaxAmount).Value = CStr(mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfCalculatedTaxAmount).DoubleValue + pAmountClaimed)
      mvTotalNetAmount = mvTotalNetAmount + pNetAmount
      mvTaxPercent = pTaxPercent
    End Sub

    Public Sub CalculateTaxToReclaim()
      If TotalNetAmount <> 0 Then
        mvClassFields.Item(GaSponsorshipTaxClaimFields.gstcfAmountClaimed).Value = CStr(FixTwoPlaces(TotalNetAmount * (mvTaxPercent / (100 - mvTaxPercent))))
      End If
    End Sub
  End Class
End Namespace
