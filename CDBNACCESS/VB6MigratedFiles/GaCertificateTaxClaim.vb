Namespace Access
  Public Class GaCertificateTaxClaim

    Public Enum GaCertificateTaxClaimRecordSetTypes 'These are bit values
      gctcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GaCertificateTaxClaimFields
      gctcfAll = 0
      gctcfClaimNumber
      gctcfClaimGDate
      gctcfClaimTYStart
      gctcfAmountClaimed
      gctcfHigherRAClaimed
      gctcfAmountPaid
      gctcfHigherRAPaid
      gctcfAmendedBy
      gctcfAmendedOn
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
          .DatabaseTableName = "ga_certificate_tax_claims"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("claim_number", CDBField.FieldTypes.cftLong)
          .Add("claim_generated_date", CDBField.FieldTypes.cftDate)
          .Add("claim_tax_year_start", CDBField.FieldTypes.cftDate)
          .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
          .Add("higher_rate_amount_claimed", CDBField.FieldTypes.cftNumeric)
          .Add("amount_paid", CDBField.FieldTypes.cftNumeric)
          .Add("higher_rate_amount_paid", CDBField.FieldTypes.cftNumeric)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimNumber).SetPrimaryKeyOnly()

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults(Optional ByVal pClaimNumber As Integer = 0)
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimNumber).Value = CStr(pClaimNumber)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimGDate).Value = TodaysDate()
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimTYStart).Value = TodaysDate()
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountClaimed).Value = CStr(0)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAClaimed).Value = CStr(0)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountPaid).Value = CStr(0)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAPaid).Value = CStr(0)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmendedBy).Value = mvEnv.User.Logname
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmendedOn).Value = TodaysDate()

    End Sub

    Private Sub SetValid(ByVal pField As GaCertificateTaxClaimFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GaCertificateTaxClaimRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GaCertificateTaxClaimRecordSetTypes.gctcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "gtc")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pClaimNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      SetDefaults(pClaimNumber)

      If pClaimNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GaCertificateTaxClaimRecordSetTypes.gctcrtAll) & " FROM ga_certificate_tax_claims gtc WHERE gtc.claim_number = " & pClaimNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GaCertificateTaxClaimRecordSetTypes.gctcrtAll)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GaCertificateTaxClaimRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And GaCertificateTaxClaimRecordSetTypes.gctcrtAll) = GaCertificateTaxClaimRecordSetTypes.gctcrtAll Then
          .SetItem(GaCertificateTaxClaimFields.gctcfClaimNumber, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfClaimGDate, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfClaimTYStart, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfAmountClaimed, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfHigherRAClaimed, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfAmountPaid, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfHigherRAPaid, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfAmendedBy, vFields)
          .SetItem(GaCertificateTaxClaimFields.gctcfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(GaCertificateTaxClaimFields.gctcfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Update(ByVal pParams As CDBParameters)
      With mvClassFields
        If pParams.Exists("TaxStatus") Then
          If pParams("TaxStatus").Value = "S" Then
            If pParams.Exists("AmountPaid") Then .Item(GaCertificateTaxClaimFields.gctcfAmountPaid).Value = CStr(pParams("AmountPaid").DoubleValue)
          Else
            If pParams.Exists("HigherAmountPaid") Then .Item(GaCertificateTaxClaimFields.gctcfHigherRAPaid).Value = CStr(pParams("HigherAmountPaid").DoubleValue)
          End If
        End If
      End With
    End Sub

    Public Function GetAmountPaid(ByVal pOldAmount As Double, ByVal pNewAmount As Double, ByVal pTaxStatus As String) As Double
      If pTaxStatus = "S" Then
        GetAmountPaid = CDbl(mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountPaid).Value) - pOldAmount + pNewAmount
      Else
        GetAmountPaid = CDbl(mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAPaid).Value) - pOldAmount + pNewAmount
      End If
    End Function
    Public Sub SetTaxYearStart(ByVal pTaxYearStart As Date)
      mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimTYStart).Value = CStr(pTaxYearStart)
    End Sub

    Public Function SetClaimAmounts(ByVal pCertificateAmount As Double, ByVal pTaxPercent As Integer, ByVal pHRTaxPercent As Integer, ByVal pTaxStatus As String) As Double
      Dim vAmountClaimed As Double
      Dim vHRAmountClaimed As Double

      If pTaxStatus = "H" Then
        vHRAmountClaimed = FixTwoPlaces(pCertificateAmount * (pHRTaxPercent / (100 - pHRTaxPercent)))
        mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAClaimed).DoubleValue = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAClaimed).DoubleValue + vHRAmountClaimed
        mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAPaid).DoubleValue = HigherRAClaimed
        SetClaimAmounts = vHRAmountClaimed
      Else
        vAmountClaimed = FixTwoPlaces(pCertificateAmount * (pTaxPercent / (100 - pTaxPercent)))
        mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountClaimed).Value = CStr(mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountClaimed).DoubleValue + vAmountClaimed)
        mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountPaid).DoubleValue = AmountClaimed
        SetClaimAmounts = vAmountClaimed
      End If
    End Function
    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property ClaimNumber() As Integer
      Get
        ClaimNumber = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ClaimGDate() As Date
      Get
        ClaimGDate = CDate(mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimGDate).Value)
      End Get
    End Property
    Public ReadOnly Property ClaimTYStart() As Date
      Get
        ClaimTYStart = CDate(mvClassFields.Item(GaCertificateTaxClaimFields.gctcfClaimTYStart).Value)
      End Get
    End Property
    Public ReadOnly Property AmountClaimed() As Double
      Get
        AmountClaimed = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountClaimed).DoubleValue
      End Get
    End Property
    Public ReadOnly Property HigherRAClaimed() As Double
      Get
        HigherRAClaimed = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAClaimed).DoubleValue
      End Get
    End Property
    Public ReadOnly Property AmountPaid() As Double
      Get
        AmountPaid = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfAmountPaid).DoubleValue
      End Get
    End Property
    Public ReadOnly Property HigherRAPaid() As Double
      Get
        HigherRAPaid = mvClassFields.Item(GaCertificateTaxClaimFields.gctcfHigherRAPaid).DoubleValue
      End Get
    End Property
  End Class
End Namespace
