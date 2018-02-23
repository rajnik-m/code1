

Namespace Access
  Public Class ContactCreditCard

    Public Enum ContactCreditCardRecordSetTypes 'These are bit values
      cccrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ContactCreditCardFields
      cccfall = 0
      cccfCreditCardDetailsNumber
      cccfCreditCardNumber
      cccfContactNumber
      cccfExpiryDate
      cccfIssuer
      cccfAccountName
      cccfCreditCardType
      cccfIssueNumber
      cccfTokenDesc
      cccfTokenId
      cccfAmendedBy
      cccfAmendedOn
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
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "contact_credit_cards"
          .Add("credit_card_details_number", CDBField.FieldTypes.cftLong)
          .Add("credit_card_number")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("expiry_date")
          .Add("issuer")
          .Add("account_name")
          .Add("credit_card_type")
          .Add("issue_number", CDBField.FieldTypes.cftInteger)
          .Add("token_desc")
          .Add("token_id")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(ContactCreditCardFields.cccfCreditCardDetailsNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(ContactCreditCardFields.cccfIssueNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCCIssueNumber)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As ContactCreditCardFields)
      'Add code here to ensure all values are valid before saving
      If pField = ContactCreditCardFields.cccfall Then
        mvClassFields.Item(ContactCreditCardFields.cccfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(ContactCreditCardFields.cccfAmendedBy).Value = mvEnv.User.UserID
        If Len(mvClassFields.Item(ContactCreditCardFields.cccfCreditCardDetailsNumber).Value) = 0 Then mvClassFields.Item(ContactCreditCardFields.cccfCreditCardDetailsNumber).Value = CStr(mvEnv.GetControlNumber("CC"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ContactCreditCardRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ContactCreditCardRecordSetTypes.cccrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ccc")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCreditCardDetailsNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      If pCreditCardDetailsNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactCreditCardRecordSetTypes.cccrtAll) & " FROM contact_credit_cards WHERE credit_card_details_number = " & pCreditCardDetailsNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ContactCreditCardRecordSetTypes.cccrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactCreditCardRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ContactCreditCardFields.cccfCreditCardDetailsNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ContactCreditCardRecordSetTypes.cccrtAll) = ContactCreditCardRecordSetTypes.cccrtAll Then
          .SetItem(ContactCreditCardFields.cccfCreditCardNumber, vFields)
          .SetItem(ContactCreditCardFields.cccfContactNumber, vFields)
          .SetItem(ContactCreditCardFields.cccfExpiryDate, vFields)
          .SetItem(ContactCreditCardFields.cccfIssuer, vFields)
          .SetItem(ContactCreditCardFields.cccfAccountName, vFields)
          .SetItem(ContactCreditCardFields.cccfCreditCardType, vFields)
          .SetOptionalItem(ContactCreditCardFields.cccfIssueNumber, vFields)
          .SetItem(ContactCreditCardFields.cccfAmendedBy, vFields)
          .SetItem(ContactCreditCardFields.cccfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(ContactCreditCardFields.cccfall)
      mvClassFields.Save(mvEnv, mvExisting)
    End Sub

    Public Sub Create(ByRef pContactNumber As Integer, ByRef pCardNumber As String, ByRef pExpiryDate As String, ByRef pIssuer As String, ByRef pAccountName As String, ByRef pCreditCardType As String, Optional ByRef pIssueNumber As String = "")
      With mvClassFields
        .Item(ContactCreditCardFields.cccfContactNumber).IntegerValue = pContactNumber
        .Item(ContactCreditCardFields.cccfExpiryDate).Value = Replace(pExpiryDate, "/", "")
        .Item(ContactCreditCardFields.cccfIssuer).Value = pIssuer
        .Item(ContactCreditCardFields.cccfAccountName).Value = pAccountName
        .Item(ContactCreditCardFields.cccfCreditCardType).Value = pCreditCardType
        If Len(pIssueNumber) > 0 Then .Item(ContactCreditCardFields.cccfIssueNumber).Value = pIssueNumber
      End With
      SetValid(ContactCreditCardFields.cccfall)
    End Sub

    Public Sub Update(ByRef pCardNumber As String, ByRef pExpiryDate As String, ByRef pIssuer As String, ByRef pAccountName As String, ByRef pCreditCardType As String, Optional ByRef pIssueNumber As String = "")
      Update(pCardNumber, pExpiryDate, pIssueNumber, pAccountName, pCreditCardType, pIssueNumber, String.Empty)
    End Sub

    Public Sub Update(ByRef pCardNumber As String, ByRef pExpiryDate As String, ByRef pIssuer As String, ByRef pAccountName As String, ByRef pCreditCardType As String, ByRef pIssueNumber As String, ByRef pTokenDesc As String)
      With mvClassFields
        .Item(ContactCreditCardFields.cccfExpiryDate).Value = Replace(pExpiryDate, "/", "")
        .Item(ContactCreditCardFields.cccfIssuer).Value = pIssuer
        .Item(ContactCreditCardFields.cccfAccountName).Value = pAccountName
        .Item(ContactCreditCardFields.cccfCreditCardType).Value = pCreditCardType
        If Len(pIssueNumber) > 0 Then .Item(ContactCreditCardFields.cccfIssueNumber).Value = pIssueNumber
        If Not String.IsNullOrEmpty(pTokenDesc) Then .Item(ContactCreditCardFields.cccfTokenDesc).Value = pTokenDesc
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

    Public ReadOnly Property AccountName() As String
      Get
        AccountName = mvClassFields.Item(ContactCreditCardFields.cccfAccountName).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(ContactCreditCardFields.cccfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ContactCreditCardFields.cccfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CreditCardDetailsNumber() As Integer
      Get
        SetValid(ContactCreditCardFields.cccfCreditCardDetailsNumber)
        CreditCardDetailsNumber = mvClassFields.Item(ContactCreditCardFields.cccfCreditCardDetailsNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ContactCreditCardFields.cccfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreditCardNumber() As String
      Get
        Return String.Empty
      End Get
    End Property

    Public ReadOnly Property CreditCardType() As String
      Get
        CreditCardType = mvClassFields.Item(ContactCreditCardFields.cccfCreditCardType).Value
      End Get
    End Property

    Public ReadOnly Property FormattedExpiryDate() As String
      Get
        FormattedExpiryDate = Left(ExpiryDate, 2) & "/" & Right(ExpiryDate, 2)
      End Get
    End Property
    Public ReadOnly Property ExpiryDate() As String
      Get
        ExpiryDate = mvClassFields.Item(ContactCreditCardFields.cccfExpiryDate).Value
      End Get
    End Property

    Public ReadOnly Property Issuer() As String
      Get
        Issuer = mvClassFields.Item(ContactCreditCardFields.cccfIssuer).Value
      End Get
    End Property

    Public ReadOnly Property IssueNumber() As String
      Get
        IssueNumber = mvClassFields.Item(ContactCreditCardFields.cccfIssueNumber).Value
      End Get
    End Property

    Public ReadOnly Property TokenDesc() As String
      Get
        Return mvClassFields.Item(ContactCreditCardFields.cccfTokenDesc).Value
      End Get
    End Property

    Public ReadOnly Property TokenId() As String
      Get
        Return mvClassFields.Item(ContactCreditCardFields.cccfTokenId).Value
      End Get
    End Property

    Public Function Delete(pCreditCardDetailsNumber As Integer) As Integer
      Return mvEnv.Connection.DeleteRecords("contact_credit_cards", New CDBFields(New CDBField(mvClassFields.Item(ContactCreditCardFields.cccfCreditCardDetailsNumber).Name, pCreditCardDetailsNumber.ToString)), False)
    End Function

  End Class
End Namespace
