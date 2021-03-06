Namespace Access

  Partial Public Class CardSale
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum CardSaleFields
      AllFields = 0
      BatchNumber
      TransactionNumber
      CardNumber
      IssueNumber
      ExpiryDate
      AuthorisationCode
      CreditCardDetailsNumber
      CreditCardType
      ValidDate
      NoClaimRequired
      SecurityCode
      TemplateNumber
      ProtxCardType
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("batch_number", CDBField.FieldTypes.cftLong)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("card_number")
        .Add("issue_number", CDBField.FieldTypes.cftInteger)
        .Add("expiry_date")
        .Add("authorisation_code")
        .Add("credit_card_details_number", CDBField.FieldTypes.cftLong)
        .Add("credit_card_type")
        .Add("valid_date")
        .Add("no_claim_required")
        .Add("security_code")
        .Add("template_number")
        .Add("protx_card_type")
        .Item(CardSaleFields.BatchNumber).PrimaryKey = True

        .Item(CardSaleFields.TransactionNumber).PrimaryKey = True

        .Item(CardSaleFields.TemplateNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMerchantDetails)
        .Item(CardSaleFields.ProtxCardType).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cbdProtxCardType)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cs"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "card_sales"
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
    Public Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(CardSaleFields.BatchNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CardSaleFields.BatchNumber).IntegerValue = Value
      End Set
    End Property
    Public Property TransactionNumber() As Integer
      Get
        Return mvClassFields(CardSaleFields.TransactionNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CardSaleFields.TransactionNumber).IntegerValue = Value
      End Set
    End Property
    Public ReadOnly Property CardNumber() As String
      Get
        Return mvEnv.DecryptCreditCardNumber(mvClassFields(CardSaleFields.CardNumber).Value)
      End Get
    End Property
    Public Property IssueNumber() As String
      Get
        Return mvClassFields(CardSaleFields.IssueNumber).Value
      End Get
      Set(ByVal pValue As String)
        mvClassFields(CardSaleFields.IssueNumber).Value = pValue
      End Set
    End Property
    Public ReadOnly Property ExpiryDate() As String
      Get
        Return mvClassFields(CardSaleFields.ExpiryDate).Value
      End Get
    End Property
    Public ReadOnly Property AuthorisationCode() As String
      Get
        Return mvClassFields(CardSaleFields.AuthorisationCode).Value
      End Get
    End Property
    Public Property CreditCardDetailsNumber() As Integer
      Get
        Return mvClassFields(CardSaleFields.CreditCardDetailsNumber).IntegerValue
      End Get
      Set(ByVal pValue As Integer)
        mvClassFields(CardSaleFields.CreditCardDetailsNumber).IntegerValue = pValue
      End Set
    End Property
    Public Property CreditCardType() As String
      Get
        Return mvClassFields(CardSaleFields.CreditCardType).Value
      End Get
      Set(ByVal pValue As String)
        mvClassFields(CardSaleFields.CreditCardType).Value = pValue
      End Set
    End Property
    Public Property ProtXCardType() As String
      Get
        Return mvClassFields(CardSaleFields.ProtxCardType).Value
      End Get
      Set(ByVal pValue As String)
        mvClassFields(CardSaleFields.ProtxCardType).Value = pValue
      End Set
    End Property
    Public ReadOnly Property ValidDate() As String
      Get
        Return mvClassFields(CardSaleFields.ValidDate).Value
      End Get
    End Property
    Public Property NoClaimRequired() As Boolean
      Get
        Return mvClassFields(CardSaleFields.NoClaimRequired).Bool
      End Get
      Set(ByVal value As Boolean)
        mvClassFields(CardSaleFields.NoClaimRequired).Bool = value
      End Set
    End Property
    Public ReadOnly Property SecurityCode() As String
      Get
        Return mvClassFields(CardSaleFields.SecurityCode).Value
      End Get
    End Property
    Public Property TemplateNumber() As String
      Get
        Return mvClassFields(CardSaleFields.TemplateNumber).Value
      End Get
      Set(ByVal value As String)
        mvClassFields(CardSaleFields.TemplateNumber).Value = value
      End Set
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      If mvClassFields(CardSaleFields.NoClaimRequired).Value = "" Then mvClassFields(CardSaleFields.NoClaimRequired).Bool = False
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      mvClassFields.Item(CardSaleFields.SecurityCode).Value = ""    'Never save this to the database
      If TemplateNumber.Length > 0 Then 'Do not save the card details if we have a Template Number (i.e. using SecureCXL SCP)
        mvClassFields(CardSaleFields.CardNumber).Value = ""
        mvClassFields(CardSaleFields.CreditCardType).Value = ""
        mvClassFields(CardSaleFields.ExpiryDate).Value = ""
        mvClassFields(CardSaleFields.IssueNumber).Value = ""
        mvClassFields(CardSaleFields.ValidDate).Value = ""
      End If
    End Sub

    Public Overloads Sub Init(ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer)
      If pBatchNumber > 0 AndAlso pTransactionNumber > 0 Then
        CheckClassFields()
        Dim vWhereFields As New CDBFields
        vWhereFields.Add(mvClassFields(CardSaleFields.BatchNumber).Name, pBatchNumber)
        vWhereFields.Add(mvClassFields(CardSaleFields.TransactionNumber).Name, pTransactionNumber)
        InitWithPrimaryKey(vWhereFields)
      Else
        Init()
      End If
    End Sub

    Public Overloads Sub Update(ByVal pCardNumber As String, ByVal pIssueNumber As String, ByVal pValidDate As String, ByVal pExpiryDate As String, ByVal pAuthorisationCode As String, ByVal pCreditCardType As String, ByRef pNoClaimRequired As Boolean)
      mvClassFields.Item(CardSaleFields.CardNumber).Value = mvEnv.EncryptCreditCardNumber(pCardNumber)
      mvClassFields.Item(CardSaleFields.IssueNumber).Value = pIssueNumber
      mvClassFields.Item(CardSaleFields.ValidDate).Value = pValidDate
      mvClassFields.Item(CardSaleFields.ExpiryDate).Value = pExpiryDate
      mvClassFields.Item(CardSaleFields.AuthorisationCode).Value = pAuthorisationCode
      mvClassFields.Item(CardSaleFields.CreditCardType).Value = pCreditCardType
      mvClassFields.Item(CardSaleFields.NoClaimRequired).Bool = pNoClaimRequired
    End Sub

    Public Sub ClearCardNumber()
      mvClassFields.Item(CardSaleFields.CardNumber).Value = ""
    End Sub

    Public Sub SetTemplateNumber()
      If mvEnv.GetConfig("fp_cc_authorisation_type") = "SCXLVPCSCP" Then
        'This will not add a new CCA record
        Dim vCCA As New CreditCardAuthorisation
        vCCA.Init(mvEnv)
        vCCA.GetTemplateNumber(Me)
        If TemplateNumber.Length > 0 Then
          NoClaimRequired = True
          Save()
        End If
      End If
    End Sub
#End Region

  End Class
End Namespace
