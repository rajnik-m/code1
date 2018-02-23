Imports System.Collections.Specialized
Imports System.Web
Imports System.Xml
Imports System.Xml.Linq

Friend Class VSeriesAccountValidation : Implements IBankAccountValidation
  Private Enum VSeriesInvalidIssueCodes
    Valid = 0
    SortCodeInvalidLength
    SortCodeInvalidMisplacedCharacters
    AccountNumberTooShort
    AccountNumberTooLong
    AccountNumberInvalidMisplacedCharacters   '5
    ModulusCheckFailed
    SortCodeNotFound
    AccountNumberTooShortForTranscription
    AccountNumberTooLongForTranscription
    AccountNumberInvalidPrefixForTranscription    '10
    AccountNumberInvalidFormatForTranscription
    MissingBuildingSocietyRollNumber
    InvalidBuildingSocietyRollNumber
  End Enum

  Private mvEnv As CDBEnvironment

  'These are required for the Properties only
  Private mvAccountNumberOutput As String = String.Empty
  Private mvBankBICOutput As String = String.Empty
  Private mvBankName As String = String.Empty
  Private mvBranchAddressLine1 As String = String.Empty
  Private mvBranchAddressLine2 As String = String.Empty
  Private mvBranchAddressLine3 As String = String.Empty
  Private mvBranchAddressLine4 As String = String.Empty
  Private mvBranchBICOutput As String = String.Empty
  Private mvBranchCountryDesc As String = String.Empty
  Private mvBranchCounty As String = String.Empty
  Private mvBranchName As String = String.Empty
  Private mvBranchPostCode As String = String.Empty
  Private mvBranchTelephone As String = String.Empty
  Private mvBranchTitle As String = String.Empty
  Private mvBranchTown As String = String.Empty
  Private mvIBANOutput As String = String.Empty
  Private mvInvalidReasonCode As String = String.Empty
  Private mvInvalidReasonDesc As String = String.Empty
  Private mvInvalidParameterName As String = String.Empty
  Private mvIsValid As Boolean = False
  Private mvSortCodeOutput As String = String.Empty
  Private mvVerifyResult As AccountNoVerify.AccountNoVerifyResult = AccountNoVerify.AccountNoVerifyResult.avrvalid
  Private mvVerifyURL As String = String.Empty

  Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pSortCode As String, ByVal pAccountNumber As String, ByVal pVerifyType As AccountNoVerify.UseVerifyType)
    mvEnv = pEnv
    SortCodeInput = pSortCode
    AccountNumberInput = pAccountNumber
    VerifyType = pVerifyType
    VerifyURL = pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAccountValidationURL)
  End Sub

  Friend Sub ValidateBankAccount() Implements IBankAccountValidation.ValidateBankAccount
    Try
      If SortCodeInput.Length = 0 OrElse VerifyType = AccountNoVerify.UseVerifyType.uvtNone Then
        'No validation required
        IsValid = True
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrvalid

      ElseIf VerifyURL.Length = 0 Then
        'URL is not set so just error
        IsValid = False
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid
        RaiseError(DataAccessErrors.daeAccountValidationURLNotSet)

      Else
        'OK - validate the SortCode & AccountNumber
        IsValid = True
        InvalidReasonCode = String.Empty
        InvalidReasonDesc = String.Empty
        InvalidParameterName = String.Empty
        If VerifyURL.EndsWith("/") = False Then VerifyURL &= "/"

        'Add SortCode & AccountNumber to a collection and encode to UTF-8
        Dim vColl As New NameValueCollection
        vColl.Add("sortCode", SortCodeInput)
        vColl.Add("accountNumber", AccountNumberInput)
        Dim vQueryString As NameValueCollection = HttpUtility.ParseQueryString(String.Empty)
        vQueryString.Add(vColl)

        'Build the URL
        Dim vURL As New StringBuilder
        With vURL
          .Append(VerifyURL)
          .Append("api/ValidateUKBankAccount?")
          .Append(vQueryString.ToString)
        End With

        'Send the request
        Dim vResponse As System.Net.WebResponse
        Dim vHttpRequest As System.Net.WebRequest = System.Net.WebRequest.Create(vURL.ToString)
        vResponse = vHttpRequest.GetResponse
        If vResponse IsNot Nothing Then
          Dim vStream As IO.Stream = vResponse.GetResponseStream()
          Dim vReader As New IO.StreamReader(vStream)

          Dim vXMLDocument As XmlDocument = JsonToXML(vReader.ReadToEnd)
          Dim vNodeList As XmlNodeList = vXMLDocument.GetElementsByTagName("Valid")
          If vNodeList IsNot Nothing Then
            For Each vNode As XmlNode In vNodeList
              If Not String.IsNullOrEmpty(vNode.InnerText) Then IsValid = CBool(vNode.InnerText)
            Next
          End If

          If IsValid Then
            '
          Else
            SetErrorInfo(vXMLDocument)
          End If

          'Set AccountDetails output even if invalid
          vNodeList = vXMLDocument.GetElementsByTagName("AccountDetailsOutput")
          If vNodeList IsNot Nothing Then
            For Each vNode As XmlNode In vNodeList
              If vNode.ChildNodes.Count > 1 Then
                For Each vChildNode As XmlNode In vNode.ChildNodes
                  Select Case vChildNode.Name.ToLower
                    Case "accountnumber"
                      AccountNumberOutput = vChildNode.InnerText.Trim
                    Case "sortcode"
                      SortCodeOutput = vChildNode.InnerText.Trim
                  End Select
                Next
              End If
            Next
          End If
          vNodeList = vXMLDocument.GetElementsByTagName("AccountIBAN")
          If vNodeList IsNot Nothing Then
            For Each vNode As XmlNode In vNodeList
              If Not String.IsNullOrEmpty(vNode.InnerText) Then IBANOutput = vNode.InnerText.Trim
            Next
          End If
          SetBankBranchInfo(vXMLDocument)
        End If
      End If

    Catch vEX As Exception
      IsValid = False
      Throw vEX
    End Try
  End Sub

  Private Function JsonToXML(ByVal vJson As String) As XmlDocument
    Dim vDoc As New XmlDocument()
    Using vReader = Runtime.Serialization.Json.JsonReaderWriterFactory.CreateJsonReader(Encoding.UTF8.GetBytes(vJson), XmlDictionaryReaderQuotas.Max)
      Dim vXml As XElement = XElement.Load(vReader)
      vDoc.LoadXml(vXml.ToString())
    End Using

    Return vDoc
  End Function

  Private Sub SetErrorInfo(ByVal pXMLDoc As XmlDocument)
    'Set InvalidReasonCode, InvalidReasonDesc and InvalidParameter
    Dim vNodeList As XmlNodeList = pXMLDoc.GetElementsByTagName("InvalidIssue")
    If vNodeList IsNot Nothing Then
      For Each vNode As XmlNode In vNodeList
        If Not String.IsNullOrEmpty(vNode.InnerText) Then InvalidReasonCode = vNode.InnerText.Trim()
      Next
    End If

    vNodeList = pXMLDoc.GetElementsByTagName("InvalidReason")
    If vNodeList IsNot Nothing Then
      For Each vNode As XmlNode In vNodeList
        If Not String.IsNullOrEmpty(vNode.InnerText) Then InvalidReasonDesc = vNode.InnerText.Trim()
      Next
    End If

    vNodeList = pXMLDoc.GetElementsByTagName("InvalidParameter")
    If vNodeList IsNot Nothing Then
      For Each vNode As XmlNode In vNodeList
        If Not String.IsNullOrEmpty(vNode.InnerText) Then InvalidParameterName = vNode.InnerText.Trim()
      Next
    End If

    'Using the InvalidReasonCode, set the VerifyResult
    Select Case CType(IntegerValue(InvalidReasonCode), VSeriesInvalidIssueCodes)
      Case VSeriesInvalidIssueCodes.Valid
        'Valid - we shound never come in here with this
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrvalid
      Case VSeriesInvalidIssueCodes.SortCodeInvalidLength, VSeriesInvalidIssueCodes.SortCodeInvalidMisplacedCharacters,
           VSeriesInvalidIssueCodes.SortCodeNotFound
        'Sort Code invalid
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid
      Case VSeriesInvalidIssueCodes.AccountNumberTooShort, VSeriesInvalidIssueCodes.AccountNumberTooShortForTranscription
        'Sort Code valid but Account Number too short
        If AccountNumberInput.Length = 0 Then
          'We were just validating Sort Code so valid
          VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrvalid
        Else
          VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrSortcodeValidAccountInvalid
        End If
      Case VSeriesInvalidIssueCodes.AccountNumberTooLong, VSeriesInvalidIssueCodes.AccountNumberInvalidMisplacedCharacters,
           VSeriesInvalidIssueCodes.AccountNumberTooLongForTranscription, VSeriesInvalidIssueCodes.AccountNumberInvalidPrefixForTranscription,
           VSeriesInvalidIssueCodes.AccountNumberInvalidFormatForTranscription
        'Sort Code valid and Account Number invalid
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrSortcodeValidAccountInvalid
      Case VSeriesInvalidIssueCodes.ModulusCheckFailed
        'Sort Code and Account Number combination invalid
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid
      Case VSeriesInvalidIssueCodes.MissingBuildingSocietyRollNumber, VSeriesInvalidIssueCodes.InvalidBuildingSocietyRollNumber
        'Roll Number invalid / too short - we do not support Roll Number!!
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid
      Case Else
        'We got some other unexpected error code from V-Series
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid
    End Select

    If VerifyType = AccountNoVerify.UseVerifyType.uvtWarn Then
      If VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrInvalid Then
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrWarning
      ElseIf VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrSortcodeValidAccountInvalid Then
        VerifyResult = AccountNoVerify.AccountNoVerifyResult.avrSortcodeValidAccountWarn
      End If
    End If
  End Sub

  Private Sub SetBankBranchInfo(ByVal pXMLDoc As XmlDocument)
    BankBICOutput = String.Empty
    BankName = String.Empty
    BranchBICOutput = String.Empty
    BranchTitle = String.Empty
    BranchName = String.Empty
    BranchAddressLine1 = String.Empty
    BranchAddressLine2 = String.Empty
    BranchAddressLine3 = String.Empty
    BranchAddressLine4 = String.Empty
    BranchTown = String.Empty
    BranchCounty = String.Empty
    BranchPostCode = String.Empty
    BranchCountryDesc = String.Empty
    BranchTelephone = String.Empty

    Dim vNodeList As XmlNodeList = pXMLDoc.GetElementsByTagName("UKBankBranch")
    If vNodeList IsNot Nothing Then
      For Each vNode As XmlNode In vNodeList
        If vNode.ChildNodes.Count > 1 Then
          For Each vChildNode As XmlNode In vNode.ChildNodes
            If Not (String.IsNullOrEmpty(vChildNode.Name)) AndAlso Not (String.IsNullOrEmpty(vChildNode.InnerText)) Then
              Select Case vChildNode.Name.ToLower
                Case "bankbic"
                  BankBICOutput = vChildNode.InnerText.Trim
                Case "bankname"
                  BankName = vChildNode.InnerText.Trim
                Case "branchbic"
                  BranchBICOutput = vChildNode.InnerText.Trim
                Case "branchname"
                  BranchName = vChildNode.InnerText.Trim
                Case "contactaddress1"
                  BranchAddressLine1 = vChildNode.InnerText.Trim
                Case "contactaddress2"
                  BranchAddressLine2 = vChildNode.InnerText.Trim
                Case "contactaddress3"
                  BranchAddressLine3 = vChildNode.InnerText.Trim
                Case "contactaddress4"
                  BranchAddressLine4 = vChildNode.InnerText.Trim
                Case "contactaddresscity"
                  BranchTown = vChildNode.InnerText.Trim
                Case "contactaddresscounty"
                  BranchCounty = vChildNode.InnerText.Trim
                Case "contactaddresspostcode"
                  BranchPostCode = vChildNode.InnerText.Trim
                Case "contactaddresspostcountry"
                  BranchCountryDesc = vChildNode.InnerText.Trim
                Case "contactphonenumber"
                  BranchTelephone = vChildNode.InnerText.Trim
                Case "officetitle"
                  BranchTitle = vChildNode.InnerText.Trim
              End Select
            End If
          Next
        End If
      Next
    End If

    If SortCodeOutput.Length > 0 AndAlso BankName.Length > 0 Then
      Dim vBank As New Bank()
      vBank.Init(mvEnv, SortCodeOutput)
      If vBank.Existing = False Then
        'Create a new record
        Dim vParams As New CDBParameters()
        With vParams
          .Add("SortCode", SortCodeOutput)
          Dim vBankName As String = StrConv(Common.Substring(BankName, 0, 30), VbStrConv.ProperCase)    'Data is in upper case, so use correct case
          If vBankName.Contains(" Tsb ") Then vBankName = vBankName.Replace(" Tsb ", " TSB ") 'It's an abbreviation so should be upper case
          .Add("Bank", vBankName)
          .Add("BranchName", Common.Substring(If(BranchTitle.Length > 0, BranchTitle, BranchName), 0, 30))
          Dim vAddressData As New StringBuilder()
          If BranchAddressLine1.Length > 0 Then vAddressData.AppendLine(BranchAddressLine1)
          If BranchAddressLine2.Length > 0 Then vAddressData.AppendLine(BranchAddressLine2)
          If BranchAddressLine3.Length > 0 Then vAddressData.AppendLine(BranchAddressLine3)
          If BranchAddressLine4.Length > 0 Then vAddressData.AppendLine(BranchAddressLine4)
          .Add("Address", CDBField.FieldTypes.cftMemo, vAddressData.ToString)
          .Add("Town", Common.Substring(BranchTown, 0, 35))
          .Add("County", Common.Substring(BranchCounty, 0, 35))
          .Add("Postcode", Common.Substring(BranchPostCode, 0, 10))
        End With
        vBank.Create(vParams)
        vBank.Save(mvEnv.User.UserID, True)
      End If
    End If

  End Sub

#Region " Friend Read/Write Properties "

  ''' <summary>Sort Code to be validated.</summary>
  Friend Property SortCodeInput As String = String.Empty Implements IBankAccountValidation.SortCodeInput
  ''' <summary>Account Number to be validated.</summary>
  Friend Property AccountNumberInput As String = String.Empty Implements IBankAccountValidation.AccountNumberInput
  ''' <summary>The type of verification to be applied.  The default is None.</summary>
  Friend Property VerifyType As AccountNoVerify.UseVerifyType = AccountNoVerify.UseVerifyType.uvtNone Implements IBankAccountValidation.VerifyType

#End Region

#Region " Friend Read / Private Write Properties "

  ''' <summary>The URL required for Account Validation.</summary>
  Friend Property VerifyURL As String Implements IBankAccountValidation.VerifyURL
    Get
      Return mvVerifyURL
    End Get
    Private Set(value As String)
      mvVerifyURL = value
    End Set
  End Property

  Friend Property VerifyResult As AccountNoVerify.AccountNoVerifyResult Implements IBankAccountValidation.VerifyResult
    Get
      Return mvVerifyResult
    End Get
    Private Set(value As AccountNoVerify.AccountNoVerifyResult)
      mvVerifyResult = value
    End Set
  End Property

  Friend Property IsValid As Boolean Implements IBankAccountValidation.IsValid
    Get
      Return mvIsValid
    End Get
    Private Set(value As Boolean)
      mvIsValid = value
    End Set
  End Property

  Friend Property InvalidReasonCode As String Implements IBankAccountValidation.InvalidReasonCode
    Get
      Return mvInvalidReasonCode
    End Get
    Private Set(value As String)
      mvInvalidReasonCode = value
    End Set
  End Property

  Friend Property InvalidReasonDesc As String Implements IBankAccountValidation.InvalidReasonDesc
    Get
      Return mvInvalidReasonDesc
    End Get
    Private Set(value As String)
      mvInvalidReasonDesc = value
    End Set
  End Property

  Friend Property InvalidParameterName As String Implements IBankAccountValidation.InvalidParameterName
    Get
      Return mvInvalidParameterName
    End Get
    Private Set(value As String)
      mvInvalidParameterName = value
    End Set
  End Property

  Friend Property SortCodeOutput As String Implements IBankAccountValidation.SortCodeOutput
    Get
      Return mvSortCodeOutput
    End Get
    Private Set(value As String)
      mvSortCodeOutput = value
    End Set
  End Property

  Friend Property AccountNumberOutput As String Implements IBankAccountValidation.AccountNumberOutput
    Get
      Return mvAccountNumberOutput
    End Get
    Private Set(value As String)
      mvAccountNumberOutput = value
    End Set
  End Property

  Friend Property IBANOutput As String Implements IBankAccountValidation.IBANOutput
    Get
      Return mvIBANOutput
    End Get
    Private Set(value As String)
      mvIBANOutput = value
    End Set
  End Property

  Friend Property BankBICOutput As String Implements IBankAccountValidation.BankBICOutput
    Get
      Return mvBankBICOutput
    End Get
    Private Set(value As String)
      mvBankBICOutput = value
    End Set
  End Property

  Friend Property BranchBICOutput As String Implements IBankAccountValidation.BranchBICOutput
    Get
      Return mvBranchBICOutput
    End Get
    Private Set(value As String)
      mvBranchBICOutput = value
    End Set
  End Property

  Friend Property BankName As String Implements IBankAccountValidation.BankName
    Get
      Return mvBankName
    End Get
    Private Set(value As String)
      mvBankName = value
    End Set
  End Property

  Friend Property BranchTitle As String Implements IBankAccountValidation.BranchTitle
    Get
      Return mvBranchTitle
    End Get
    Private Set(value As String)
      mvBranchTitle = value
    End Set
  End Property

  Friend Property BranchName As String Implements IBankAccountValidation.BranchName
    Get
      Return mvBranchName
    End Get
    Private Set(value As String)
      mvBranchName = value
    End Set
  End Property

  Friend Property BranchAddressLine1 As String Implements IBankAccountValidation.BranchAddressLine1
    Get
      Return mvBranchAddressLine1
    End Get
    Private Set(value As String)
      mvBranchAddressLine1 = value
    End Set
  End Property

  Friend Property BranchAddressLine2 As String Implements IBankAccountValidation.BranchAddressLine2
    Get
      Return mvBranchAddressLine2
    End Get
    Private Set(value As String)
      mvBranchAddressLine2 = value
    End Set
  End Property

  Friend Property BranchAddressLine3 As String Implements IBankAccountValidation.BranchAddressLine3
    Get
      Return mvBranchAddressLine3
    End Get
    Private Set(value As String)
      mvBranchAddressLine3 = value
    End Set
  End Property

  Friend Property BranchAddressLine4 As String Implements IBankAccountValidation.BranchAddressLine4
    Get
      Return mvBranchAddressLine4
    End Get
    Private Set(value As String)
      mvBranchAddressLine4 = value
    End Set
  End Property

  Friend Property BranchTown As String Implements IBankAccountValidation.BranchTown
    Get
      Return mvBranchTown
    End Get
    Private Set(value As String)
      mvBranchTown = value
    End Set
  End Property

  Friend Property BranchCounty As String Implements IBankAccountValidation.BranchCounty
    Get
      Return mvBranchCounty
    End Get
    Private Set(value As String)
      mvBranchCounty = value
    End Set
  End Property

  Friend Property BranchPostCode As String Implements IBankAccountValidation.BranchPostCode
    Get
      Return mvBranchPostCode
    End Get
    Private Set(value As String)
      mvBranchPostCode = value
    End Set
  End Property

  Friend Property BranchCountryDesc As String Implements IBankAccountValidation.BranchCountryDesc
    Get
      Return mvBranchCountryDesc
    End Get
    Private Set(value As String)
      mvBranchCountryDesc = value
    End Set
  End Property

  Friend Property BranchTelephone As String Implements IBankAccountValidation.BranchTelephone
    Get
      Return mvBranchTelephone
    End Get
    Private Set(value As String)
      mvBranchTelephone = value
    End Set
  End Property

#End Region

End Class
