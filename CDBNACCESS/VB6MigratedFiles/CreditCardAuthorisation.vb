Namespace Access
  Public Class CreditCardAuthorisation

    Public Enum CreditCardAuthorisationRecordSetTypes 'These are bit values
      ccaurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CreditCardAuthorisationFields
      ccafAll = 0
      ccafAuthorisationNumber
      ccafBatchNumber
      ccafTransactionNumber
      ccafAuthorisationType
      ccafAuthorisedOn
      ccafAuthorisationCode
      ccafAuthorisationResponseCode
      ccafAuthorisationResponseMessage
      ccafAuthorisedAmount
      ccafAddressVerificationResult
      ccafAddressVerificationMessage
      ccafCSCResultCode
      ccafCSCResultMessage
      ccafAuthorisedTransactionNumber
      ccafAuthorisedTextID
    End Enum

    Public Enum CreditCardAuthorisationTypes
      ccatNone = 0
      ccatNormal = 1
      ccatBackOrder = 2
      ccatNotional = 4
      ccatRefund = 8
      ccatWeb = 16
      ccatTemplateOnly = 32 'Used to save the card details and get the Template Number only. CCA record will not be saved.
    End Enum

    Public Enum OnlineAuthorisationTypes
      None = 1
      CommsXL     'Config fp_cc_authorisation_directory value is set to CSXL210FE
      SecureCXL   'Config fp_cc_authorisation_directory value is set to SCXLVPCSCP
      ProtX       'Config fp_cc_authorisation_directory value is set to ProtX 
      TnsHosted   'Config fp_cc_authorisation_directory value is set to TnsHosted
      SagePayHosted 'Config fp_cc_authorisation_directory value is set to SagePayHosted
    End Enum

    Public Event AuthorisingCreditCard(ByRef pMaxTime As Integer, ByRef pTime As Integer)

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvOnlineAuthorisationType As OnlineAuthorisationTypes = Nothing
    Private mvContactNumber As Integer

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "credit_card_authorisations"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("authorisation_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("authorisation_type")
          .Add("authorised_on", CDBField.FieldTypes.cftTime)
          .Add("authorisation_code")
          .Add("authorisation_response_code")
          .Add("authorisation_response_message")
          .Add("authorised_amount", CDBField.FieldTypes.cftNumeric)
          .Add("address_verification_result")
          .Add("address_verification_message")
          .Add("csc_result_code")
          .Add("csc_result_message")
          .Add("authorised_transaction_number")
          .Add("authorised_text_id")
        End With

        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAddressVerificationResult).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCreditCardAVSCVV2)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMerchantDetails) = False Then
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAddressVerificationMessage).InDatabase = False
          mvClassFields.Item(CreditCardAuthorisationFields.ccafCSCResultCode).InDatabase = False
          mvClassFields.Item(CreditCardAuthorisationFields.ccafCSCResultMessage).InDatabase = False
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).InDatabase = False
        End If
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAuthorisedTextID) = False Then mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisedTextID).InDatabase = False

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CreditCardAuthorisationFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CreditCardAuthorisationRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CreditCardAuthorisationRecordSetTypes.ccaurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ccau")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAuthorisationNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pAuthorisationNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CreditCardAuthorisationRecordSetTypes.ccaurtAll) & " FROM credit_card_authorisations ccau WHERE authorisation_number = " & pAuthorisationNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CreditCardAuthorisationRecordSetTypes.ccaurtAll)
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

    Public Sub InitFromTransaction(ByVal pEnv As CDBEnvironment, ByVal pBatchNo As Integer, ByVal pTransNo As Integer)
      InitFromTransaction(pEnv, pBatchNo, pTransNo, False)
    End Sub
    Public Sub InitFromTransaction(ByVal pEnv As CDBEnvironment, ByVal pBatchNo As Integer, ByVal pTransNo As Integer, ByVal pBackOrderOnly As Boolean)
      mvEnv = pEnv

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("batch_number", pBatchNo)
      vWhereFields.Add("transaction_number", pTransNo)
      Dim vOrderBy As String = ""
      If pBackOrderOnly Then
        'Currently used in ConfirmedStockAllocation.
        vWhereFields.Add("authorisation_type", "B")
        vOrderBy = "authorisation_number DESC"  'Sort it with DESC as there might be old records of type 'B' for the same batch and trans
      Else
        vWhereFields.Add("authorisation_type", "A")
      End If
      Dim vRecordSet As CDBRecordSet = New SQLStatement(pEnv.Connection, GetRecordSetFields(CreditCardAuthorisationRecordSetTypes.ccaurtAll), "credit_card_authorisations ccau", vWhereFields, vOrderBy).GetRecordSet
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, CreditCardAuthorisationRecordSetTypes.ccaurtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CreditCardAuthorisationRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CreditCardAuthorisationFields.ccafAuthorisationNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CreditCardAuthorisationRecordSetTypes.ccaurtAll) = CreditCardAuthorisationRecordSetTypes.ccaurtAll Then
          .SetItem(CreditCardAuthorisationFields.ccafBatchNumber, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafTransactionNumber, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisationType, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisedOn, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisationCode, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisationResponseCode, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage, vFields)
          .SetItem(CreditCardAuthorisationFields.ccafAuthorisedAmount, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafAddressVerificationResult, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafAddressVerificationMessage, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafCSCResultCode, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafCSCResultMessage, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber, vFields)
          .SetOptionalItem(CreditCardAuthorisationFields.ccafAuthorisedTextID, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CreditCardAuthorisationFields.ccafAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Private Function OnlineAuthorisationType() As OnlineAuthorisationTypes
      Return OnlineAuthorisationType(False)
    End Function

    Private Function OnlineAuthorisationType(ByVal pForceSecureCXL As Boolean) As OnlineAuthorisationTypes
      'pForceSecureCXL will override the config value and always used SCXLVPCSCP (for refund and subsequent payments)
      If mvOnlineAuthorisationType = 0 Then
        If pForceSecureCXL And mvEnv.GetConfig("fp_cc_authorisation_type") <> "PROTX" Then
          mvOnlineAuthorisationType = OnlineAuthorisationTypes.SecureCXL
        Else
          Dim vType As String = mvEnv.GetConfig("fp_cc_authorisation_type")
          Select Case vType
            Case "CSXL210FE"
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.CommsXL
            Case "SCXLVPCSCP"
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.SecureCXL
            Case "PROTX"
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.ProtX
            Case "TNSHOSTED"
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.TnsHosted
            Case "SAGEPAYHOSTED"
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.SagePayHosted
            Case Else
              mvOnlineAuthorisationType = OnlineAuthorisationTypes.None
          End Select
        End If
      End If
      Return mvOnlineAuthorisationType
    End Function

    Public Sub CheckOnlineAuthorisation()
      If OnlineAuthorisationType() = OnlineAuthorisationTypes.CommsXL Then
        Dim vPath As String = mvEnv.GetConfig("fp_cc_authorisation_directory")
        If vPath.Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "fp_cc_authorisation_directory")
        Dim vFileName As String = System.IO.Path.Combine(vPath, "ccard.lck")
        If Not My.Computer.FileSystem.FileExists(vFileName) Then RaiseError(DataAccessErrors.daeCCAuthorisationServerNotRunning)
      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.SecureCXL Then
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSCPURL).Length = 0 OrElse _
          mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSCPAPIVersion).Length = 0 OrElse _
          mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlVPCURL).Length = 0 OrElse _
          mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlVPCAPIVersion).Length = 0 Then
          RaiseError(DataAccessErrors.daeSecureCXLNotSetup)
        End If
      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.ProtX Then
        If mvEnv.GetConfig("protX_web_url").Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "protX_web_url")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAuthorisedTextID) = False Then RaiseError(DataAccessErrors.daeProtXNotSetup)
      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.TnsHosted Then
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTnsHostedPayment) = False Then RaiseError(DataAccessErrors.daeTNSHostedPaymentNotSetUp)
      End If
    End Sub

    Public Function AuthoriseTransaction(ByRef pCS As CardSale, ByRef pType As CreditCardAuthorisationTypes, ByRef pAmount As Double, ByVal pAddressNumber As Integer) As Boolean
      Return AuthoriseTransaction(pCS, pType, pAmount, pAddressNumber, "", "", "", 0)
    End Function

    Public Function AuthoriseTransaction(ByRef pCS As CardSale, ByRef pType As CreditCardAuthorisationTypes, ByRef pAmount As Double, ByVal pAddressNumber As Integer, ByVal pOriginalTransNumber As String,
                                         ByVal pAuthorisationTextId As String, ByVal pSecurityKey As String, ByVal pAuthorisationNumber As Integer) As Boolean
      Return AuthoriseTransaction(pCS, pType, pAmount, pAddressNumber, pOriginalTransNumber, pAuthorisationTextId, pSecurityKey, pAuthorisationNumber, "", "")
    End Function

    Public Function AuthoriseTransaction(ByRef pCS As CardSale, ByRef pType As CreditCardAuthorisationTypes, ByRef pAmount As Double, ByVal pAddressNumber As Integer,
                                         ByVal pOriginalTransNumber As String, ByVal pAuthorisationTextId As String, ByVal pSecurityKey As String, ByVal pAuthorisationNumber As Integer,
                                         ByVal pSession As String, ByVal pBatchCategory As String) As Boolean
      'Class must be initialised first
      Dim vPath As String = ""
      'Setup validation
      Select Case OnlineAuthorisationType(pOriginalTransNumber.Length > 0 OrElse pCS.TemplateNumber.Length > 0)
        Case OnlineAuthorisationTypes.None
          RaiseError(DataAccessErrors.daeInvalidConfig, "fp_cc_authorisation_type")
        Case OnlineAuthorisationTypes.CommsXL
          vPath = mvEnv.GetConfig("fp_cc_authorisation_directory")
          If vPath.Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "fp_cc_authorisation_directory")
        Case OnlineAuthorisationTypes.SecureCXL
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMerchantDetails) = False Then RaiseError(DataAccessErrors.daeSecureCXLNotSetup)
        Case OnlineAuthorisationTypes.ProtX
          If mvEnv.GetConfig("protX_web_url").Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "protX_web_url")
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAuthorisedTextID) = False Then RaiseError(DataAccessErrors.daeProtXNotSetup)
        Case OnlineAuthorisationTypes.TnsHosted
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTnsHostedPayment) = False Then RaiseError(DataAccessErrors.daeTNSHostedPaymentNotSetUp)
      End Select

      'Get Merchant Details
      Dim vMerchantNo As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMultipleMerchantRetailNos) Then
        vMerchantNo = mvEnv.Connection.GetValue("SELECT merchant_retail_number FROM batches b, batch_categories bc WHERE b.batch_number = " & pCS.BatchNumber & " AND bc.batch_category = b.batch_category")
      End If
      If vMerchantNo.Length = 0 Then vMerchantNo = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMerchantRetailNumber)
      If OnlineAuthorisationType() = OnlineAuthorisationTypes.SecureCXL Then
        SecureCXLRequest.MerchantDetails.Init(New SQLStatement(mvEnv.Connection, "merchant_id,access_code,user_name,user_password", "merchant_details", New CDBFields(New CDBField("merchant_retail_number", vMerchantNo))).GetRecordSet)
        If SecureCXLRequest.MerchantDetails.MerchantID.Length = 0 Then RaiseError(DataAccessErrors.daeSecureCXLNotSetup)
      End If

      If OnlineAuthorisationType() <> OnlineAuthorisationTypes.SagePayHosted AndAlso pType <> CreditCardAuthorisationTypes.ccatTemplateOnly Then  'Don't set the class fields for Template only
        'Set the class fields
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationNumber).IntegerValue = mvEnv.GetControlNumber("AN")
        mvClassFields(CreditCardAuthorisationFields.ccafBatchNumber).IntegerValue = pCS.BatchNumber
        mvClassFields(CreditCardAuthorisationFields.ccafTransactionNumber).IntegerValue = pCS.TransactionNumber
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedAmount).DoubleValue = pAmount
        If pCS.AuthorisationCode.Length > 0 Then mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationCode).Value = pCS.AuthorisationCode
      End If

      Dim vDesc As String = ""
      Dim vTT As String = GetTransactionTypeInfo(pType, vDesc)

      Dim vAddress As New Address(mvEnv)
      vAddress.Init(pAddressNumber)


      Dim vAmount As Integer = CInt(pAmount * 100)  'Amount in pence

      If OnlineAuthorisationType() = OnlineAuthorisationTypes.CommsXL Then
        Dim vLine As String = "-mc" & vMerchantNo 'Merchant retail number
        vLine = vLine & " -tr" & vTT 'Transaction Type
        vLine = vLine & " -cd" & pCS.CardNumber 'Card Number
        If pCS.IssueNumber.Length > 0 Then
          vLine = vLine & " -is" & pCS.IssueNumber 'Issue Number
        End If
        If pCS.ValidDate.Length > 0 Then
          vLine = vLine & " -sd" & Right(pCS.ValidDate, 2) & Left(pCS.ValidDate, 2) 'Start Date  YYMM
        End If
        vLine = vLine & " -ed" & Right(pCS.ExpiryDate, 2) & Left(pCS.ExpiryDate, 2) 'Expiry Date YYMM
        vLine = vLine & " -am" & vAmount 'Amount in pence
        vLine = vLine & " -rf" & AuthorisationNumber 'Unique Reference
        If pType = (pType Or CreditCardAuthorisationTypes.ccatWeb) Then
          vLine = vLine & " -et32" 'EMV Terminal Type
        End If
        vLine = vLine & " -cc826" 'Country Code
        vLine = vLine & " -cu826" 'Currency Code
        vLine = vLine & " -ds1" & vDesc 'Description
        If pCS.SecurityCode.Length > 0 Then vLine = vLine & " -sc" & pCS.SecurityCode 'Security Code
        If AuthorisationCode.Length > 0 Then vLine = vLine & " -ac" & AuthorisationCode 'Authorisation Code
        If vAddress.Postcode.Length > 0 And vAddress.UK = True Then
          vLine = vLine & " -hn" & vAddress.AddressNumbersOrName 'House Numbers / Name
          vLine = vLine & " -zp" & vAddress.CCAuthorisationPostcode 'Postcode
        End If
        vLine = vLine & " -x"

        'Now create the temporary file in the EFT directory
        Dim vOutFile As String = System.IO.Path.Combine(vPath, "CDB" & AuthorisationNumber & ".TMP")
        Dim vWriter As New IO.StreamWriter(vOutFile, False)
        vWriter.WriteLine(vLine)
        vWriter.Close()
        'My.Computer.FileSystem.RenameFile(vOutFile, Replace(vOutFile, ".TMP", ".REQ"))
        FileSystem.Rename(vOutFile, Replace(vOutFile, ".TMP", ".REQ"))

        Dim vInFile As String = Replace(vOutFile, ".TMP", ".RSP")
        Dim vStartTime As Date = Now
        Dim vTimeout As Integer = IntegerValue(mvEnv.GetConfig("fp_cc_authorisation_timeout", "195"))
        Dim vEndTime As Date = DateAdd(Microsoft.VisualBasic.DateInterval.Second, vTimeout, vStartTime)
        'Now we have created the file we have to wait for the response file to be created
        Dim vSeconds As Integer
        Do
          System.Threading.Thread.Sleep(200)

          vSeconds = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Second, vStartTime, Now))
          If vSeconds > 1 Then RaiseEvent AuthorisingCreditCard(vTimeout, vSeconds)
        Loop While Not My.Computer.FileSystem.FileExists(vInFile) And vEndTime > Now

        If Now >= vEndTime Then
          'On Error Resume Next
          My.Computer.FileSystem.DeleteFile(Replace(vOutFile, ".TMP", ".REQ")) 'If we timed out the try to delete the file so it doesn't get submitted again
          'On Error GoTo 0
          RaiseError(DataAccessErrors.daeCCAuthorisationTimeout)
        End If
        vLine = My.Computer.FileSystem.ReadAllText(vInFile)

        Dim vItems() As String = Split(vLine, "-")
        Dim vValue As String
        For vIndex As Integer = 0 To UBound(vItems)
          vValue = Trim(Mid(vItems(vIndex), 3))
          Select Case Left(vItems(vIndex), 2)
            Case "rf"
              If CDbl(vValue) <> AuthorisationNumber Then RaiseError(DataAccessErrors.daeInvalidCCAuthResponse)
            Case "rc"
              mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = vValue
            Case "ac"
              If AuthorisationCode.Length > 0 Then
                If vValue <> AuthorisationCode Then RaiseError(DataAccessErrors.daeInvalidCCAuthResponse)
              Else
                mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationCode).Value = vValue
              End If
            Case "ms"
              mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = Replace(vValue, Chr(34), "")
            Case "cd"
              If vValue <> pCS.CardNumber Then RaiseError(DataAccessErrors.daeInvalidCCAuthResponse)
            Case "am"
              If CDbl(vValue) <> vAmount Then RaiseError(DataAccessErrors.daeInvalidCCAuthResponse)
            Case "av"
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAddressVerificationResult).Value = vValue
          End Select
        Next
        My.Computer.FileSystem.DeleteFile(vInFile)
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
        Dim vAuthorised As Boolean = False
        If AuthorisationResponseCode = "00" Or AuthorisationResponseCode = "08" Or AuthorisationResponseCode = "11" Or AuthorisationResponseCode = "Y1" Or AuthorisationResponseCode = "Y3" Then
          'Valid authorisation response code
          vAuthorised = True
        ElseIf AuthorisationResponseCode = "01" Or AuthorisationResponseCode = "02" Or AuthorisationResponseCode = "60" Then
          RaiseError(DataAccessErrors.daeCardRejectedAsOverCeilingLimit)
        ElseIf AuthorisationResponseCode = "05" Then
          RaiseError(DataAccessErrors.daeAuthorisationHasBeenRefused)
        ElseIf AuthorisationResponseCode = "30" Then
          RaiseError(DataAccessErrors.daeMerchantNumberNotSetUp)
        ElseIf AuthorisationResponseCode = "54" Then
          RaiseError(DataAccessErrors.daeCardHasExpired)
        End If
        Save()
        pCS.NoClaimRequired = True 'Did some kind of authorisation so claim not required
        pCS.Save()
        Return vAuthorised

      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.SecureCXL Then

        Dim vVPCURL As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlVPCURL)
        Dim vVPCAPIVersion As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlVPCAPIVersion)

        Dim vVPCRequest As SecureCXLRequest = Nothing
        Dim vContinue As Boolean = True
        Dim vAddCardDetails As Boolean = True
        SecureCXLRequest.MerchantDetails.Timeout = IntegerValue(mvEnv.GetConfig("fp_cc_authorisation_timeout", "195"))

        If vTT = "doRequest" Then
          Dim vSCPURL As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSCPURL)
          Dim vSCPAPIVersion As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSCPAPIVersion)

          If pType = CreditCardAuthorisationTypes.ccatTemplateOnly Then
            'Just save the card details and get a Template Number
            vContinue = False
            'Save the card details
            vVPCRequest = New SecureCXLRequest(vSCPURL, vSCPAPIVersion)
            vVPCRequest.AddDigitalOrderField("vpc_Command", vTT)
            vVPCRequest.AddDigitalOrderField("vpc_RequestType", "payTemplate")
            vVPCRequest.AddDigitalOrderField("vpc_RequestCommand", "doCreateTemplate")
            AddCardDetails(vVPCRequest, pCS)  'AVS and CSC does not work in this call
            SendRequest(vVPCRequest, vVPCURL, vVPCAPIVersion, pCS, True)

          ElseIf pCS.TemplateNumber.Length > 0 Then
            'Normal Transaction: Use SCP Subsequent Trans
            vVPCRequest = New SecureCXLRequest(vSCPURL, vSCPAPIVersion)
            vVPCRequest.AddDigitalOrderField("vpc_RequestType", "payTemplate")
            vVPCRequest.AddDigitalOrderField("vpc_RequestCommand", "doSubTxn")
            vVPCRequest.AddDigitalOrderField("vpc_TemplateNo", pCS.TemplateNumber)
            vAddCardDetails = False

          Else 'If AuthorisationCode.Length = 0 Then
            'Normal Transaction: Use SCP Initial Trans
            vVPCRequest = New SecureCXLRequest(vSCPURL, vSCPAPIVersion)
            vVPCRequest.AddDigitalOrderField("vpc_RequestType", "payTemplate")
            vVPCRequest.AddDigitalOrderField("vpc_RequestCommand", "doInitTxn")
          End If
        Else
          'Refund: Use VCP for Refund (AMA Refund Trans)
          vVPCRequest = New SecureCXLRequest(vVPCURL, vVPCAPIVersion, True)
          vVPCRequest.AddDigitalOrderField("vpc_TransNo", pOriginalTransNumber.ToString)
          vAddCardDetails = False
        End If

        If vContinue Then
          vVPCRequest.AddDigitalOrderField("vpc_Command", vTT)
          vVPCRequest.AddDigitalOrderField("vpc_MerchTxnRef", AuthorisationNumber.ToString)
          vVPCRequest.AddDigitalOrderField("vpc_Amount", vAmount.ToString)

          If vAddCardDetails Then AddCardDetails(vVPCRequest, pCS, vAddress)
          'Perform the transaction
          SendRequest(vVPCRequest, vVPCURL, vVPCAPIVersion, pCS)
        End If
        'If Not AuthorisationResponseCode = "0" Then 
        pCS.ClearCardNumber()
        If pType <> CreditCardAuthorisationTypes.ccatTemplateOnly Then
          Save()
          pCS.NoClaimRequired = True 'Did some kind of authorisation so claim not required
          pCS.Save()
        End If
        If AuthorisationResponseCode = "0" Then Return True
      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.ProtX Then
        Dim vUrl As String = String.Empty
        Dim vProtXRequest As ProtXRequest = Nothing
        Dim vTimeOut As Integer = IntegerValue(mvEnv.GetConfig("fp_cc_authorisation_timeout", "195"))
        Dim vMode As String = String.Empty
        Dim vVendorName As String = mvEnv.GetConfig("protX_vendor_name")
        If vVendorName.Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "protX_vendor_name")

        'For protx use protx card type
        If pCS.CreditCardType.Length <> 0 And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cbdProtxCardType) Then
          Dim vDataTable As CDBDataTable = GetProtxCardType(pCS.CreditCardType)
          If vDataTable IsNot Nothing Then
            pCS.ProtXCardType = vDataTable.Rows(0).Item("protx_card_type")
            pCS.CreditCardType = ""
          End If
        End If

        If vTT = "doRequest" Then
          vMode = "PAYMENT"
          vUrl = mvEnv.GetConfig("protX_web_url")
          vProtXRequest = New ProtXRequest(vUrl, vTimeOut)
          vProtXRequest.AddDigitalOrderField("VPSProtocol", "2.22")
          vProtXRequest.AddDigitalOrderField("TxType", vMode)
          vProtXRequest.AddDigitalOrderField("Vendor", vVendorName)
          vProtXRequest.AddDigitalOrderField("VendorTxCode", AuthorisationNumber.ToString)
          vProtXRequest.AddDigitalOrderField("Amount", pAmount.ToString)
          vProtXRequest.AddDigitalOrderField("Currency", "GBP")
          vProtXRequest.AddDigitalOrderField("Description", AuthorisationNumber.ToString)
          AddAccountType(pType, vProtXRequest)
          AddCardAndAddressDetails(vProtXRequest, pCS, vAddress)
        Else
          vMode = "REFUND"
          If mvEnv.GetConfig("protX_web_refund_url").Length = 0 Then RaiseError(DataAccessErrors.daeInvalidConfig, "protX_web_refund_url")
          vUrl = mvEnv.GetConfig("protX_web_refund_url")
          vProtXRequest = New ProtXRequest(vUrl, vTimeOut)
          vProtXRequest.AddDigitalOrderField("VPSProtocol", "2.22")
          vProtXRequest.AddDigitalOrderField("TxType", vMode)
          vProtXRequest.AddDigitalOrderField("Vendor", vVendorName)
          vProtXRequest.AddDigitalOrderField("VendorTxCode", AuthorisationNumber.ToString)
          vProtXRequest.AddDigitalOrderField("Amount", pAmount.ToString)
          vProtXRequest.AddDigitalOrderField("Currency", "GBP")
          vProtXRequest.AddDigitalOrderField("Description", AuthorisationNumber.ToString)
          AddAccountType(pType, vProtXRequest)
          AddRefundDetails(vProtXRequest, pAuthorisationNumber.ToString, pOriginalTransNumber, pAuthorisationTextId, pSecurityKey)
        End If
        SendValidationRequest(vProtXRequest, pCS)
        If mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value.Length = 0 Then mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
        Save()
        If Not AuthorisationResponseCode = "0" Then pCS.ClearCardNumber() 'the authorisation was not successful so make sure the card number wont be saved.
        pCS.NoClaimRequired = True 'Did some kind of authorisation so claim not required
        pCS.Save()
        If AuthorisationResponseCode = "0" Then Return True
      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.TnsHosted Then
        Dim vCardAuthService As ICardAuthorisationService = Nothing
        If pBatchCategory.Length > 0 Then
          vCardAuthService = CardAuthorisationServiceFactory.GetAuthorisationServiceProvider(mvEnv, pBatchCategory)
        Else
          Dim vBatchCategory As String = mvEnv.Connection.GetValue("SELECT b.batch_category FROM batches b, batch_categories bc WHERE b.batch_number = " & pCS.BatchNumber & " AND bc.batch_category = b.batch_category")
          vCardAuthService = CardAuthorisationServiceFactory.GetAuthorisationServiceProvider(mvEnv, vBatchCategory)
        End If

        Dim vRequest As String = vCardAuthService.GetRequestData(pSession, AuthorisationNumber, pAmount.ToString)
        SendTransactionRequest(vCardAuthService, vRequest, pCS)
        Save()
        If Not AuthorisationResponseCode = "0" Then pCS.ClearCardNumber() 'the authorisation was not successful so make sure the card number wont be saved.
        pCS.NoClaimRequired = True 'Did some kind of authorisation so claim not required
        pCS.Save()
        If AuthorisationResponseCode = "0" Then Return True

      ElseIf OnlineAuthorisationType() = OnlineAuthorisationTypes.SagePayHosted Then
        Return AuthoriseTransactionUsingSagePay(pCS, pAuthorisationNumber)
      End If
      Return False
    End Function

    Private Function AuthoriseTransactionUsingSagePay(pCS As CardSale, pAuthorisationNumber As Integer) As Boolean
      Try
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("authorisation_number", pAuthorisationNumber)
        Dim vUpdateFields As New CDBFields
        vUpdateFields.Add("batch_number", pCS.BatchNumber)
        vUpdateFields.Add("transaction_number", pCS.TransactionNumber)
        mvEnv.Connection.UpdateRecords("credit_card_authorisations", vUpdateFields, vWhereFields, False)
        pCS.NoClaimRequired = True 'Did some kind of authorisation so claim not required
        pCS.ClearCardNumber()
        pCS.Save()
        Return True
      Catch vEx As Exception
        Return False
      End Try
    End Function

    Public Function StoreCardToken(vParams As ParameterList) As Boolean
      Try
        If vParams IsNot Nothing AndAlso vParams.ContainsKey("TokenId") AndAlso Not String.IsNullOrWhiteSpace(vParams("TokenId")) Then
          Dim vInsertFields As New CDBFields
          With vInsertFields
            .Add("credit_card_details_number", mvEnv.GetControlNumber("CC"))
            .Add("contact_number", CInt(vParams.Item("ContactNumber")))
            .Add("credit_card_number", vParams.Item("CardDigits"))
            .Add("token_desc", If(vParams.Contains("TokenDesc") AndAlso vParams("TokenDesc").Length > 0,
                                  String.Format("{0}{1}", vParams("TokenDesc") + " **** ", vParams.Item("CardDigits")),
                                  String.Format("{0}{1}", ProjectText.CardEndingIn + " **** ", vParams.Item("CardDigits"))))
            .Add("token_id", System.Web.HttpUtility.UrlDecode(vParams.Item("TokenId")))
            .Add("expiry_date", System.Web.HttpUtility.UrlDecode(vParams.Item("CardExpiryDate")))
            .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
            .Add("amended_by", mvEnv.User.Logname)

          End With
          mvEnv.Connection.InsertRecord("contact_credit_cards", vInsertFields, True) '`, vWhereFields, False)
          Return True
        Else
          Return True 'Tokens are not set 
        End If
      Catch vEx As Exception
        Return False
      End Try
    End Function

    Private Sub AddAccountType(ByVal pType As CreditCardAuthorisationTypes, ByVal pProtXRequest As ProtXRequest)
      If pType = (pType Or CreditCardAuthorisationTypes.ccatWeb) Then
        pProtXRequest.AddDigitalOrderField("AccountType", "E")
      Else
        pProtXRequest.AddDigitalOrderField("AccountType", "M")
      End If
    End Sub

    Private Function GetTransactionTypeInfo(ByVal pAuthType As CreditCardAuthorisationTypes, ByRef pDesc As String) As String
      Dim vTT As String = ""
      Dim vType As String = ""
      Select Case pAuthType
        Case CreditCardAuthorisationTypes.ccatNormal, CType((pAuthType Or (CreditCardAuthorisationTypes.ccatNormal + CreditCardAuthorisationTypes.ccatWeb)), CreditCardAuthorisationTypes)
          If pAuthType = CreditCardAuthorisationTypes.ccatNormal Then
            vTT = "09" 'Normal Transaction  Type
          Else
            vTT = "B2" 'e-commerce
          End If
          pDesc = "Transaction"
          vType = "A" 'Normal
        Case CreditCardAuthorisationTypes.ccatBackOrder
          vTT = "09" 'Normal Transaction  Type
          pDesc = "Back Order"
          vType = "B" 'Back Order
        Case CreditCardAuthorisationTypes.ccatRefund, CType((pAuthType Or (CreditCardAuthorisationTypes.ccatRefund + CreditCardAuthorisationTypes.ccatWeb)), CreditCardAuthorisationTypes)
          If pAuthType = CreditCardAuthorisationTypes.ccatRefund Then
            vTT = "47" 'Refund Transaction  Type
          Else
            vTT = "B4"  'e-commerce
          End If
          pDesc = "Refund"
          vType = "R" 'Refund
        Case CreditCardAuthorisationTypes.ccatNotional
          vTT = "09" 'Normal Transaction  Type
          pDesc = "Notional Claim"
          vType = "N" 'Notional
      End Select

      If OnlineAuthorisationType() = OnlineAuthorisationTypes.SecureCXL Or OnlineAuthorisationType() = OnlineAuthorisationTypes.ProtX Then
        If vTT = "09" OrElse vTT = "B2" OrElse pAuthType = CreditCardAuthorisationTypes.ccatTemplateOnly Then
          vTT = "doRequest"
        ElseIf vTT = "47" Then
          vTT = "refund"  'Refund
        End If
      End If

      mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationType).Value = vType
      Return vTT
    End Function

#Region "SecureCXL"

    Private Sub AddCardDetails(ByVal pRequest As SecureCXLRequest, ByVal pCS As CardSale)
      AddCardDetails(pRequest, pCS, Nothing)
    End Sub
    Private Sub AddCardDetails(ByVal pRequest As SecureCXLRequest, ByVal pCS As CardSale, ByVal pAddress As Address)
      Throw New NotSupportedException("SecureCXL no longer supported due to PCI compliance requirements")
    End Sub

    Private Sub SendRequest(ByVal pRequest As SecureCXLRequest, ByVal pVPCURL As String, ByVal pVPCAPIVersion As String, ByVal pCS As CardSale)
      SendRequest(pRequest, pVPCURL, pVPCAPIVersion, pCS, False)
    End Sub
    Private Sub SendRequest(ByVal pRequest As SecureCXLRequest, ByVal pVPCURL As String, ByVal pVPCAPIVersion As String, ByVal pCS As CardSale, ByVal pTemplateNumberOnly As Boolean)
      Dim vError As String = pRequest.SendRequest
      If vError.Length > 0 Then
        'Timeout- Check the transaction
        Dim vAttemptsRemaining As Integer = IntegerValue(mvEnv.GetConfig("fp_cc_authorisation_retries", "1"))
        While vAttemptsRemaining > 0
          pRequest = New SecureCXLRequest(pVPCURL, pVPCAPIVersion, True)
          If QueryDR(pRequest, pVPCURL, vError) Then
            'Transaction has been found
            If pTemplateNumberOnly Then
              If pRequest.GetResultField("PCT_ErrMsg", "").Length > 0 Then
                mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "07"
                mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = AuthorisationResponseMessage & ". An error occurred when saving Card Details: " & pRequest.GetResultField("PCT_ErrMsg")
              ElseIf pRequest.GetResultField("vpc_TxnResponseCode") = "0" Then
                pCS.TemplateNumber = pRequest.GetResultField("vpc_TemplateNo")
              End If
            Else
              ProcessTransactionResponse(pRequest, pCS)
            End If
            vAttemptsRemaining = 0
          End If
          vAttemptsRemaining -= 1
        End While
      Else
        'Transaction successful
        If pTemplateNumberOnly Then
          If pRequest.GetResultField("PCT_ErrMsg", "").Length > 0 Then
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "07"
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = AuthorisationResponseMessage & ". An error occurred when saving Card Details: " & pRequest.GetResultField("PCT_ErrMsg")
          ElseIf pRequest.GetResultField("vpc_TxnResponseCode") = "0" Then
            pCS.TemplateNumber = pRequest.GetResultField("vpc_TemplateNo")
          End If
        Else
          ProcessTransactionResponse(pRequest, pCS)
        End If
      End If
    End Sub

    Private Function QueryDR(ByRef pRequest As SecureCXLRequest, ByVal pVPCURL As String, ByVal pFirstError As String) As Boolean
      'Used to check the status of a transaction
      pRequest.AddDigitalOrderField("vpc_Command", "queryDR")
      pRequest.AddDigitalOrderField("vpc_MerchTxnRef", AuthorisationNumber.ToString)
      Dim vError As String = pRequest.SendRequest
      If vError.Length = 0 Then
        If pRequest.GetResultField("vpc_DRExists") = "Y" Then
          'Transaction found
          Return True
        Else
          'Transaction not found - Payment Server did not responded the first request
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "7"
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = "Payment Server error: " & pFirstError
          mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
        End If
      Else
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "7"
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = "Payment Server error: " & vError
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
      End If
      Return False
    End Function

    Private Sub ProcessTransactionResponse(ByVal pRequest As SecureCXLRequest, ByVal pCS As CardSale)
      'Check if the transaction was successful or if there was an error
      Dim vResponseCode As String = pRequest.GetResultField("vpc_TxnResponseCode")

      'Set the fields for the receipt with the result fields
      'Core fields
      mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = vResponseCode
      If vResponseCode = "7" Then
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = pRequest.GetResultField("vpc_Message")
      ElseIf vResponseCode = "" AndAlso pRequest.GetResultField("PCT_ErrMsg", "").Length > 0 Then
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = pRequest.GetResultField("PCT_ErrMsg")
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "07"
      Else
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = SecureCXLCodes.GetResponseCodeDescription(vResponseCode)
      End If

      'pRequest.GetResultField("vpc_AcqResponseCode", "Unknown") 'Generated by the financial institution to indicate the status of the transaction. The results can vary between institutions so it is advisable to use the vpc_TxnResponseCode as it is consistent across all acquirers.
      If pRequest.GetResultField("vpc_AuthorizeId").Length > 0 Then
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationCode).Value = pRequest.GetResultField("vpc_AuthorizeId")
      End If
      mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).Value = pRequest.GetResultField("vpc_TransactionNo")

      ' Card Security Code Fields
      Dim vCSCResultCode As String = pRequest.GetResultField("vpc_CSCResultCode")
      mvClassFields(CreditCardAuthorisationFields.ccafCSCResultCode).Value = vCSCResultCode
      mvClassFields(CreditCardAuthorisationFields.ccafCSCResultMessage).Value = SecureCXLCodes.GetCSCDescription(vCSCResultCode)

      ' Address Verification / Advanced Address Verfication Fields
      Dim vAddressVerificationCode As String = pRequest.GetResultField("vpc_AVSResultCode")
      mvClassFields(CreditCardAuthorisationFields.ccafAddressVerificationResult).Value = vAddressVerificationCode
      mvClassFields(CreditCardAuthorisationFields.ccafAddressVerificationMessage).Value = SecureCXLCodes.GetAVSDescription(vAddressVerificationCode)
      mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()

      If pCS.TemplateNumber.Length = 0 Then pCS.TemplateNumber = pRequest.GetResultField("vpc_TemplateNo")
    End Sub

#End Region

#Region "ProtX"
    Private Sub AddCardAndAddressDetails(ByVal pProtX As ProtXRequest, ByVal pCS As CardSale, ByVal pAddress As Address)
      Throw New NotSupportedException("ProtX no longer supported due to PCI compliance requirements")
    End Sub

    Private Sub AddRefundDetails(ByVal pProtX As ProtXRequest, ByVal pAuthorisationNumber As String, ByVal pAuthorisationTransactionNumber As String, ByVal pAuthorisationTextId As String, ByVal pSecurityKey As String)
      If pProtX IsNot Nothing Then
        pProtX.AddDigitalOrderField("RelatedVPSTxId", pAuthorisationTextId)
        pProtX.AddDigitalOrderField("RelatedVendorTxCode", pAuthorisationNumber)
        pProtX.AddDigitalOrderField("RelatedSecurityKey", pSecurityKey)
        pProtX.AddDigitalOrderField("RelatedTxAuthNo", pAuthorisationTransactionNumber.ToString)
      End If
    End Sub

    Private Sub SendValidationRequest(ByVal pProtX As ProtXRequest, ByVal pCS As CardSale)
      If pProtX IsNot Nothing Then
        Dim vError As String = pProtX.SendRequest
        If vError.Length > 0 Then
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "07"
          mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = AuthorisationResponseMessage & ". An error occurred when saving Card Details: " & vError
        Else
          Dim vResponseString As String = pProtX.GetResultField("Status")
          Select Case vResponseString
            Case "OK", "REGISTERED", "AUTHENTICATED"
              'Transaction successful
              ProcessProtXTransactionResponse(pProtX, pCS)
            Case "INVALID", "MALFORMED"
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = pProtX.GetResultField("StatusDetail").Substring(0, 2)
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = AuthorisationResponseMessage & pProtX.GetResultField("StatusDetail")
            Case Else
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = pProtX.GetResultField("StatusDetail").Substring(0, 2)
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = AuthorisationResponseMessage & pProtX.GetResultField("StatusDetail")
          End Select
        End If
      End If
    End Sub


    Private Sub ProcessProtXTransactionResponse(ByVal pRequest As ProtXRequest, ByVal pCS As CardSale)
      'Check if the transaction was successful or if there was an error
      Dim vResponseCode As String = pRequest.GetResultField("StatusDetail")

      If vResponseCode.Length > 0 Then
        'Set the fields for the receipt with the result fields
        'Core fields
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "0"
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = vResponseCode.Substring(5, vResponseCode.Length - 5)

        'store secure key in the database for authorisation code
        If pRequest.GetResultField("SecurityKey").Length > 0 Then
          mvClassFields(CreditCardAuthorisationFields.ccafAuthorisationCode).Value = pRequest.GetResultField("SecurityKey")
        End If
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).Value = pRequest.GetResultField("TxAuthNo")

        ' Address Verification / Advanced Address Verfication Fields
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = pRequest.GetResultField("VPSTxId")
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
      End If
    End Sub

#End Region


#Region "TNS"
    ''' <summary>
    ''' Sends the transaction request to TNS and set the CCA fields with response data
    ''' </summary>
    ''' <param name="pTns">Authorisation Service interface</param>
    ''' <param name="pRequestData">Request string with all the transactiondetails</param>
    ''' <param name="pCS">Card Sales</param>
    ''' <remarks>Connection error will be stored as ''</remarks>
    Private Sub SendTransactionRequest(ByVal pTns As ICardAuthorisationService, ByVal pRequestData As String, ByVal pCS As CardSale)
      Dim vResponse As Boolean = pTns.SendRequest(pRequestData)
      If pTns.GetResponseData.Count > 0 Then
        Select Case pTns.GetResponseData("result").ToUpper
          Case "SUCCESS"
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "0"
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = pTns.GetResponseData("result").ToUpper
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).Value = pTns.GetResponseData("transaction.authorizationCode").ToUpper
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = pTns.GetResponseData("transaction.reference").ToUpper
          Case "FAILURE"
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = pTns.GetErrorCode(pTns.GetResponseData("result").ToUpper).ToUpper
            If pTns.GetResponseData().ContainsKey("response.acquirerCode") Then
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = pTns.GetErrorCode(pTns.GetResponseData("response.acquirerCode").ToUpper).ToUpper
            ElseIf pTns.GetResponseData().ContainsKey("response.gatewayCode") Then
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = pTns.GetResponseData("response.gatewayCode").ToUpper
            Else
              mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = "Unknown Response"
            End If
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = If(pTns.GetResponseData.ContainsKey("transaction.reference"), pTns.GetResponseData("transaction.reference").ToUpper, "")
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
          Case Else
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = pTns.GetErrorCode(pTns.GetResponseData("result").ToUpper).ToUpper
            mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = If(pTns.GetResponseData.ContainsKey("error.explanation"), pTns.GetResponseData("error.explanation").ToUpper, "UnKnown Error")
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = If(pTns.GetResponseData.ContainsKey("transaction.reference"), pTns.GetResponseData("transaction.reference").ToUpper, "Unknown Reference")
            mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
        End Select
      ElseIf Not vResponse And pTns.GetRawResponseData().Length > 0 Then
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = "99" ' Unknown error
        mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = If(pTns.GetRawResponseData().Length < 255, pTns.GetRawResponseData(), pTns.GetRawResponseData().Substring(0, 250))
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = "" ' Transaction ID will not be available as there were issues while connecting to TNS Server
        mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
      End If
    End Sub
#End Region


    Public Sub UpdateToConfirmedTransaction(ByRef pBT As BatchTransaction)
      mvClassFields.Item(CreditCardAuthorisationFields.ccafBatchNumber).IntegerValue = pBT.BatchNumber
      mvClassFields.Item(CreditCardAuthorisationFields.ccafTransactionNumber).IntegerValue = pBT.TransactionNumber
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AuthorisationCode() As String
      Get
        AuthorisationCode = mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationCode).Value
      End Get
    End Property

    Public ReadOnly Property AuthorisationNumber() As Integer
      Get
        AuthorisationNumber = CInt(mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationNumber).Value)
      End Get
    End Property

    Public ReadOnly Property AuthorisationResponseCode() As String
      Get
        AuthorisationResponseCode = mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value
      End Get
    End Property

    Public ReadOnly Property AuthorisationResponseMessage() As String
      Get
        AuthorisationResponseMessage = mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value
      End Get
    End Property

    Public ReadOnly Property AuthorisationType() As CreditCardAuthorisationTypes
      Get
        Select Case mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisationType).Value
          Case "A"
            AuthorisationType = CreditCardAuthorisationTypes.ccatNormal
          Case "B"
            AuthorisationType = CreditCardAuthorisationTypes.ccatBackOrder
          Case "N"
            AuthorisationType = CreditCardAuthorisationTypes.ccatNotional
          Case "R"
            AuthorisationType = CreditCardAuthorisationTypes.ccatRefund
        End Select
      End Get
    End Property

    Public ReadOnly Property AuthorisedAmount() As Double
      Get
        AuthorisedAmount = CDbl(mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisedAmount).Value)
      End Get
    End Property

    Public ReadOnly Property AuthorisedOn() As String
      Get
        AuthorisedOn = mvClassFields.Item(CreditCardAuthorisationFields.ccafAuthorisedOn).Value
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = CInt(mvClassFields.Item(CreditCardAuthorisationFields.ccafBatchNumber).Value)
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = CInt(mvClassFields.Item(CreditCardAuthorisationFields.ccafTransactionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property AddressVerificationResult() As String
      Get
        AddressVerificationResult = mvClassFields.Item(CreditCardAuthorisationFields.ccafAddressVerificationResult).Value
      End Get
    End Property

    Public ReadOnly Property AuthorisedTransactionNo As String
      Get
        Return mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).Value '.IntegerValue
      End Get
    End Property

    Public ReadOnly Property AuthorisedTextId As String
      Get
        Return mvClassFields(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value
      End Get
    End Property

    Public Sub GetTemplateNumber(ByVal pCS As CardSale)
      If OnlineAuthorisationType() = OnlineAuthorisationTypes.SecureCXL Then AuthoriseTransaction(pCS, CreditCardAuthorisationTypes.ccatTemplateOnly, 0, 0)
    End Sub

    Public Sub UpdateBatchTransForBandR(ByVal pNewBatchNo As Integer, ByVal pNewTransNo As Integer)
      'Called from ConfirmStockAllocation only when adding a new card sale record
      'This is for Authorisation Type B = Back Orders
      If mvExisting Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("authorisation_number", AuthorisationNumber)
        vWhereFields.Add("authorisation_number#2", AuthorisationNumber, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("batch_number", BatchNumber)
        vWhereFields.Add("transaction_number", TransactionNumber)
        vWhereFields.Add("authorisation_type", "R", CDBField.FieldWhereOperators.fwoCloseBracket)
        Dim vUpdateFields As New CDBFields
        vUpdateFields.Add("batch_number", pNewBatchNo)
        vUpdateFields.Add("transaction_number", pNewTransNo)
        mvEnv.Connection.UpdateRecords("credit_card_authorisations", vUpdateFields, vWhereFields, False)
      End If
    End Sub

    Private Function GetProtxCardType(ByVal pCardType As String) As CDBDataTable
      Dim vTable As CDBDataTable = Nothing
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cbdProtxCardType) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("credit_card_type", pCardType)
        vTable = New CDBDataTable
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "protx_card_type", "credit_card_types", vWhereFields)
        vTable.FillFromSQL(mvEnv, vSQLStatement, False)
      Else
        RaiseError(DataAccessErrors.daeProtXNotSetup)
      End If
      Return vTable
    End Function

#Region "Public Property"
    Public WriteOnly Property ContactNumber As Integer
      Set(ByVal value As Integer)
        mvContactNumber = value
      End Set
    End Property
#End Region

    Sub InitFromParameters(pEnv As CDBEnvironment, vParameterList As ParameterList)
      mvEnv = pEnv
      InitClassFields()
      With mvClassFields
        'Always include the primary key attributes
        .Item(CreditCardAuthorisationFields.ccafAuthorisationNumber).IntegerValue = CInt(vParameterList("TransactionNumber").ToString)
        'Modify below to handle each recordset type as required
        If vParameterList IsNot Nothing Then
          .Item(CreditCardAuthorisationFields.ccafAuthorisationType).Value = "A"
          .Item(CreditCardAuthorisationFields.ccafAuthorisedOn).Value = TodaysDateAndTime()
          .Item(CreditCardAuthorisationFields.ccafAuthorisedTextID).Value = vParameterList("VPSTxId")
          .Item(CreditCardAuthorisationFields.ccafAuthorisationResponseCode).Value = vParameterList("Response").ToString
          .Item(CreditCardAuthorisationFields.ccafAuthorisedAmount).Value = vParameterList("Amount").ToString
          .Item(CreditCardAuthorisationFields.ccafAuthorisationCode).Value = vParameterList("SecurityKey").ToString
          .Item(CreditCardAuthorisationFields.ccafAuthorisedTransactionNumber).Value = vParameterList("TransactionNumber").ToString
          .Item(CreditCardAuthorisationFields.ccafAuthorisationResponseMessage).Value = vParameterList("Status").ToString
        End If
      End With
    End Sub
  End Class
End Namespace
