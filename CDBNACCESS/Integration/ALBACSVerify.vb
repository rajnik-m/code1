Imports System.Runtime.InteropServices
Imports CARE.Config

Namespace Access


  Public Class AccountNoVerify
    Implements IDisposable
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   Note:    The ValCode.dll will require the following *.idx, *.dbf, *.fdf files:
    '   V.1.0
    '            Enhanced database system:
    '
    '            1.  Sortcddb    (Sortcode database) contains all valid UK sort codes and
    '                            the data required to validate their account numbers.
    '
    '            2.  BkAddrs     (Bank address database) contains details of all main branch
    '                            bank addresses
    '
    '
    '            Standard database system:
    '
    '            1.  Acmpdata    Contains all BACS sortcode ranges and the appropriate formula
    '                            to perform validation checks on their account numbers.
    '                            (Also requires *.hdr file)
    '
    '            2.  Acmpsubs    Contains list of sort codes and their substitutions.
    '
    '
    '    The Valcode.dll will return a VAL error code if any of the above files are missing
    '    for the appropriate database system and a call to ValidateAcc will not be successful.
    '
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   VERSION constants are declared.
    '
    '   Sent to ValStartUp telling the dll what database system is to be used and
    '   what set of internal processes are to be carried out on the sortcode and
    '   account number.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Const STANDARD As Short = 0 ' Use standard database system and processes
    Private Const ENHANCED As Short = 1 ' Use enhanced database system and processes

    Public Enum AccountNoVerifyResult
      avrvalid = 1
      avrSortcodeValidAccountInvalid
      avrInvalid
      avrSortcodeValidAccountWarn
      avrWarning
    End Enum

    Public Enum BankAccountValidatorType
      bavtNone
      bavtAFD
      bavtAllBacs
      bavtVSeries
    End Enum

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   VALIDATION STATUS constants are declared
    '
    '   These fields are returned in the ValidityLevel member of the AccValStruct structure and
    '   contain level of validation achieved by the sortcode/ account number combination.
    '
    '   The ValidityLevel is a bit field and must be interpreted differently depending on what
    '   mode ValCode.dll is running in.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Const INVALID As Short = 0 ' (enh/std) Sortcode and account number invalid
    Private Const SEARCH_VALID As Short = 1 ' bit 1 (enh only) sort code is valid
    Private Const FORMULA_POSITIVE_VALID As Short = 2 ' bit 2 (enh) account number is valid
    '       (std) sortcode and account no. valid
    Private Const FORMULA_NEGATIVE_VALID As Short = 4 ' bit 3 (std only) sortcode and acount no.
    '       valid by default

    '''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   EXPLICIT validation values declared
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''

    Private Const ENH_SORTCODE_ACCOUNT_INVALID As Short = 0
    Private Const ENH_SORTCODE_VALID_ACCOUNT_INVALID As Short = 1 ' bit 1 set
    Private Const ENH_SORTCODE_ACCOUNT_VALID As Short = 3 ' bit 1 and 2 set

    Private Const STD_SORTCODE_ACCOUNT_INVALID As Short = 0
    Private Const STD_SORTCODE_ACCOUNT_POS_VALID As Short = 2 ' bit 2 set
    Private Const STD_SORTCODE_ACCOUNT_NEG_VALID As Short = 4 ' bit 4 set

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   BIT VALUES used for the DataReturnMap parameter sent to ValidateAcc.
    '
    '   The DataReturnMap tells the ValCode.dll what data is to be placed in the AccValStruct.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Const FLD_VALID_NUMBER As Short = 1 ' bit 1
    Private Const FLD_SORT_CODE As Short = 2 ' bit 2
    Private Const FLD_ACCOUNT_NUMBER As Short = 4 ' bit 3
    Private Const FLD_BRANCH As Short = 8 ' bit 4
    Private Const FLD_ADDRESS_ONE As Short = 16 ' bit 5
    Private Const FLD_ADDRESS_TWO As Short = 32 ' bit 6
    Private Const FLD_ADDRESS_THREE As Short = 64 ' bit 7
    Private Const FLD_POSTCODE As Short = 128 ' bit 8
    Private Const FLD_TELEPHONE As Short = 256 ' bit 9
    Private Const FLD_SPECIAL As Short = 512 ' bit 10
    Private Const FLD_BRANCH_TITLE As Short = 1024 ' bit 11
    Private Const FLD_BANK_NAME As Short = 2048 ' bit 12
    Private Const FLD_TOWN As Short = 4096 ' bit 13
    Private Const FLD_LAST_AMEND As Short = 8192 ' bit 14
    Private Const FLD_TRANSACTIONS As Short = 16384 ' bit 15

    Private Const ALL_FIELDS As Short = 32767 ' all 15 bits set (all fields returned)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   ERROR codes returned from ValCode.dll
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '   Error flags     '

    Private Const VAL_NO_ERROR As Integer = 0 ' ValCode.dll is working correctly
    Private Const VAL_FIELD_ERROR As Integer = 1 ' Error retrieving data from a database field
    Private Const VAL_DATABASE_ERROR As Integer = 2 ' Error reading from, or opening a database file
    Private Const VAL_BACSVAL_ERROR As Integer = 3 ' Error occured during account validation process
    Private Const VAL_ERROR_OPENING_BK_ADDR_DB As Integer = 4 ' Error opening Bkaddrs database files
    Private Const VAL_ERROR_OPENING_SORTCODE_DB As Integer = 5 ' Error opening Sortcddb database files
    Private Const VAL_ERROR_OPENING_STANDARD_DB As Integer = 6 ' Error opening Acmpdata database files
    Private Const VAL_DATABASE_SYS_STARTUP_ERROR As Integer = 7 ' Error initialising the database system
    Private Const VAL_ERROR_CLOSING_STANDARD_DB As Integer = 8 ' Error closing Acmpdata database files
    Private Const VAL_ERROR_CLOSING_SORTCODE_DB As Integer = 9 ' Error closing Sortcddb database files
    Private Const VAL_ERROR_CLOSING_BK_ADDR_DB As Integer = 10 ' Error closing Bkaddrs database files
    Private Const VAL_DATABASE_SYS_SHUTDOWN_ERROR As Integer = 11 ' Error shuting down the database system
    Private Const VAL_ERROR_CHANGING_DIRECTORY As Integer = 12 ' Error changing the working directory
    Private Const VAL_PASSED_UPGRADE_DATE As Short = 14 ' Database upgrade due
    Private Const VAL_ENHANCED_MODE_DENIED As Short = 20 ' Enhanced mode denied (automatically opens in STANDARD)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   ACCOUNT VALIDATION STRUCTURE: AccValStruct is declared
    '
    '   ValCode.dll populates this structure with null terminated strings (except ValidityLevel)
    '
    '   It is important to note that the address given is that of the main branch that
    '   handles the transactions of accounts with the given sort code, this is not necessarily the
    '   the address of the actual bank that is referenced by the sort code (sub branch), this is
    '   a limitation of the database system.  The branch title/location of the bank actually referenced
    '   by the sort code is placed within the Branch member of the AccValStruct.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' <summary>
    ''' Always use the custom constructor when declaring New. This will set the correct length of fields that is required in third party DLL
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential, Size:=545), Serializable()> _
    Public Structure AccValStruct
      Public ValidityLevel As Short      ' Represents validation results 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=7)> Private SortCode As String          ' Sort code (set by the user) to validate 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=11)> Private AccountNumber As String    ' Account no. (set by the user) to validate 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=31)> Private Branch As String           ' Branch address details 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=36)> Private AddressOne As String       ' Branch Address Line One details 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=36)> Private AddressTwo As String       ' Branch Address Line Two details 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=36)> Private AddressThree As String     ' Branch Address Line Three details 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=36)> Private AddressFour As String      ' Branch Address Line Four details 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=11)> Private Postcode As String         ' Branch Postcode 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=22)> Private Telephone As String        ' Branch Telephone Number 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=121)> Private BranchTitle As String     ' Branch Title 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=101)> Private BankName As String        ' Bank Name 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=31)> Private Town As String             ' Town where Branch is located 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=31)> Private County As String           ' County where Branch is located 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=13)> Private DateLastAmended As String  ' Date these details were last amended 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=9)> Private Transactions As String      ' Type of transactions prohibited 
      Private AccountType As Short                                                                    ' Account type (0-9) 
      <MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=9)> Private RecordIndex As String       ' Reserved 

      Public Sub New(ByVal pSort As String, ByVal pAccount As String)
        SortCode = pSort
        AccountNumber = pAccount
      End Sub
    End Structure

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    '   ValCode.dll function are declared.
    '
    '   ValStartUp:     initialises database system and sets internal flags that effect the way sort codes
    '                   and account numbers are processed and validated.
    '
    '   ValShutDown:    closes down the database system and frees any allocated system memory.
    '
    '   ValidateAcc:    Validates a give sort code and account number and populates the AccValStruc
    '                   structure according to the bits set in the DataReturnMap parameter.
    '                   Returns a VAL_ error code.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Declare Function ValStartUp Lib "valsrt32.dll" (ByVal VersionFlag As Short, ByVal Path As String, ByVal UpgradeDate As String) As Integer
    Private Declare Function ValShutDown Lib "valsrt32.dll" () As Integer
    'UPGRADE_WARNING: Structure AccValStruct may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ValidateAcc Lib "valsrt32.dll" (ByRef ValStruct As AccValStruct, ByVal DataReturnMap As Short) As Integer

    Private Structure SearchValStruct
      'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
      <VBFixedString(101), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=101)> Public BranchName() As Char
      'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
      <VBFixedString(121), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=121)> Public BranchTitle() As Char
      'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
      <VBFixedString(25), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=25)> Public Town() As Char
      'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
      <VBFixedString(37), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=37)> Public AddressThree() As Char
      'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
      'UPGRADE_NOTE: Postcode was upgraded to Postcode_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
      <VBFixedString(11), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=11)> Public Postcode_Renamed() As Char
    End Structure

    'UPGRADE_WARNING: Structure SearchValStruct may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ValMakeAddrList Lib "valsrt32.dll" (ByRef SearchStruct As SearchValStruct, ByRef Count As Short) As Integer
    'UPGRADE_WARNING: Structure AccValStruct may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ValGetAddrItem Lib "valsrt32.dll" (ByVal Idx As Short, ByRef Item As AccValStruct) As Integer
    'UPGRADE_WARNING: Structure AccValStruct may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function ValMakeSortCodeList Lib "valsrt32.dll" (ByRef Item As AccValStruct, ByRef Count As Short) As Integer
    Private Declare Function ValGetSortCode Lib "valsrt32.dll" (ByVal Idx As Short, ByVal SortCode As String) As Integer

    Friend Enum UseVerifyType
      uvtNone
      uvtWarn
      uvtError
    End Enum

    Private mvEnv As CDBEnvironment
    Private mvUseVerify As UseVerifyType
    Private mvInitialised As Boolean
    Private mvStarted As Boolean
    Private mvVersionFlag As Short
    Private mvBankValidatorType As BankAccountValidatorType

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pAlbacsBankDetails As String = "C")
      mvEnv = pEnv
      mvBankValidatorType = BankAccountValidatorType.bavtNone

      Dim vConfig As String = String.Empty
      Select Case mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAccountValidationType).ToUpper
        Case "UC"
          'Use Configs
          vConfig = pEnv.GetConfig("afd_software")
          If ((vConfig.Equals("BOTH", StringComparison.InvariantCultureIgnoreCase) OrElse vConfig.Equals("BANK", StringComparison.InvariantCultureIgnoreCase)) _
          AndAlso mvEnv.GetConfig("afd_everywhere_server").Length > 0) Then
            mvBankValidatorType = BankAccountValidatorType.bavtAFD
          ElseIf pEnv.GetConfig("albacs_verify").Length > 0 Then
            mvBankValidatorType = BankAccountValidatorType.bavtAllBacs
          End If
        Case "VS"
          'V-Series
          mvBankValidatorType = BankAccountValidatorType.bavtVSeries
      End Select

      If Not mvInitialised Then
        'See if verification is required
        mvUseVerify = UseVerifyType.uvtNone
        vConfig = String.Empty
        Select Case mvBankValidatorType
          Case BankAccountValidatorType.bavtAFD
            vConfig = pEnv.GetConfig("afd_verify").ToLower

          Case BankAccountValidatorType.bavtAllBacs, BankAccountValidatorType.bavtVSeries
            vConfig = pEnv.GetConfig("albacs_verify").ToLower
            If pAlbacsBankDetails.Equals("C", StringComparison.InvariantCultureIgnoreCase) = False Then
              Select Case pAlbacsBankDetails.ToUpper
                Case "W"
                  vConfig = "warn"
                Case Else
                  vConfig = "error"
              End Select
            End If
        End Select
        Select Case vConfig
          Case "warn"
            mvUseVerify = UseVerifyType.uvtWarn
          Case "error"
            mvUseVerify = UseVerifyType.uvtError
          Case Else
            mvUseVerify = UseVerifyType.uvtNone
        End Select

        If mvUseVerify <> UseVerifyType.uvtNone Then
          Select Case mvBankValidatorType
            Case BankAccountValidatorType.bavtAFD
              Dim vAFDServerLocation As String = pEnv.GetConfig("afd_everywhere_server") 'afd_software
              Dim vAFDSoftware As String = pEnv.GetConfig("afd_software")
              If vAFDServerLocation.Length = 0 OrElse vAFDSoftware.Length = 0 Then RaiseError(DataAccessErrors.daeAFDVerify)

            Case BankAccountValidatorType.bavtAllBacs
              Dim vUpgradeDate As String = "            "
              Dim vDataFilePath As String = pEnv.GetConfig("albacs_verify_path")
              If vDataFilePath.Length = 0 Then vDataFilePath = "C:\Verify4w"
              If pEnv.GetConfigOption("albacs_verify_enhanced", True) Then
                mvVersionFlag = ENHANCED
              Else
                mvVersionFlag = STANDARD
              End If
              Dim vErrorCode As Integer = ValStartUp(mvVersionFlag, vDataFilePath, vUpgradeDate)
              If vErrorCode = VAL_NO_ERROR Then
                mvStarted = True
              Else
                RaiseError(DataAccessErrors.daeALBACSVerify, vErrorCode.ToString)
              End If

            Case BankAccountValidatorType.bavtVSeries
              If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlAccountValidationURL).Length = 0 Then
                RaiseError(DataAccessErrors.daeAccountValidationURLNotSet)
              End If
          End Select
        End If
        mvInitialised = True
      End If
    End Sub

    Public Function Validate(ByRef pSortCode As String, ByRef pAccountNo As String, ByRef pMsg As String) As AccountNoVerifyResult
      Select Case mvBankValidatorType
        Case BankAccountValidatorType.bavtNone
          Return AccountNoVerifyResult.avrvalid
        Case BankAccountValidatorType.bavtAFD
          Return ValidateUsingAFD(pSortCode, pAccountNo, pMsg)
        Case BankAccountValidatorType.bavtAllBacs
          Return ValidateUsingAllBacs(pSortCode, pAccountNo, pMsg)
        Case Else
          Return ValidateBankDetails(pSortCode, pAccountNo, pMsg)
      End Select
    End Function
    Private Function ValidateUsingAllBacs(ByRef pSortCode As String, ByRef pAccountNo As String, ByRef pMsg As String) As AccountNoVerifyResult
      Dim vErrorCode As Integer
      Dim vDataReturnMap As Short
      Dim vReturn As AccountNoVerifyResult

      If mvStarted And Len(pSortCode) > 0 Then
        Dim vAccValStruct As New AccValStruct(pSortCode, pAccountNo)
        'Debug.Print "ALBACS Verify SortCode: " & pSortCode & " Account: " & pAccountNo
        vDataReturnMap = FLD_BRANCH Or FLD_ADDRESS_ONE Or FLD_ADDRESS_TWO Or FLD_ADDRESS_THREE Or FLD_TOWN
        If Len(pSortCode) < 6 Then
          vReturn = AccountNoVerifyResult.avrInvalid
        Else
          vErrorCode = ValidateAcc(vAccValStruct, vDataReturnMap)
          If vErrorCode <> VAL_NO_ERROR Then
            RaiseError(DataAccessErrors.daeALBACSVerify, CStr(vErrorCode))
          Else
            'Debug.Print "ALBACS Verify Result: " & .ValidityLevel
            Select Case vAccValStruct.ValidityLevel
              Case ENH_SORTCODE_ACCOUNT_VALID, STD_SORTCODE_ACCOUNT_POS_VALID, STD_SORTCODE_ACCOUNT_NEG_VALID
                vReturn = AccountNoVerifyResult.avrvalid
              Case ENH_SORTCODE_VALID_ACCOUNT_INVALID
                If pAccountNo.Length > 0 Then
                  vReturn = AccountNoVerifyResult.avrSortcodeValidAccountInvalid
                Else
                  vReturn = AccountNoVerifyResult.avrvalid
                End If
              Case ENH_SORTCODE_ACCOUNT_INVALID, STD_SORTCODE_ACCOUNT_INVALID
                If Len(pAccountNo) = 0 And mvVersionFlag = STANDARD Then
                  vReturn = AccountNoVerifyResult.avrvalid
                Else
                  vReturn = AccountNoVerifyResult.avrInvalid
                End If
            End Select
          End If
          If vReturn <> AccountNoVerifyResult.avrvalid Then
            If pAccountNo.Length > 0 Then
              pMsg = "Sort Code and Account Number Combination is Not Valid"
            Else
              pMsg = "Sort Code is Not Valid"
            End If
            If mvUseVerify = UseVerifyType.uvtWarn Then
              Select Case vReturn
                Case AccountNoVerifyResult.avrInvalid
                  vReturn = AccountNoVerifyResult.avrWarning
                Case AccountNoVerifyResult.avrSortcodeValidAccountInvalid
                  vReturn = AccountNoVerifyResult.avrSortcodeValidAccountWarn
              End Select
            End If
          End If
          ValidateUsingAllBacs = vReturn
        End If
      Else
        ValidateUsingAllBacs = AccountNoVerifyResult.avrvalid
      End If
    End Function

    Private Enum AFDInvalidIssueCodes
      InvalidAccountNumber = -14
    End Enum

    Private Function ValidateUsingAFD(ByVal pSortCode As String, ByVal pAccountNo As String, ByRef pMsg As String) As AccountNoVerifyResult
      Dim vXMLDoc As Xml.XmlDocument
      Dim vRoot As Xml.XmlElement
      Dim vDataNode As Xml.XmlNode
      Dim vItemNodes As System.Xml.XmlNodeList
      Dim vXmlParams As String
      Dim vReturn As AccountNoVerifyResult = AccountNoVerifyResult.avrvalid

      Dim vSerialNumber As String = String.Empty
      Dim vPassword As String = String.Empty
      If NfpConfigrationManager.BankFinderAuthenticationValues IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.BankFinderAuthenticationValues.UsernameValue) AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.BankFinderAuthenticationValues.PasswordValue) Then
        vSerialNumber = NfpConfigrationManager.BankFinderAuthenticationValues.UsernameValue
        vPassword = NfpConfigrationManager.BankFinderAuthenticationValues.PasswordValue
      End If

      ' Initialise the Microsoft XML Document Object Model
      vXMLDoc = New System.Xml.XmlDocument()
      ' Build up the XML query string
      vXmlParams = mvEnv.GetConfig("afd_everywhere_server") & "/afddata.pce?"
      vXmlParams = vXmlParams & "Serial=" & vSerialNumber & "&"
      vXmlParams = vXmlParams & "Password=" & vPassword & "&"
      'vXmlParams = vXmlParams & "Serial=" & "" & "&"
      'vXmlParams = vXmlParams & "Password=" & "" & "&"
      vXmlParams = vXmlParams & "UserID=" & "" & "&"
      vXmlParams = vXmlParams + "Data=Bank&Task=Account&Clearing=Both&Fields=Account"

      ' Set the Sort Code and Account Number
      vXmlParams = vXmlParams + "&SortCode=" + pSortCode
      vXmlParams = vXmlParams + "&AccountNumber=" + pAccountNo

      ' Load the XML from the webserver with the query string
      Try
        vXMLDoc.Load(vXmlParams)
      Catch vEx As Xml.XmlException
        RaiseError(DataAccessErrors.daeUniservError, "Error: " & vEx.Message)
      End Try

      ' Check if PCE returned an error and if the document is valid
      vRoot = vXMLDoc.DocumentElement
      vDataNode = vRoot.SelectSingleNode("Result")
      vItemNodes = vRoot.SelectNodes("Item")
      If vDataNode Is Nothing Or vItemNodes Is Nothing Then
        RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
      End If
      If Val(vDataNode.InnerText) < 1 Then
        Dim vErrorResult As Integer = IntegerValue(vDataNode.InnerText)
        vDataNode = vRoot.SelectSingleNode("ErrorText")
        If vDataNode Is Nothing Then
          RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
        Else
          If vErrorResult = AFDInvalidIssueCodes.InvalidAccountNumber AndAlso String.IsNullOrEmpty(pAccountNo) Then
            'We were just validating Sort Code so valid
            vReturn = AccountNoVerifyResult.avrvalid
          Else
            pMsg = vDataNode.InnerText ' Show the user the error
            vReturn = AccountNoVerifyResult.avrInvalid
            If mvUseVerify = UseVerifyType.uvtWarn Then
              Select Case vReturn
                Case AccountNoVerifyResult.avrInvalid
                  vReturn = AccountNoVerifyResult.avrWarning
                Case AccountNoVerifyResult.avrSortcodeValidAccountInvalid
                  vReturn = AccountNoVerifyResult.avrSortcodeValidAccountWarn
              End Select
            End If
          End If
        End If
      End If
      Return vReturn
    End Function

    ''' <summary>Validate Bank Details using classes that inherit the IBankAccountValidation interface.</summary>
    ''' <param name="pSortCode">The Sort Code to be validated.  This must be present.</param>
    ''' <param name="pAccountNumber">The Account Number to be validated.  This can be an empty string if only the Sort Code is to be validated.</param>
    ''' <returns>Returns an <see cref="AccountNoVerifyResult">Account Number Verified</see>enum.</returns>
    Private Function ValidateBankDetails(ByVal pSortCode As String, ByVal pAccountNumber As String, ByRef pMsg As String) As AccountNoVerifyResult
      Dim vVerifyResult As AccountNoVerifyResult = AccountNoVerifyResult.avrvalid
      pMsg = String.Empty

      Dim vBankVerifier As IBankAccountValidation = Nothing
      If mvBankValidatorType = BankAccountValidatorType.bavtVSeries Then
        vBankVerifier = New VSeriesAccountValidation(mvEnv, pSortCode, pAccountNumber, mvUseVerify)
      End If

      If vBankVerifier IsNot Nothing Then
        With vBankVerifier
          .ValidateBankAccount()
          vVerifyResult = .VerifyResult
          If .IsValid = False Then pMsg = .InvalidReasonDesc
        End With
      End If

      Return vVerifyResult

    End Function

#Region " IDisposable Support "
    Private DisposedValue As Boolean = False    ' To detect redundant calls

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) below.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
      If Not Me.DisposedValue Then
        If disposing Then
          ' free managed resources when explicitly called
          If mvStarted Then
            ValShutDown()
            mvInitialised = False
            mvStarted = False
          End If
        End If
        ' free shared unmanaged resources
      End If
      Me.DisposedValue = True
    End Sub
#End Region

  End Class

End Namespace
