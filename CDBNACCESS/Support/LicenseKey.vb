Namespace Access

  Public Class LicenseKey

    Private Const NUMBER_OF_MODULES As Integer = 13

    Public Enum LicenceKeyModules As Integer
      Tabs
      Actions
      EventsManager
      QuestionairreProcessing
      IntranetCreation
      SelectionManager
      Meetings
      DocumentDistributor
      PostCodeValidator
      WebPageDesigner
      WPDCareModules
      WPDSponsorMe
      Dashboard
      'lkmTransactions         ' First module of Key #2 (=9)
      'lkmEventsEnquiry
      'lkmFinancialManager
      'lkmMemberManager
      'lkmCampaignManager
      'lkmGiftAidManager
      'lkmLegacyManager
      'lkmGAYEManager          ' Last module of Key #2  (=16)
      OutlookAddIn
      Unknown
    End Enum

    Public Enum LicenceKeyErrors As Integer
      NoError
      ErrNoKey
      ErrClientCode
      ErrExpired
      ErrCheckDigit
      ErrInvalidKeyNum

      ErrTooManyUsers
      ErrLicenceInvalid
      ErrInstanceInUse
    End Enum

    Private mvLicensedUsers(NUMBER_OF_MODULES) As Integer
    Private mvLicenseUsage(NUMBER_OF_MODULES) As Boolean
    Private mvExpiryDate As String
    Private mvEnv As CDBEnvironment
    Private mvStartTime As Date
    Private mvBuildNumber As Integer
    Private mvKeyError As LicenceKeyErrors

    Public Sub New(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      mvKeyError = DecodeKey(pEnv.GetConfig("option_system_key"), pEnv.ClientCode)
    End Sub

    Public Function ClearModule(ByVal pModule As LicenceKeyModules) As LicenceKeyErrors
      If mvKeyError <> LicenceKeyErrors.NoError Then Return mvKeyError
      If pModule = LicenceKeyModules.Unknown Then
        Return LicenceKeyErrors.ErrInvalidKeyNum
      End If
      Dim vModuleCode As String = GetLicenseModuleCode(pModule)
      Dim vUpdateFields As New CDBFields
      vUpdateFields.Add("active", "N")
      Dim vWhereFields As New CDBFields
      Select Case pModule
        Case LicenceKeyModules.Tabs
          vWhereFields.Add("module", "WPD%", CDBField.FieldWhereOperators.fwoNotLike).SpecialColumn = True
        Case LicenceKeyModules.WebPageDesigner
          vWhereFields.Add("module", "WPD%", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
        Case Else
          vWhereFields.Add("module", vModuleCode).SpecialColumn = True
      End Select
      vWhereFields.Add("logname", mvEnv.User.Logname)
      '.Add "start_time", cftTime, mvStartTime
      vWhereFields.TableAlias = "sys_module_users"
      mvEnv.Connection.UpdateRecords("sys_module_users", vUpdateFields, vWhereFields, False)
      'mvLicenceUsage(pLicenceKeyModule) = False
    End Function

    Public Function ClearModule(ByVal pModuleCode As String) As LicenceKeyErrors
      Dim vLicenseKeyModule As LicenceKeyModules = GetLicenseKeyModule(pModuleCode)
      Return ClearModule(vLicenseKeyModule)
    End Function

    Public Function StartModule(ByVal pModule As LicenceKeyModules, ByVal pBuildNumber As Integer, ByVal pReportActiveUsers As Boolean, ByVal pRemoveActiveUsers As Boolean) As LicenceKeyErrors
      Dim vModuleCode As String = GetLicenseModuleCode(pModule)
      Return StartModule(vModuleCode, pBuildNumber, pRemoveActiveUsers, pRemoveActiveUsers)
    End Function

    Public Function StartModule(ByVal pModuleCode As String, ByVal pBuildNumber As Integer, ByVal pReportActiveUsers As Boolean, ByVal pRemoveActiveUsers As Boolean) As LicenceKeyErrors
      If mvKeyError <> LicenceKeyErrors.NoError Then Return mvKeyError
      'First figure out which module we are interested in
      Dim vLicenseKeyModule As LicenceKeyModules = GetLicenseKeyModule(pModuleCode)
      If vLicenseKeyModule = LicenceKeyModules.Unknown Then
        Return LicenceKeyErrors.ErrInvalidKeyNum
      End If
      'Check if we think we are already using a license for this module
      If mvLicenseUsage(vLicenseKeyModule) Then
        'Wont ever happen for the smart client
        Debug.Assert(False, "Cannot happen in the smart client usage")
      Else
        'If not we need to add as a user
        mvStartTime = Now
        If mvBuildNumber = 0 Then mvBuildNumber = pBuildNumber

        'First check to see if the system thinks the user is already using the module
        Dim vModuleUser As New ModuleUser(mvEnv)
        vModuleUser.Init()
        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("module", pModuleCode).SpecialColumn = True
        vWhereFields.Add("logname", mvEnv.User.Logname)
        vWhereFields.TableAlias = "smu"
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vModuleUser.GetRecordSetFields, "sys_module_users smu", vWhereFields, "active DESC")
        Dim vRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet
        Dim vExistingRecord As Boolean = False
        Dim vNamedUser As Boolean = False
        Dim vExistingModuleUser As ModuleUser = Nothing
        While vRecordSet.Fetch()
          vModuleUser.InitFromRecordSet(vRecordSet)
          If vModuleUser.Active = True And pReportActiveUsers = True Then
            vRecordSet.CloseRecordSet()
            Return LicenceKeyErrors.ErrInstanceInUse
          End If
          If (vModuleUser.Active = False Or pRemoveActiveUsers = True) And (vExistingRecord = False Or vModuleUser.NamedUser = True) Then
            vExistingRecord = True
            If vModuleUser.NamedUser Then vNamedUser = True
            vExistingModuleUser = vModuleUser       'Re-use this record
            vModuleUser = New ModuleUser(mvEnv)
          End If
        End While
        vRecordSet.CloseRecordSet()

        'Remove active user records for this logname if required
        vWhereFields.TableAlias = "sys_module_users"
        Dim vUpdateFields As New CDBFields
        If pRemoveActiveUsers Then
          vUpdateFields.Add("active", "N")
          vWhereFields.Clear()
          Select Case vLicenseKeyModule
            Case LicenceKeyModules.Tabs
              vWhereFields.Add("module", "WPD%", CDBField.FieldWhereOperators.fwoNotLike).SpecialColumn = True
            Case LicenceKeyModules.WebPageDesigner
              vWhereFields.Add("module", "WPD%", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
            Case Else
              vWhereFields.Add("module", pModuleCode).SpecialColumn = True
          End Select
          vWhereFields.Add("logname", mvEnv.User.Logname)
          vWhereFields.Add("active", "Y")
          mvEnv.Connection.UpdateRecords("sys_module_users", vUpdateFields, vWhereFields)
        End If
        Dim vRefusedAccess As Boolean = False
        Dim vRetries As Integer = 0
        If vNamedUser = False Then
          Do
            vWhereFields.Clear()
            vWhereFields.Add("module", pModuleCode).SpecialColumn = True  'This module only
            vWhereFields.Add("active", "Y", CDBField.FieldWhereOperators.fwoOpenBracket)
            vWhereFields.Add("named_user", "Y", CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoOR)

            If mvEnv.Connection.GetCount("sys_module_users", vWhereFields) >= GetNumberOfUsers(vLicenseKeyModule) Then
              'Remove users who have not updated their record for more than 2 hours
              vUpdateFields.Clear()
              vUpdateFields.Add("active", "N")
              vWhereFields.Clear()
              vWhereFields.Add("module", pModuleCode).SpecialColumn = True  'This module only
              vWhereFields.Add("active", "Y")
              vWhereFields.Add("last_updated_on", CDBField.FieldTypes.cftTime, DateAdd("h", -2, Now).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThan)
              mvEnv.Connection.UpdateRecords("sys_module_users", vUpdateFields, vWhereFields, False)
              vRefusedAccess = True
              vRetries = vRetries + 1
            Else
              vRefusedAccess = False
            End If
          Loop While vRefusedAccess And (vRetries < 2)
        End If

        If Not vExistingRecord Then
          vModuleUser.Init()
          vModuleUser.Create(pModuleCode, mvStartTime, mvBuildNumber, vRefusedAccess)
          vModuleUser.Save()
        Else
          With vExistingModuleUser
            If vRefusedAccess Then
              .SetRefusedAccess(mvBuildNumber)
            Else
              .SetActive(mvStartTime, mvBuildNumber)
            End If
            .Save()
          End With
        End If
        If vRefusedAccess Then
          Return LicenceKeyErrors.ErrTooManyUsers
        Else
          mvLicenseUsage(vLicenseKeyModule) = True
        End If
      End If
    End Function

    Public Function GetNumberOfUsers(ByVal pModule As LicenceKeyModules) As Integer
      Return mvLicensedUsers(pModule)
    End Function

    Public Function GetNumberOfUsers(ByVal pModuleCode As String) As Integer
      Return mvLicensedUsers(GetLicenseKeyModule(pModuleCode))
    End Function

    Public Function GetLicenseKeyModule(ByVal pCode As String) As LicenceKeyModules
      Select Case pCode
        Case "IN"
          Return LicenceKeyModules.IntranetCreation         'Option Intranet
        Case "PV"
          Return LicenceKeyModules.PostCodeValidator        'Postcode Validation
        Case "DD"
          Return LicenceKeyModules.DocumentDistributor      'Document Distribution
        Case "ME"
          Return LicenceKeyModules.Meetings                 'Meetings
        Case "GM"
          Return LicenceKeyModules.SelectionManager         'General Mailing
        Case "QP"
          Return LicenceKeyModules.QuestionairreProcessing  'Questionnaire Processing
        Case "EV"
          Return LicenceKeyModules.EventsManager            'Events
        Case "AM"
          Return LicenceKeyModules.Actions                  'Action Manager
        Case "CD"
          Return LicenceKeyModules.Tabs                     'CARE
        Case "WPD"
          Return LicenceKeyModules.WebPageDesigner          'Web Page Designer
        Case "WPDC"
          Return LicenceKeyModules.WPDCareModules           'WPD Care User Modules
        Case "WPDS"
          Return LicenceKeyModules.WPDSponsorMe             'WPD Sponsor Me Modules
        Case "DASH"
          Return LicenceKeyModules.Dashboard                'Dashboard
        Case "OLA"
          Return LicenceKeyModules.OutlookAddIn
        Case Else
          Return LicenceKeyModules.Unknown
      End Select
    End Function

    Public Function GetLicenseModuleCode(ByVal pModule As LicenceKeyModules) As String
      Select Case pModule
        Case LicenceKeyModules.IntranetCreation         'Option Intranet
          Return "IN"
        Case LicenceKeyModules.PostCodeValidator        'Postcode Validation
          Return "PV"
        Case LicenceKeyModules.DocumentDistributor      'Document Distribution
          Return "DD"
        Case LicenceKeyModules.Meetings                 'Meetings
          Return "ME"
        Case LicenceKeyModules.SelectionManager         'General Mailing
          Return "GM"
        Case LicenceKeyModules.QuestionairreProcessing  'Questionnaire Processing
          Return "QP"
        Case LicenceKeyModules.EventsManager            'Events
          Return "EV"
        Case LicenceKeyModules.Actions                  'Action Manager
          Return "AM"
        Case LicenceKeyModules.Tabs                     'CARE
          Return "CD"
        Case LicenceKeyModules.WebPageDesigner          'Web Page Designer
          Return "WPD"
        Case LicenceKeyModules.WPDCareModules           'WPD Care User Modules
          Return "WPDC"
        Case LicenceKeyModules.WPDSponsorMe             'WPD Sponsor Me Modules
          Return "WPDS"
        Case LicenceKeyModules.Dashboard                'Dashboard
          Return "DASH"
        Case LicenceKeyModules.OutlookAddIn
          Return "OLA"
        Case Else                                       'Change to fix compiler warning
          Return ""
      End Select
    End Function

    Private Function DecodeKey(ByVal pKeyString As String, ByVal pClientCode As String) As LicenceKeyErrors
      Dim vSourceString As String
      Dim vCheckDigit As Integer
      Dim vIndex As Integer
      Dim vCode As Integer
      Dim vCount As Integer
      Dim vDecodeDigit As Integer
      Dim vNumber As Integer
      Dim vClient As Integer

      If pKeyString.Length = 0 Then
        Return LicenceKeyErrors.ErrNoKey
        Exit Function
      End If

      Debug.Print("Decoding Key:")
      vCheckDigit = Asc(pKeyString.Chars(0)) - Asc("A")
      Debug.Print("Check Digit: " & vCheckDigit)
      'Now use the check digit to shift all the characters apart from the check digit
      If vCheckDigit = 0 Then
        vDecodeDigit = 4
      Else
        vDecodeDigit = vCheckDigit
      End If
      For vIndex = 2 To pKeyString.Length
        vCode = Asc(pKeyString.Substring(vIndex - 1, 1))
        If vCode <> Asc("-") Then
          If vCode < 65 Then 'Numerics
            vCode = vCode - vDecodeDigit
            Do While vCode < 48
              vCode = vCode + 10
            Loop
          Else 'Alpha
            vCode = vCode - vDecodeDigit
            Do While vCode < 65
              vCode = vCode + 26
            Loop
          End If
          Mid(pKeyString, vIndex, 1) = Chr(vCode)
        Else
          vDecodeDigit = vDecodeDigit + 1
        End If
      Next
      Debug.Print("Decoded:      " & pKeyString)

      'Validate the check digit
      vCount = 0
      For vIndex = 2 To pKeyString.Length
        vCount = vCount + Asc(pKeyString.Substring(vIndex - 1, 1))
      Next
      If vCount Mod 26 = vCheckDigit Then
        Debug.Print("Check Digit Matches")
      Else
        Return LicenceKeyErrors.ErrCheckDigit
      End If

      'Get the expiry date
      If pKeyString.Substring(1, 1) <> "0" Then
        vNumber = CodeToNumber(pKeyString.Substring(1, 3))

        Dim vYear As Integer = vNumber \ 12
        Dim vMonth As Integer = vNumber Mod 12
        If vMonth = 0 Then
          vMonth = 12
          vYear = vYear - 1
        End If
        mvExpiryDate = vMonth & "/" & vYear
        Debug.Print("Expiry Date: " & mvExpiryDate)
        If (vYear < Year(Now)) Or ((vYear = Year(Now)) And vMonth < Month(Now)) Then
          Return LicenceKeyErrors.ErrExpired
        End If
      End If

      'Check the client code matches
      vSourceString = pKeyString.Substring(5)
      For vIndex = 1 To pClientCode.Length
        vClient = vClient + (Asc(pClientCode.Substring(vIndex - 1, 1)) - Asc("A")) * (((vIndex - 1) * 26) + 1)
      Next
      If Not ("0000" & Hex(vClient)).EndsWith(vSourceString.Substring(0, 4)) Then
        Return LicenceKeyErrors.ErrClientCode
      End If
      vSourceString = pKeyString.Substring(10)

      'Now get the counts for the modules
      Do While vSourceString.Length > 0
        If Char.IsDigit(vSourceString.Chars(0)) OrElse (vSourceString.Chars(0) >= "A"c AndAlso vSourceString.Chars(0) <= "Z"c) Then
          vNumber = CodeToNumber(vSourceString.Substring(1, 3))
          Select Case vSourceString.Chars(0)
            Case "0"c
              mvLicensedUsers(LicenceKeyModules.Tabs) = vNumber
              Debug.Print("Tabs:                     " & vNumber)
            Case "1"c
              mvLicensedUsers(LicenceKeyModules.Actions) = vNumber
              Debug.Print("Actions:                  " & vNumber)
            Case "2"c
              mvLicensedUsers(LicenceKeyModules.EventsManager) = vNumber
              Debug.Print("Events:                   " & vNumber)
            Case "3"c
              mvLicensedUsers(LicenceKeyModules.QuestionairreProcessing) = vNumber
              Debug.Print("Questionnaire Processing: " & vNumber)
            Case "4"c
              mvLicensedUsers(LicenceKeyModules.IntranetCreation) = vNumber
              Debug.Print("Intranet:                 " & vNumber)
            Case "5"c
              mvLicensedUsers(LicenceKeyModules.SelectionManager) = vNumber
              Debug.Print("Selection Manager:        " & vNumber)
            Case "6"c
              mvLicensedUsers(LicenceKeyModules.Meetings) = vNumber
              Debug.Print("Meetings:                 " & vNumber)
            Case "7"c
              mvLicensedUsers(LicenceKeyModules.DocumentDistributor) = vNumber
              Debug.Print("Document Distribution:    " & vNumber)
            Case "8"c
              mvLicensedUsers(LicenceKeyModules.PostCodeValidator) = vNumber
              Debug.Print("Postcode Validation:      " & vNumber)
            Case "9"c
              mvLicensedUsers(LicenceKeyModules.WebPageDesigner) = vNumber
              Debug.Print("Web Page Designer:        " & vNumber)
            Case "A"c
              mvLicensedUsers(LicenceKeyModules.WPDCareModules) = vNumber
              Debug.Print("WPD CARE Modules:         " & vNumber)
            Case "B"c
              mvLicensedUsers(LicenceKeyModules.WPDSponsorMe) = vNumber
              Debug.Print("WPD Sponsor Me:           " & vNumber)
            Case "C"c
              mvLicensedUsers(LicenceKeyModules.Dashboard) = vNumber
              Debug.Print("Dashboard:                " & vNumber)
          End Select
        End If
        If vSourceString.Length < 5 Then Exit Do
        vSourceString = vSourceString.Substring(5)
      Loop
    End Function

    Private Function CodeToNumber(ByVal pCode As String) As Integer
      Dim pDigit1 As Integer
      Dim pDigit2 As Integer
      Dim pDigit3 As Integer
      Dim pChar As Integer

      pChar = Asc(pCode.Chars(0))
      If pChar > 64 Then
        pDigit1 = pChar - 65
      Else
        pDigit1 = (pChar - 48) + 26
      End If
      pChar = Asc(pCode.Substring(1, 1))
      If pChar > 64 Then
        pDigit2 = pChar - 65
      Else
        pDigit2 = (pChar - 48) + 26
      End If
      pChar = Asc(pCode.Substring(2, 1))
      If pChar > 64 Then
        pDigit3 = pChar - 65
      Else
        pDigit3 = (pChar - 48) + 26
      End If
      Return (pDigit1 * (36 * 36)) + (pDigit2 * 36) + pDigit3
    End Function

    Public Function GetInitialisationKey() As String
      Return GenerateKey("", mvEnv.ClientCode, {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1})
    End Function

    Private Function GenerateKey(pExpiry As String, pClientCode As String, vModuleCounts() As Integer) As String
      Dim vClient As Integer
      Dim vCode As Integer
      Dim vIndex As Integer
      Dim vCount As Integer
      Dim vSourceString As String
      Dim vKeyString As String
      Dim vItems As Integer
      Dim vItem As Integer
      Dim vPos As Integer
      Dim vCheckDigit As Integer
      Dim vEncodeDigit As Integer

      Randomize()
      'Place a space where the check digit will go
      vSourceString = " "
      'If there is an expiry date then store the year * 12 plus the month
      'Else use a zero as the second character and randomize the others
      If IsDate(pExpiry) Then
        vCode = Year(CDate(pExpiry)) * 12 + Month(CDate(pExpiry))
        vSourceString = vSourceString & NumberToCode(vCode)
      Else
        vSourceString = vSourceString & "0"
        vSourceString = vSourceString & Chr(CInt(26 * Rnd() + 1) + Asc("@")) 'Random char
        vSourceString = vSourceString & Chr(CInt(26 * Rnd() + 1) + Asc("@")) 'Random char
      End If

      'Now add the client code block
      For vIndex = 1 To pClientCode.Length
        vClient = vClient + (Asc(Mid(pClientCode, vIndex, 1)) - Asc("A")) * (((vIndex - 1) * 26) + 1)
      Next
      'Having totaled the ascii values of all the characters turn it into hex
      Dim vTemp As String = "0000" & Hex(vClient)
      vSourceString = vSourceString & "-" & vTemp.Substring(vTemp.Length - 4)

      'Now add the numbers for each module
      'The first digit is the indentifier for the module
      'Next come the number of users encoded to base 36
      For vIndex = 0 To vModuleCounts.Length - 1
        vCode = vModuleCounts(vIndex)
        If vCode > 0 Then
          If vIndex < 10 Then
            vSourceString = vSourceString & "-" & vIndex.ToString & NumberToCode(vCode)
          Else
            vSourceString = vSourceString & "-" & Chr(Asc("A") + vIndex - 10) & NumberToCode(vCode)
          End If
        End If
      Next

      'Now pad the string to at least 49 characters
      While vSourceString.Length < 49
        vSourceString = vSourceString & "-" & Chr(CInt(25 * Rnd() + 1) + Asc("@")) & NumberToCode(CInt(1000 * Rnd()))
      End While

      'Generate a check digit and store it as a character at the start of the string
      vCount = 0
      For vIndex = 2 To Len(vSourceString)
        vCount = vCount + Asc(Mid(vSourceString, vIndex, 1))
      Next
      vCheckDigit = vCount Mod 26
      Debug.Print("Check Digit: " & vCheckDigit)
      Mid(vSourceString, 1, 1) = Chr(vCheckDigit + Asc("A"))
      'lblSystemCodes.Text = vSourceString
      Debug.Print("System Codes: " & vSourceString)

      'Now we have the system codes let's generate the system key
      'First take the check digit and expiry date segment and the client code segment
      'Keep them at the start of the system key
      vKeyString = vSourceString.Substring(0, 9)
      vSourceString = vSourceString.Substring(10)
      'Now randomze the order of the other blocks
      While vSourceString.Length > 0
        vItems = (vSourceString.Length + 1) \ 5
        vItem = CInt(Int(vItems * Rnd() + 1))
        vPos = ((vItem - 1) * 5) + 1
        vKeyString = vKeyString & "-" & Mid(vSourceString, ((vItem - 1) * 5) + 1, 4)
        If vItem > 1 Then
          vSourceString = TruncateString(vSourceString, vPos - 1) & Mid(vSourceString, vPos + 5)
        Else
          vSourceString = Mid(vSourceString, 6)
        End If
      End While
      Debug.Print("Re-Ordered:   " & vKeyString)

      'Now use the check digit to shift all the characters apart from the check digit
      If vCheckDigit = 0 Then
        vEncodeDigit = 4
      Else
        vEncodeDigit = vCheckDigit
      End If
      For vIndex = 2 To Len(vKeyString)
        vCode = Asc(Mid(vKeyString, vIndex, 1))
        If vCode <> Asc("-") Then
          If vCode < 65 Then 'Numerics
            vCode = vCode + vEncodeDigit
            While vCode > 57
              vCode = vCode - 10
            End While
          Else 'Alpha
            vCode = vCode + vEncodeDigit
            While vCode > 90
              vCode = vCode - 26
            End While
          End If
          Mid(vKeyString, vIndex, 1) = Chr(vCode)
        Else
          vEncodeDigit = vEncodeDigit + 1
        End If
      Next
      Debug.Print("Encoded:      " & vKeyString)
      DecodeKey(vKeyString, pClientCode)
      Return vKeyString
    End Function

    Private Function NumberToCode(ByVal pNumber As Integer) As String
      'Max value supported is 36 * 36 * 36 = 46656
      'Turn number into a 3 character code in modulo 36
      Dim pDigit1 As Integer
      Dim pDigit2 As Integer
      Dim pDigit3 As Integer
      Dim pString As String = ""

      pDigit1 = pNumber \ (36 * 36)
      pNumber = pNumber - pDigit1 * (36 * 36)
      pDigit2 = pNumber \ 36
      pNumber = pNumber - pDigit2 * 36
      pDigit3 = pNumber

      If pDigit1 < 26 Then
        pString = pString & Chr(pDigit1 + 65)
      Else
        pString = pString & CStr(pDigit1 - 26)
      End If
      If pDigit2 < 26 Then
        pString = pString & Chr(pDigit2 + 65)
      Else
        pString = pString & CStr(pDigit2 - 26)
      End If
      If pDigit3 < 26 Then
        pString = pString & Chr(pDigit3 + 65)
      Else
        pString = pString & CStr(pDigit3 - 26)
      End If
      Return pString
    End Function

  End Class

End Namespace
