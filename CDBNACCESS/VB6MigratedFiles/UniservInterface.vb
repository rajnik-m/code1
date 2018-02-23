Namespace Access
  Public Class UniservInterface

    Public Event SetStatus(ByRef pMsg As String)
    Public Event ShowError(ByRef pMsg As String)
    Public Event ShowWarning(ByRef pMsg As String)
    Public Event SetErrorMessage(ByRef pMsg As String)

    Public Enum UniservPostReturnValues
      uspNone
      ustTowns
      ustStreets
      ustParts
      ustCheckedOK
    End Enum

    Private Structure PostConnection
      Dim Initialised As Boolean
      Dim NeedsPost As Boolean
      Dim Active As Boolean
      Dim Session As Integer
      Dim TestMode As Boolean
    End Structure

    Private Structure MailConnection
      Dim Initialised As Boolean
      Dim Active As Boolean
      Dim Session As Integer
      Dim TestMode As Boolean
    End Structure

    Private Shared mvUniservLock As New Object 'BR21061 Uniserv cannot run concurrently. Use SyncLock on this object to lock Uniserv. 
    Private mvEnv As CDBEnvironment
    Private mvLastError As String
    Private mvLastWarning As String
    Private mvPost As PostConnection
    Private mvMail As MailConnection
    Private mvPhoneBook As MailConnection
    Private mvRaiseErrors As Boolean

    Private Const ERROR_LANGUAGE As String = "D"

    Private Const CHECK_ADDRESS As String = "check_address"
    Private Const CHECK_ADRESS As String = "check_adress"
    Private Const MAIL_INSERT As String = "mail_insert"
    Private Const MAIL_DELETE As String = "mail_delete"
    Private Const MAIL_UPDATE As String = "mail_update"
    Private Const MAIL_SEARCH As String = "mail_search"

    Private Const OUT_SEL_LIST_COUNT As String = "out_sel_list_count"
    Private Const OUT_LIST_COUNT As String = "out_list_count"

    ' *******************************
    '  definition of input arguments
    ' *******************************

    Private Const IN_DBREF As String = "in_dbref"
    Private Const IN_COUNTRY_CODE As String = "in_country_code"
    Private Const IN_STR_LINE As String = "in_str_line"
    Private Const IN_COMPANY_NAME As String = "in_company_name"
    Private Const IN_NAME As String = "in_name"
    Private Const IN_NAME_LINE As String = "in_name_line"
    Private Const IN_FIRST_NAME As String = "in_first_name"
    Private Const IN_HNO As String = "in_hno"
    Private Const IN_STR_HNO As String = "in_str_hno"
    Private Const IN_CITY As String = "in_city"
    Private Const IN_ZIP As String = "in_zip"
    Private Const IN_DATE As String = "in_date"

    ' ********************************
    '  definition of output arguments
    ' ********************************

    Private Const OUT_STR_HNO As String = "out_str_hno"
    Private Const OUT_CITY As String = "out_city"
    Private Const OUT_ZIP As String = "out_zip"
    Private Const OUT_MVAL_CITY As String = "out_mval_city"
    Private Const OUT_MIN_ZIP As String = "out_min_zip"
    Private Const OUT_CITY_DETAIL As String = "out_city_detail"
    Private Const OUT_STR As String = "out_str"
    Private Const OUT_HNO_FROM As String = "out_hno_from"
    Private Const OUT_HNO_FROM_AL As String = "out_hno_from_al"
    Private Const OUT_HNO_TO As String = "out_hno_to"
    Private Const OUT_HNO_TO_AL As String = "out_hno_to_al"
    Private Const OUT_HNO_ZIP As String = "out_hno_zip"
    Private Const OUT_HNO_CITY_DISTRICT As String = "out_hno_city_district"
    Private Const OUT_REGION As String = "out_region"
    Private Const OUT_DBREF As String = "out_dbref"

    Private Const OUT_RES_STR As String = "out_res_str"

    Private Const PAR_DEF_COUNTRY_CODE As String = "par_def_country_code"

    Private Const PAR_STR_LEN As String = "par_str_len"
    Private Const MAX_STR_LEN As Integer = 60

    Private Const PAR_LIST_MAX As String = "par_list_max"
    Private Const DEF_LIST_MAX As Integer = 256
    Private Const PAR_MIN_MVAL As String = "par_min_mval"
    Private Const DEF_MIN_MVAL As Integer = 60
    Private Const PAR_DATE_FORMAT As String = "par_date_format"

    Private Const UNI_SEL_CITY As Integer = 1014
    Private Const UNI_SEL_CITY_TRUNC As Integer = 1015
    Private Const UNI_SEL_STR As Integer = 1016
    Private Const UNI_SEL_STR_TRUNC As Integer = 1017
    Private Const UNI_SEL_BOX As Integer = 1018
    Private Const UNI_SEL_BOX_TRUNC As Integer = 1019

    Private Const UNISTR As Integer = 0
    Private Const UNICSTR As Integer = 1
    Private Const UNILONG As Integer = 2
    Private Const UNISHORT As Integer = 3
    Private Const UNIBYTE As Integer = 4

    Private Const UNI_OK As Integer = 0
    Private Const UNI_ERR As Integer = 1
    Private Const UNI_WARN As Integer = 2
    Private Const UNI_BREAK As Integer = 3
    Private Const UNI_MSG As Integer = 4

    Private Const UNIABS As Integer = 0
    Private Const UNIREL As Integer = 1

    Private Const UNI_NOT_CHECKED As Integer = 7

    Private Const UNI_OVERFLOW As Integer = 2003
    Private Const UNI_NO_RSRC As Integer = 2017 '  No resource

    Private Declare Function uni_get_ret_type Lib "clirts32.dll" (ByVal uniresult As Integer) As Integer
    Private Declare Function uni_get_ret_info Lib "clirts32.dll" (ByVal uniresult As Integer) As Integer
    Private Declare Function uni_get_error_msg Lib "clirts32.dll" (ByVal Session As Integer, ByVal Result As Integer, ByVal lang As String, ByVal buf As String, ByVal leng As Integer) As Integer

    Private Declare Function uni_open_session Lib "clirts32.dll" (ByVal args As String, ByRef pSession As Integer) As Integer
    Private Declare Function uni_close_session Lib "clirts32.dll" (ByVal Session As Integer) As Integer

    Private Declare Function uni_start_request Lib "clirts32.dll" (ByVal Session As Integer, ByVal utype As String, ByRef pRequest As Integer) As Integer
    Private Declare Function uni_exec_request Lib "clirts32.dll" (ByVal request As Integer) As Integer
    Private Declare Function uni_close_request Lib "clirts32.dll" (ByVal request As Integer) As Integer

    Private Declare Function uni_set_param Lib "clirts32.dll" (ByVal Session As Integer, ByVal param As String, ByVal Value As Integer) As Integer
    Private Declare Function uni_get_param Lib "clirts32.dll" (ByVal Session As Integer, ByVal param As String, ByRef Value As Integer) As Integer
    Private Declare Function uni_set_string_param Lib "clirts32.dll" (ByVal Session As Integer, ByVal param As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Integer) As Integer
    Private Declare Function uni_get_string_param Lib "clirts32.dll" (ByVal Session As Integer, ByVal param As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Integer) As Integer

    Private Declare Function uni_set_arg Lib "clirts32.dll" (ByVal request As Integer, ByVal arg As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Integer) As Integer
    Private Declare Function uni_get_arg Lib "clirts32.dll" (ByVal request As Integer, ByVal arg As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Integer) As Integer

    Private Declare Function uni_set_cursor Lib "clirts32.dll" (ByVal request As Integer, ByVal Mode As Integer, ByVal pos As Integer) As Integer
    Private Declare Function uni_rollback Lib "clirts32.dll" (ByVal Session As Integer) As Integer
    Private Declare Function uni_commit Lib "clirts32.dll" (ByVal Session As Integer) As Integer

    'Private Declare Function uni_get_ret_type Lib "clirts32.dll" Alias "_uni_get_ret_type@4" (ByVal uniresult As Long) As Long
    'Private Declare Function uni_get_ret_info Lib "clirts32.dll" Alias "_uni_get_ret_info@4" (ByVal uniresult As Long) As Long
    'Private Declare Function uni_get_error_msg Lib "clirts32.dll" Alias "_uni_get_error_msg@20" (ByVal Session As Long, ByVal Result As Long, ByVal lang As String, ByVal buf As String, ByVal leng As Integer) As Long

    'Private Declare Function uni_open_session Lib "clirts32.dll" Alias "_uni_open_session@8" (ByVal args As String, psession As Long) As Long
    'Private Declare Function uni_close_session Lib "clirts32.dll" Alias "_uni_close_session@4" (ByVal Session As Long) As Long

    'Private Declare Function uni_start_request Lib "clirts32.dll" Alias "_uni_start_request@12" (ByVal Session As Long, ByVal utype As String, prequest As Long) As Long
    'Private Declare Function uni_exec_request Lib "clirts32.dll" Alias "_uni_exec_request@4" (ByVal request As Long) As Long
    'Private Declare Function uni_close_request Lib "clirts32.dll" Alias "_uni_close_request@4" (ByVal request As Long) As Long

    'Private Declare Function uni_set_arg Lib "clirts32.dll" Alias "_uni_set_arg@20" (ByVal request As Long, ByVal arg As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Long) As Long
    'Private Declare Function uni_get_arg Lib "clirts32.dll" Alias "_uni_get_arg@20" (ByVal request As Long, ByVal arg As String, ByVal buf As String, ByVal leng As Integer, ByVal utype As Long) As Long

    'Private Declare Function uni_set_cursor Lib "clirts32.dll" Alias "_uni_set_cursor@12" (ByVal request As Long, ByVal Mode As Long, ByVal pos As Long) As Long
    'Private Declare Function uni_rollback Lib "clirts32.dll" Alias "_uni_rollback@4" (ByVal Session As Long) As Long
    'Private Declare Function uni_commit Lib "clirts32.dll" Alias "_uni_commit@4" (ByVal Session As Long) As Long

    'TBD

    Public Sub Init(ByRef pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    Private Function ExtractArg(ByVal pArg As String) As String
      ExtractArg = Left(pArg, InStr(pArg, Chr(0)) - 1)
    End Function

    Public Function UniAddAddress(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByRef pAddress As String, ByRef pTown As String, ByRef pPostCode As String, ByRef pCountry As String, ByRef pDate As String) As Integer
      Dim vErrorNumber As Integer
      Dim vResult As Integer
      Dim vRequest As Integer
      Dim vDBRef As String
      Dim vOrgParm As String

      vErrorNumber = UniOpenMail(mvMail, False)
      If mvMail.Active Then
        RaiseEvent SetStatus((ProjectText.String15101)) 'Adding Address Data to Mail

        'Start a request
        vResult = uni_start_request(mvMail.Session, MAIL_INSERT, vRequest)
        vErrorNumber = UniMailError(mvMail, vResult)
        If vErrorNumber = UNI_OK And Len(pForename) > 0 Then
          vResult = uni_set_arg(vRequest, IN_FIRST_NAME, pForename, Len(pForename) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pSurname) > 0 Then
          vResult = uni_set_arg(vRequest, IN_NAME, pSurname, Len(pSurname) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pName) > 0 Then
          vOrgParm = LCase(mvEnv.GetConfig("uniserv_organisation_parameter", "IN_COMPANY_NAME"))
          vResult = uni_set_arg(vRequest, vOrgParm, pName, Len(pName) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pAddress) > 0 Then
          vResult = uni_set_arg(vRequest, IN_STR_LINE, pAddress, Len(pAddress) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pTown) > 0 Then
          vResult = uni_set_arg(vRequest, IN_CITY, pTown, Len(pTown) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pPostCode) > 0 Then
          vResult = uni_set_arg(vRequest, IN_ZIP, pPostCode, Len(pPostCode) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pCountry) > 0 Then
          vResult = uni_set_arg(vRequest, IN_COUNTRY_CODE, pCountry, Len(pCountry) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pDate) > 0 Then
          vResult = uni_set_arg(vRequest, IN_DATE, CDate(pDate).ToString("yyyyMMdd"), 9, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          vDBRef = CStr(pContactNumber) & "A" & CStr(pAddressNumber)
          vResult = uni_set_arg(vRequest, IN_DBREF, vDBRef, Len(vDBRef) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Issue the request
          vResult = uni_exec_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Close the request
          vResult = uni_close_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          vResult = uni_commit(mvMail.Session)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        RaiseEvent SetStatus("")
      End If
      UniAddAddress = vErrorNumber
    End Function

    Public Sub UniCloseMail()
      Dim vResult As Integer
      Dim vErrorNumber As Integer

      If mvMail.Active Then
        vResult = uni_close_session(mvMail.Session)
        vErrorNumber = UniMailError(mvMail, vResult)
        With mvMail
          .Active = False
          .Initialised = False
        End With
      End If
      If mvPhoneBook.Active Then
        vResult = uni_close_session(mvPhoneBook.Session)
        vErrorNumber = UniMailError(mvPhoneBook, vResult)
        With mvPhoneBook
          .Active = False
          .Initialised = False
        End With
      End If
    End Sub

    Public Sub UniClosePost()
      Dim vResult As Integer
      Dim vErrorNumber As Integer

      If mvPost.Active Then
        vResult = uni_close_session(mvPost.Session)
        vErrorNumber = UniPostError(vResult)
        With mvPost
          .Active = False
          .Initialised = False
        End With
      End If
    End Sub

    Public Function FindContact(ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByVal pDateOfBirth As String, ByRef pContactNumbers As String) As Integer
      SyncLock mvUniservLock
        If MailActive() Then
          mvRaiseErrors = True
          FindContact = UniFindContact(pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pDateOfBirth, pContactNumbers)
          UniCloseMail()
        End If
      End SyncLock
    End Function

    Public Function FindInPhoneBook(ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByVal pDateOfBirth As String, ByRef pContactNumbers As String) As Integer
      If MailActive() Then
        mvRaiseErrors = True
        FindInPhoneBook = UniFindInPhoneBook(pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pDateOfBirth, pContactNumbers)
        UniCloseMail()
      End If
    End Function

    Public Function AddAddress(ByRef pContact As Contact, ByRef pAddress As Address) As Integer
      Dim vForename As String = ""
      Dim vSurname As String = ""
      Dim vName As String = ""
      Dim vOrg As New Organisation(mvEnv)

      If MailActive() Then
        mvRaiseErrors = True
        If pContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vOrg.Init((pContact.ContactNumber))
          vName = vOrg.Name
        Else
          vForename = pContact.Forenames
          vSurname = pContact.Surname
        End If
        AddAddress = UniAddAddress(pContact.ContactNumber, pAddress.AddressNumber, vForename, vSurname, vName, (pAddress.AddressText), (pAddress.Town), (pAddress.Postcode), (pAddress.Country), (pContact.DateOfBirth))
        UniCloseMail()
      End If
    End Function

    Public Function UpdateAddress(ByRef pContact As Contact, ByRef pAddress As Address) As Integer
      Dim vForename As String = ""
      Dim vSurname As String = ""
      Dim vName As String = ""
      Dim vOrg As New Organisation(mvEnv)

      If MailActive() Then
        mvRaiseErrors = True
        If pContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vOrg.Init((pContact.ContactNumber))
          vName = vOrg.Name
        Else
          vForename = pContact.Forenames
          vSurname = pContact.Surname
        End If
        UpdateAddress = UniModAddress(pContact.ContactNumber, pAddress.AddressNumber, vForename, vSurname, vName, (pAddress.AddressText), (pAddress.Town), (pAddress.Postcode), (pAddress.Country), pContact.DateOfBirth)
        UniCloseMail()
      End If
    End Function

    Public Function DeleteAddress(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer) As Integer
      If MailActive() Then
        mvRaiseErrors = True
        DeleteAddress = UniDelAddress(pContactNumber, pAddressNumber)
        UniCloseMail()
      End If
    End Function

    Public Function MailActive() As Boolean
      If Len(mvEnv.GetConfig("uniserv_host")) > 0 And Len(mvEnv.GetConfig("uniserv_mail")) > 0 Then MailActive = True
    End Function

    Public Function UniDelAddress(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer) As Integer
      Dim vErrorNumber As Integer
      Dim vResult As Integer
      Dim vRequest As Integer
      Dim vDBRef As String

      vErrorNumber = UniOpenMail(mvMail, False)
      If mvMail.Active Then
        RaiseEvent SetStatus((ProjectText.String15102)) 'Deleting Address Data from Mail
        'Start a request
        vResult = uni_start_request(mvMail.Session, MAIL_DELETE, vRequest)
        vErrorNumber = UniMailError(mvMail, vResult)
        If vErrorNumber = UNI_OK Then
          vDBRef = CStr(pContactNumber) & "A" & CStr(pAddressNumber)
          vResult = uni_set_arg(vRequest, IN_DBREF, vDBRef, Len(vDBRef) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Issue the request
          vResult = uni_exec_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Close the request
          vResult = uni_close_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          vResult = uni_commit(mvMail.Session)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        RaiseEvent SetStatus("")
      End If
      UniDelAddress = vErrorNumber
    End Function

    Public Function UniFindContact(ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByVal pDateOfBirth As String, ByRef pContactNumbers As String) As Integer
      Dim vErrorNumber As Integer

      vErrorNumber = UniOpenMail(mvMail, False)
      If mvMail.Active Then
        vErrorNumber = UniFind(mvMail, pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pDateOfBirth, pContactNumbers)
      ElseIf mvMail.TestMode Then
        UniTestMode(pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pContactNumbers)
      End If
      UniFindContact = vErrorNumber
    End Function

    Public Function UniFindInPhoneBook(ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByVal pDateOfBirth As String, ByRef pContactNumbers As String) As Integer
      Dim vErrorNumber As Integer

      vErrorNumber = UniOpenMail(mvPhoneBook, True)
      If mvPhoneBook.Active Then
        vErrorNumber = UniFind(mvPhoneBook, pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pDateOfBirth, pContactNumbers)
      ElseIf mvPhoneBook.TestMode Then
        UniTestMode(pForename, pSurname, pName, pStreetNo, pAddress, pTown, pPostCode, pCountry, pContactNumbers)
      End If
      UniFindInPhoneBook = vErrorNumber
    End Function

    Private Sub UniTestMode(ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByRef pContactNumbers As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      pContactNumbers = ""
      If Len(pForename) > 0 Then vWhereFields.Add("forenames", CDBField.FieldTypes.cftCharacter, pForename, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If Len(pName) > 0 Then
        vWhereFields.Add("surname", CDBField.FieldTypes.cftCharacter, pName, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      ElseIf Len(pSurname) > 0 Then
        vWhereFields.Add("surname", CDBField.FieldTypes.cftCharacter, pSurname, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      End If
      If pTown.Length > 0 Then vWhereFields.Add("town", CDBField.FieldTypes.cftCharacter, pTown, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If pPostCode.Length > 0 Then vWhereFields.Add("postcode", CDBField.FieldTypes.cftCharacter, pPostCode, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT DISTINCT c.contact_number FROM contacts c, contact_addresses ca, addresses a WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND c.contact_number = ca.contact_number AND ca.address_number = a.address_number")
      While vRecordSet.Fetch() = True
        If pContactNumbers.Length > 0 Then pContactNumbers = pContactNumbers & ","
        pContactNumbers = pContactNumbers & vRecordSet.Fields(1).Value
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Function UniFind(ByRef pMailConnection As MailConnection, ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByVal pStreetNo As String, ByVal pAddress As String, ByVal pTown As String, ByVal pPostCode As String, ByVal pCountry As String, ByVal pDate As String, ByRef pContactNumbers As String) As Integer
      Dim vErrorNumber As Integer
      Dim vResult As Integer
      Dim vRequest As Integer
      Dim vCountString As String
      Dim vCount As Integer
      Dim vIndex As Integer
      Dim vArg As String
      Dim vNumber As Integer
      Dim vOrgParm As String

      RaiseEvent SetStatus((ProjectText.String15107)) 'Searching UNISERV Database
      'Start a request
      vResult = uni_start_request(pMailConnection.Session, MAIL_SEARCH, vRequest)
      vErrorNumber = UniMailError(pMailConnection, vResult)
      If vErrorNumber = UNI_OK And Len(pForename) > 0 Then
        vResult = uni_set_arg(vRequest, IN_FIRST_NAME, pForename, Len(pForename) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pSurname) > 0 Then
        vResult = uni_set_arg(vRequest, IN_NAME, pSurname, Len(pSurname) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pName) > 0 Then
        vOrgParm = LCase(mvEnv.GetConfig("uniserv_organisation_parameter", "IN_COMPANY_NAME"))
        vResult = uni_set_arg(vRequest, vOrgParm, pName, Len(pName) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pStreetNo) > 0 Then
        vResult = uni_set_arg(vRequest, IN_HNO, pStreetNo, Len(pStreetNo) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pAddress) > 0 Then
        vResult = uni_set_arg(vRequest, IN_STR_LINE, pAddress, Len(pAddress) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pTown) > 0 Then
        vResult = uni_set_arg(vRequest, IN_CITY, pTown, Len(pTown) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pPostCode) > 0 Then
        vResult = uni_set_arg(vRequest, IN_ZIP, pPostCode, Len(pPostCode) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pCountry) > 0 Then
        vResult = uni_set_arg(vRequest, IN_COUNTRY_CODE, pCountry, Len(pCountry) + 1, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK And Len(pDate) > 0 Then
        vResult = uni_set_arg(vRequest, IN_DATE, CDate(pDate).ToString("yyyyMMdd"), 9, UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK Then
        'Issue the request
        vResult = uni_exec_request(vRequest)
        vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      If vErrorNumber = UNI_OK Then
        vCountString = New String(Chr(0), 16)
        vResult = uni_get_arg(vRequest, OUT_LIST_COUNT, vCountString, Len(vCountString), UNICSTR)
        vErrorNumber = UniMailError(pMailConnection, vResult)
        If vErrorNumber = UNI_OK Then
          vCount = IntegerValue(vCountString)
          If vCount > 200 Then vCount = 201 'Restrict to first 200 contacts
          For vIndex = 0 To vCount - 1
            vResult = uni_set_cursor(vRequest, UNIABS, vIndex)
            vErrorNumber = UniMailError(pMailConnection, vResult)
            If vErrorNumber = UNI_OK Then
              vArg = New String(Chr(0), 30)
              vResult = uni_get_arg(vRequest, OUT_DBREF, vArg, Len(vArg), UNICSTR)
              vErrorNumber = UniPostError(vResult)
              If vErrorNumber = UNI_OK Then
                vNumber = IntegerValue(ExtractNumber(ExtractArg(vArg)))
                If pContactNumbers.Length > 0 Then
                  pContactNumbers = pContactNumbers & "," & vNumber
                Else
                  pContactNumbers = CStr(vNumber)
                End If
              End If
            End If
            If vErrorNumber <> 0 Then Exit For
          Next
        End If
      End If
      If vErrorNumber = UNI_OK Or vErrorNumber = UNI_OVERFLOW Or vErrorNumber = UNI_NO_RSRC Then
        'Close the request
        vResult = uni_close_request(vRequest)
        If vErrorNumber = UNI_OK Then vErrorNumber = UniMailError(pMailConnection, vResult)
      End If
      RaiseEvent SetStatus("")
      UniFind = vErrorNumber
    End Function

    Private Sub UniGetList(ByVal pPostReturnValue As UniservPostReturnValues, ByVal pRequest As Integer, ByRef pItems As Collection)
      Dim vResult As Integer
      Dim vCountString As String
      Dim vCount As Integer
      Dim vIndex As Integer
      Dim vArg As String
      Dim vItem As String = ""
      Dim vErrorNumber As Integer

      vCountString = New String(Chr(0), 16)
      vResult = uni_get_arg(pRequest, OUT_SEL_LIST_COUNT, vCountString, Len(vCountString), UNICSTR)
      vErrorNumber = UniPostError(vResult)
      If vErrorNumber = UNI_OK Then
        vCount = IntegerValue(vCountString)
        For vIndex = 0 To vCount - 1
          vResult = uni_set_cursor(pRequest, UNIABS, vIndex)
          vErrorNumber = UniPostError(vResult)
          If vErrorNumber = UNI_OK Then
            Select Case pPostReturnValue
              Case UniservPostReturnValues.ustTowns
                vArg = New String(Chr(0), 6)
                vResult = uni_get_arg(pRequest, OUT_MVAL_CITY, vArg, Len(vArg), UNICSTR)
                vErrorNumber = UniPostError(vResult)
                If vErrorNumber = UNI_OK Then
                  vItem = ExtractArg(vArg)
                  vArg = New String(Chr(0), 7)
                  vResult = uni_get_arg(pRequest, OUT_MIN_ZIP, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  vItem = vItem & " " & ExtractArg(vArg)
                  vArg = New String(Chr(0), 51)
                  vResult = uni_get_arg(pRequest, OUT_CITY, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  vItem = vItem & " " & ExtractArg(vArg)
                  'vArg = String$(51, 0)
                  'vResult = uni_get_arg(pRequest, OUT_CITY_DETAIL, vArg, Len(vArg), UNICSTR)
                  'vErrorNumber = UniPostError(Nothing, vResult)
                End If
                If vErrorNumber = UNI_OK Then vItem = vItem & " " & ExtractArg(vArg)
              Case UniservPostReturnValues.ustStreets
                vArg = New String(Chr(0), 6)
                vResult = uni_get_arg(pRequest, OUT_MVAL_CITY, vArg, Len(vArg), UNICSTR)
                vErrorNumber = UniPostError(vResult)
                If vErrorNumber = UNI_OK Then
                  vItem = ExtractArg(vArg)
                  vArg = New String(Chr(0), 51)
                  vResult = uni_get_arg(pRequest, OUT_STR, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then vItem = vItem & " " & ExtractArg(vArg)
              Case UniservPostReturnValues.ustParts
                vArg = New String(Chr(0), 7)
                vResult = uni_get_arg(pRequest, OUT_HNO_FROM, vArg, Len(vArg), UNICSTR)
                vErrorNumber = UniPostError(vResult)
                If vErrorNumber = UNI_OK Then
                  vItem = ExtractArg(vArg)
                  vArg = New String(Chr(0), 4)
                  vResult = uni_get_arg(pRequest, OUT_HNO_FROM_AL, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  vItem = vItem & ExtractArg(vArg)
                  vArg = New String(Chr(0), 7)
                  vResult = uni_get_arg(pRequest, OUT_HNO_TO, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  vItem = vItem & "-" & ExtractArg(vArg)
                  vArg = New String(Chr(0), 4)
                  vResult = uni_get_arg(pRequest, OUT_HNO_TO_AL, vArg, Len(vArg), UNICSTR)
                  vErrorNumber = UniPostError(vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  vItem = vItem & ExtractArg(vArg)
                  'vArg = String$(7, 0)
                  'vResult = uni_get_arg(pRequest, OUT_HNO_ZIP, vArg, Len(vArg), UNICSTR)
                  'vErrorNumber = UniPostError(Nothing, vResult)
                End If
                If vErrorNumber = UNI_OK Then
                  'vItem = vItem & " " & ExtractArg(vArg)
                  'vArg = String$(51, 0)
                  'vResult = uni_get_arg(pRequest, OUT_HNO_CITY_DISTRICT, vArg, Len(vArg), UNICSTR)
                  'vErrorNumber = UniPostError(Nothing, vResult)
                End If
                'If vErrorNumber = UNI_OK Then vItem = vItem & " " & ExtractArg(vArg)
            End Select
            If vErrorNumber <> 0 Then Exit For
            If Len(vItem) > 0 Then pItems.Add(vItem)
          End If
        Next
        'frmSelectModal!lstItems.ListIndex = 0
      End If
    End Sub

    Private Function UniMailError(ByRef pMailConnection As MailConnection, ByVal pResult As Integer) As Integer
      Dim vResult As Integer
      Dim vErrorMsg As String
      Dim vType As Integer
      Dim vErrorNumber As Integer

      vType = uni_get_ret_type(pResult)
      If vType = UNI_ERR Then
        vErrorMsg = New String(" "c, 255) ' allocate space for error-message
        vResult = uni_get_error_msg(pMailConnection.Session, pResult, ERROR_LANGUAGE, vErrorMsg, Len(vErrorMsg))
        vErrorNumber = uni_get_ret_info(pResult)
        mvLastError = "MAIL " & ExtractArg(vErrorMsg)
        RaiseEvent ShowError(mvLastError)
        If vErrorNumber < 1000 Then
          vResult = uni_close_session(pMailConnection.Session)
          pMailConnection.Active = False
          pMailConnection.Initialised = False
        End If
        UniMailError = vErrorNumber
        If mvRaiseErrors Then RaiseError(DataAccessErrors.daeUniservError, mvLastError)
      ElseIf vType = UNI_WARN Then
        vErrorMsg = New String(" "c, 255) ' allocate space for error-message
        vResult = uni_get_error_msg(pMailConnection.Session, pResult, ERROR_LANGUAGE, vErrorMsg, Len(vErrorMsg))
        vErrorNumber = uni_get_ret_info(pResult)
        mvLastError = "MAIL " & ExtractArg(vErrorMsg)
        RaiseEvent ShowWarning(mvLastError)
        'Only return errors from the following warnings
        If vErrorNumber = UNI_OVERFLOW Or vErrorNumber = UNI_NO_RSRC Then
          UniMailError = vErrorNumber
          If mvRaiseErrors Then RaiseError(DataAccessErrors.daeUniservError, mvLastError)
        End If
      End If
    End Function

    Public Function UniModAddress(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pForename As String, ByVal pSurname As String, ByVal pName As String, ByRef pAddress As String, ByRef pTown As String, ByRef pPostCode As String, ByRef pCountry As String, ByVal pDate As String) As Integer
      Dim vErrorNumber As Integer
      Dim vResult As Integer
      Dim vRequest As Integer
      Dim vDBRef As String
      Dim vOrgParm As String

      vErrorNumber = UniOpenMail(mvMail, False)
      If mvMail.Active Then
        RaiseEvent SetStatus((ProjectText.String15103)) 'Updating Address Data in Mail

        'Start a request
        vResult = uni_start_request(mvMail.Session, MAIL_UPDATE, vRequest)
        vErrorNumber = UniMailError(mvMail, vResult)
        If vErrorNumber = UNI_OK And Len(pForename) > 0 Then
          vResult = uni_set_arg(vRequest, IN_FIRST_NAME, pForename, Len(pForename) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pSurname) > 0 Then
          vResult = uni_set_arg(vRequest, IN_NAME, pSurname, Len(pSurname) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pName) > 0 Then
          vOrgParm = LCase(mvEnv.GetConfig("uniserv_organisation_parameter", "IN_COMPANY_NAME"))
          vResult = uni_set_arg(vRequest, vOrgParm, pName, Len(pName) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pAddress) > 0 Then
          vResult = uni_set_arg(vRequest, IN_STR_LINE, pAddress, Len(pAddress) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pTown) > 0 Then
          vResult = uni_set_arg(vRequest, IN_CITY, pTown, Len(pTown) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pPostCode) > 0 Then
          vResult = uni_set_arg(vRequest, IN_ZIP, pPostCode, Len(pPostCode) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pCountry) > 0 Then
          vResult = uni_set_arg(vRequest, IN_COUNTRY_CODE, pCountry, Len(pCountry) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK And Len(pDate) > 0 Then
          vResult = uni_set_arg(vRequest, IN_DATE, CDate(pDate).ToString("yyyyMMdd"), 9, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          vDBRef = CStr(pContactNumber) & "A" & CStr(pAddressNumber)
          vResult = uni_set_arg(vRequest, IN_DBREF, vDBRef, Len(vDBRef) + 1, UNICSTR)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Issue the request
          vResult = uni_exec_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          'Close the request
          vResult = uni_close_request(vRequest)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        If vErrorNumber = UNI_OK Then
          vResult = uni_commit(mvMail.Session)
          vErrorNumber = UniMailError(mvMail, vResult)
        End If
        RaiseEvent SetStatus("")
      End If
      UniModAddress = vErrorNumber
    End Function

    Private Function UniOpenMail(ByRef pMailConnection As MailConnection, ByVal pPhoneBook As Boolean) As Integer
      Dim vUniHost As String
      Dim vUniService As String
      Dim vResult As Integer
      Dim vErrorNumber As Integer
      Dim vCountry As String
      Dim vValue As Integer

      mvLastError = ""
      mvLastWarning = ""
      If pMailConnection.Initialised = False Then
        vUniHost = mvEnv.GetConfig("uniserv_host")
        If pPhoneBook Then
          vUniService = mvEnv.GetConfig("uniserv_phone_book")
        Else
          vUniService = mvEnv.GetConfig("uniserv_mail")
        End If
        If vUniService <> "" Then
          'Open session to uniserv server
          RaiseEvent SetStatus((ProjectText.String15104)) 'Connecting to UNISERV Mail Host
          If vUniService = "TEST_MODE" Then
            pMailConnection.Initialised = True
            pMailConnection.TestMode = True
          Else
            pMailConnection.TestMode = False
            vResult = uni_open_session("service='" & vUniService & "', servicehost='" & vUniHost & "'", pMailConnection.Session) 'NoTranslate
            vErrorNumber = UniMailError(pMailConnection, vResult)
            If vErrorNumber = UNI_OK Then
              pMailConnection.Active = True
              vCountry = "CH"
              vResult = uni_set_string_param(pMailConnection.Session, PAR_DEF_COUNTRY_CODE, vCountry, Len(vCountry) + 1, UNICSTR)
              vErrorNumber = UniMailError(pMailConnection, vResult)
              If vErrorNumber = UNI_OK Then
                vValue = IntegerValue(mvEnv.GetConfig("uniserv_list_max"))
                If vValue <= 0 Then vResult = uni_get_param(pMailConnection.Session, PAR_LIST_MAX, vValue)
                vErrorNumber = UniMailError(pMailConnection, vResult)
                If vErrorNumber = UNI_OK Then
                  vResult = uni_set_param(pMailConnection.Session, PAR_LIST_MAX, vValue)
                  vErrorNumber = UniMailError(pMailConnection, vResult)
                End If
              End If
              If vErrorNumber = UNI_OK Then
                vValue = IntegerValue(mvEnv.GetConfig("uniserv_min_mval"))
                If vValue <= 0 Then vResult = uni_get_param(pMailConnection.Session, PAR_MIN_MVAL, vValue)
                vErrorNumber = UniMailError(pMailConnection, vResult)
                If vErrorNumber = UNI_OK Then
                  vResult = uni_set_param(pMailConnection.Session, PAR_MIN_MVAL, vValue)
                  vErrorNumber = UniMailError(pMailConnection, vResult)
                End If
              End If
              If vErrorNumber = UNI_OK Then
                vResult = uni_set_string_param(pMailConnection.Session, PAR_DATE_FORMAT, "yyyymmdd", 9, UNICSTR)
                vErrorNumber = UniMailError(pMailConnection, vResult)
              End If
            End If
            pMailConnection.Initialised = pMailConnection.Active
          End If
          RaiseEvent SetStatus("")
        Else
          pMailConnection.Initialised = True
        End If
      End If
      UniOpenMail = vErrorNumber
    End Function

    Public Function UniNeedPostCheck() As Boolean
      If Not mvPost.Initialised Then
        mvPost.NeedsPost = Len(mvEnv.GetConfig("uniserv_host")) > 0
        mvPost.Initialised = True
      End If
      UniNeedPostCheck = mvPost.NeedsPost
    End Function

    Private Sub UniOpenPost(ByVal pCountry As String)
      Dim vUniHost As String
      Dim vUniService As String
      Dim vResult As Integer
      Dim vErrorNumber As Integer

      vUniHost = mvEnv.GetConfig("uniserv_host")
      vUniService = mvEnv.GetConfig(LCase("uniserv_post_" & pCountry))
      If vUniService <> "" Then
        'Open session to uniserv server
        RaiseEvent SetStatus((ProjectText.String15105)) 'Connecting to UNISERV Post Host
        If vUniService = "TEST_MODE" Then
          mvPost.TestMode = True
        Else
          mvPost.TestMode = False
          vResult = uni_open_session("service='" & vUniService & "', servicehost='" & vUniHost & "'", mvPost.Session) 'NoTranslate
          If UniPostError(vResult) = 0 Then
            mvPost.Active = True
            vResult = uni_set_param(mvPost.Session, PAR_STR_LEN, MAX_STR_LEN)
            vErrorNumber = UniPostError(vResult)
          End If
        End If
        RaiseEvent SetStatus("")
      End If
    End Sub

    Public WriteOnly Property RaiseErrors() As Boolean
      Set(ByVal Value As Boolean)
        mvRaiseErrors = Value
      End Set
    End Property

    Public Function UniPostCheck(ByRef pPostCode As String, ByRef pAddress As String, ByRef pTown As String, ByRef pCounty As String, ByRef pCountry As String, ByRef pReturnValue As UniservPostReturnValues, ByRef pTownSelect As String, ByRef pStreetSelect As String, ByRef pPartSelect As String, ByRef pItems As Collection) As Integer
      Dim vResult As Integer
      Dim vRequest As Integer
      Dim vErrorNumber As Integer
      Dim vArg As String
      Dim vLastError As String = ""
      Dim vLastWarning As String = ""
      Dim vLastLine As String = ""
      Dim vReturnAdd As String
      Dim vSelection As String = ""

      mvLastError = ""
      mvLastWarning = ""
      pReturnValue = UniservPostReturnValues.uspNone
      If Not mvPost.Active Then UniOpenPost(pCountry)
      If mvPost.Active Then
        RaiseEvent SetStatus((ProjectText.String15106)) 'Validating Address Data
        'Start a request
        vResult = uni_start_request(mvPost.Session, CHECK_ADDRESS, vRequest)
        vErrorNumber = UniPostError(vResult)
        'Set the town
        If vErrorNumber = UNI_OK Then
          If pTown.Length > 0 Then
            vResult = uni_set_arg(vRequest, IN_CITY, pTown, Len(pTown) + 1, UNICSTR)
            vErrorNumber = UniPostError(vResult)
          End If
        End If
        'Set the postcode
        If vErrorNumber = UNI_OK Then
          If pPostCode.Length > 0 Then
            vResult = uni_set_arg(vRequest, IN_ZIP, pPostCode, Len(pPostCode) + 1, UNICSTR)
            vErrorNumber = UniPostError(vResult)
          End If
        End If
        'Set the address
        If vErrorNumber = UNI_OK Then
          If pAddress.Length > 0 Then
            vLastLine = LastAddressLine(pAddress)
            If vLastLine.Length > 0 Then
              vResult = uni_set_arg(vRequest, IN_STR_HNO, vLastLine, Len(vLastLine) + 1, UNICSTR)
              vErrorNumber = UniPostError(vResult)
            End If
          End If
        End If
labeltest:
        If vErrorNumber = UNI_OK Then
          'Issue the request
          vResult = uni_exec_request(vRequest)
          vErrorNumber = UniPostError(vResult)
        End If
        If vErrorNumber = UNI_OK Then
          If uni_get_ret_type(vResult) = UNI_BREAK Then
            'We have a selection returned
            If uni_get_ret_info(vResult) = UNI_SEL_CITY Or uni_get_ret_info(vResult) = UNI_SEL_CITY_TRUNC Then
              pReturnValue = UniservPostReturnValues.ustTowns
              vSelection = pTownSelect
            ElseIf uni_get_ret_info(vResult) = UNI_SEL_STR Or uni_get_ret_info(vResult) = UNI_SEL_STR_TRUNC Then
              pReturnValue = UniservPostReturnValues.ustStreets
              vSelection = pStreetSelect
            ElseIf uni_get_ret_info(vResult) = UNI_SEL_BOX Or uni_get_ret_info(vResult) = UNI_SEL_BOX_TRUNC Then
              pReturnValue = UniservPostReturnValues.ustParts
              vSelection = pPartSelect
            End If
            'If there is a selection to choose from then go get the list
            If pReturnValue = UniservPostReturnValues.ustTowns Or pReturnValue = UniservPostReturnValues.ustStreets Or pReturnValue = UniservPostReturnValues.ustParts Then
              UniGetList(pReturnValue, vRequest, pItems)
              'If we have been here before we know which one to select so just select it
              If Len(vSelection) > 0 Then
                vResult = uni_set_arg(vRequest, "in_select_pos", vSelection, Len(vSelection) + 1, UNICSTR)
                vErrorNumber = UniPostError(vResult)
                GoTo labeltest
              End If
            End If
            'Just for the moment let's ignore them
            vResult = uni_close_request(vRequest)
            vErrorNumber = UniPostError(vResult)
          ElseIf uni_get_ret_type(vResult) = UNI_OK Or uni_get_ret_type(vResult) = UNI_WARN Then
            'Display results
            If Len(mvLastError) > 0 Then vLastError = mvLastError
            If Len(mvLastWarning) > 0 Then vLastWarning = mvLastWarning
            vArg = New String(Chr(0), 255) ' allocate space
            vResult = uni_get_arg(vRequest, OUT_ZIP, vArg, Len(vArg), UNICSTR)
            vErrorNumber = UniPostError(vResult)
            If vErrorNumber = UNI_OK Then
              pPostCode = ExtractArg(vArg)
              vArg = New String(Chr(0), 255) ' allocate space
              vResult = uni_get_arg(vRequest, OUT_CITY, vArg, Len(vArg), UNICSTR)
              vErrorNumber = UniPostError(vResult)
            End If
            If vErrorNumber = UNI_OK Then
              pTown = ExtractArg(vArg)
              vArg = New String(Chr(0), 255) ' allocate space
              vResult = uni_get_arg(vRequest, OUT_REGION, vArg, Len(vArg), UNICSTR)
              vErrorNumber = UniPostError(vResult)
            End If
            If vErrorNumber = UNI_OK Then
              pCounty = ExtractArg(vArg)
              vArg = New String(Chr(0), 255) 'allocate space
              vResult = uni_get_arg(vRequest, OUT_STR_HNO, vArg, Len(vArg), UNICSTR)
              vErrorNumber = UniPostError(vResult)
            End If
            If vErrorNumber = UNI_OK Then
              vReturnAdd = ExtractArg(vArg)
              If vLastLine.Length > 0 Then
                pAddress = Replace(pAddress, vLastLine, vReturnAdd)
              Else
                pAddress = vReturnAdd
              End If
            End If
            If vErrorNumber = UNI_OK Then
              vArg = New String(Chr(0), 255) 'allocate space
              vResult = uni_get_arg(vRequest, OUT_RES_STR, vArg, Len(vArg), UNISHORT)
              vErrorNumber = UniPostError(vResult)
              If vErrorNumber = UNI_OK And Asc(Left(vArg, 1)) = UNI_NOT_CHECKED Then
                pReturnValue = UniservPostReturnValues.uspNone
              Else
                pReturnValue = UniservPostReturnValues.ustCheckedOK
              End If
            End If
            'Close request
            vResult = uni_close_request(vRequest)
            vErrorNumber = UniPostError(vResult)
            If Len(mvLastError) = 0 And Len(vLastError) > 0 Then
              RaiseEvent SetErrorMessage(vLastError)
            ElseIf Len(mvLastWarning) = 0 And Len(mvLastWarning) > 0 Then
              RaiseEvent SetErrorMessage(vLastWarning)
            End If
          Else
            'Some error or warning has occured already handled so close the request
            vResult = uni_close_request(vRequest)
            vErrorNumber = UniPostError(vResult)
          End If
        End If
      End If
      RaiseEvent SetStatus("")
      UniPostCheck = vErrorNumber
    End Function

    Private Function UniPostError(ByVal pResult As Integer) As Integer
      Dim vResult As Integer
      Dim vErrorMsg As String
      Dim vType As Integer
      Dim vErrorNumber As Integer

      vType = uni_get_ret_type(pResult)
      If vType = UNI_ERR Then
        vErrorMsg = New String(" "c, 255) ' allocate space for error-message
        vResult = uni_get_error_msg(mvPost.Session, pResult, ERROR_LANGUAGE, vErrorMsg, Len(vErrorMsg))
        mvLastError = ExtractArg(vErrorMsg)
        RaiseEvent ShowError(mvLastError)
        vErrorNumber = uni_get_ret_info(pResult)
        If vErrorNumber < 1000 Then
          vResult = uni_close_session(mvPost.Session)
          mvPost.Active = False
        End If
        UniPostError = vErrorNumber
        If mvRaiseErrors Then RaiseError(DataAccessErrors.daeUniservError, mvLastError)
      ElseIf vType = UNI_WARN Then
        vErrorMsg = New String(" "c, 255) ' allocate space for error-message
        vResult = uni_get_error_msg(mvPost.Session, pResult, ERROR_LANGUAGE, vErrorMsg, Len(vErrorMsg))
        mvLastWarning = ExtractArg(vErrorMsg)
        RaiseEvent ShowWarning(mvLastWarning)
      End If
    End Function

    Public Function GetStreetNo(ByRef pAddress As String) As String
      Dim vFirstLine As String
      Dim vDone As Boolean
      Dim vStart As Integer
      Dim vPos As Integer

      vFirstLine = RTrim(FirstLine(pAddress))
      vStart = 1
      Do
        vPos = InStr(vStart, vFirstLine, " ")
        If vPos > 0 Then
          If IsNumeric(Mid(vFirstLine, vPos + 1, 1)) Then
            Return Mid(vFirstLine, vPos + 1)
            vDone = True
          End If
          vStart = vPos + 1
        End If
      Loop While vPos > 0 And Not vDone
      Return ""
    End Function

    Private Function LastAddressLine(ByRef pAddress As String) As String
      Dim vLines() As String
      Dim vIndex As Integer

      vLines = Split(Replace(pAddress, vbCr, ""), vbLf)
      vIndex = UBound(vLines)
      Do While vIndex >= 0
        If Len(Trim(vLines(vIndex))) > 0 Then
          Return Trim(vLines(vIndex))
          Exit Do
        End If
        vIndex = vIndex - 1
      Loop
      Return ""
    End Function

    Public ReadOnly Property LastErrorMessage As String
      Get
        Return mvLastError
      End Get
    End Property

  End Class
End Namespace
