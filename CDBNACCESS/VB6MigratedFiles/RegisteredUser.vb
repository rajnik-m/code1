Namespace Access
  Public Class RegisteredUser

    Public Enum RegisteredUserRecordSetTypes 'These are bit values
      rurtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum RegisteredUserFields
      rufAll = 0
      rufUserName
      rufPassword
      rufEmailAddress
      rufContactNumber
      rufLogOnCount
      rufLastLoggedOn
      rufCreatedOn
      rufRegistrationData
      rufSecurityQuestion
      rufSecurityAnswer
      rufLastUpdatedOn
      rufValidFrom
      rufValidTo
      rufLoginAttempts
      rufLockedOut
      rufPasswordExpiryDate
      rufAmendedBy
      rufAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvContact As Contact
    Private mvPasswordHistoryRecord As PortalPasswordHistory = Nothing

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "registered_users"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("user_name")
          .Add("password", CDBField.FieldTypes.cftBinary)
          .Add("email_address")
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("log_on_count", CDBField.FieldTypes.cftLong)
          .Add("last_logged_on", CDBField.FieldTypes.cftDate)
          .Add("created_on", CDBField.FieldTypes.cftTime)
          .Add("registration_data", CDBField.FieldTypes.cftMemo)
          .Add("security_question")
          .Add("security_answer")
          .Add("last_updated_on", CDBField.FieldTypes.cftTime)
          .Add("valid_from", CDBField.FieldTypes.cftDate)
          .Add("valid_to", CDBField.FieldTypes.cftDate)
          .Add("login_attempts", CDBField.FieldTypes.cftInteger)
          .Add("locked_out", CDBField.FieldTypes.cftTime)
          .Add("password_expiry_date", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftTime)
        End With
        mvClassFields.Item(RegisteredUserFields.rufUserName).SetPrimaryKeyOnly()

        'Removed checks to see if fields in database - no longer required
        mvClassFields.Item(RegisteredUserFields.rufPasswordExpiryDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbPasswordExpiry)
      Else
        mvClassFields.ClearItems()
        mvContact = Nothing
      End If
      mvClassFields.Item(RegisteredUserFields.rufContactNumber).PrefixRequired = True
      mvClassFields.Item(RegisteredUserFields.rufAmendedOn).PrefixRequired = True
      mvClassFields.Item(RegisteredUserFields.rufAmendedBy).PrefixRequired = True
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(RegisteredUserFields.rufLogOnCount).Value = CStr(1)
      Dim vPasswordExpiry As String = mvEnv.GetConfig("portal_password_duration")
      If IntegerValue(vPasswordExpiry) > 0 Then mvClassFields(RegisteredUserFields.rufPasswordExpiryDate).Value = DateAdd(DateInterval.Month, IntegerValue(vPasswordExpiry), Date.Today).ToString(CAREDateFormat)
    End Sub

    Private Sub SetValid(ByVal pField As RegisteredUserFields)
      'Add code here to ensure all values are valid before saving
      If Len(mvClassFields.Item(RegisteredUserFields.rufCreatedOn).Value) = 0 Then mvClassFields.Item(RegisteredUserFields.rufCreatedOn).Value = TodaysDate()
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As RegisteredUserRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = RegisteredUserRecordSetTypes.rurtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ru")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub InitByContactNumber(ByVal pEnv As CDBEnvironment, pContactNumber As Integer)
      mvEnv = pEnv
      InitClassFields()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", pContactNumber)
      Dim vSql As New SQLStatement(mvEnv.Connection, GetRecordSetFields(RegisteredUserRecordSetTypes.rurtAll), "registered_users ru", vWhereFields)
      Dim vRecordSet As CDBRecordSet = vSql.GetRecordSet()
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(mvEnv, vRecordSet, RegisteredUserRecordSetTypes.rurtAll)
      Else
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pUserName As String, Optional ByVal pMemberNumber As String = "", Optional ByVal pMemberContactNumber As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields As New CDBFields
      mvEnv = pEnv
      InitClassFields()
      If Not (String.IsNullOrEmpty(pMemberNumber) AndAlso String.IsNullOrEmpty(pMemberContactNumber)) Then
        vAnsiJoins.Add("members m", "ru.contact_number", "m.contact_number")
        If Not String.IsNullOrEmpty(pMemberNumber) Then
          vWhereFields.Add("m.member_number", pMemberNumber)
        Else
          'pMemberContactNumber set
          vWhereFields.Add("m.contact_number", pMemberContactNumber)
        End If
        vWhereFields.Add("m.cancellation_reason")
        Dim vSql As New SQLStatement(mvEnv.Connection, GetRecordSetFields(RegisteredUserRecordSetTypes.rurtAll), "registered_users ru", vWhereFields, "", vAnsiJoins)
        vRecordSet = vSql.GetRecordSet()
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, RegisteredUserRecordSetTypes.rurtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      ElseIf pUserName.Length > 0 Then
        vWhereFields.Add("user_name", pUserName)
        Dim vSql As New SQLStatement(mvEnv.Connection, GetRecordSetFields(RegisteredUserRecordSetTypes.rurtAll), "registered_users ru", vWhereFields)
        vRecordSet = vSql.GetRecordSet
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, RegisteredUserRecordSetTypes.rurtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As RegisteredUserRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And RegisteredUserRecordSetTypes.rurtAll) = RegisteredUserRecordSetTypes.rurtAll Then
          .SetItem(RegisteredUserFields.rufUserName, vFields)
          .SetItem(RegisteredUserFields.rufPassword, vFields)
          .SetItem(RegisteredUserFields.rufEmailAddress, vFields)
          .SetItem(RegisteredUserFields.rufContactNumber, vFields)
          .SetItem(RegisteredUserFields.rufLogOnCount, vFields)
          .SetItem(RegisteredUserFields.rufLastLoggedOn, vFields)
          .SetItem(RegisteredUserFields.rufCreatedOn, vFields)
          .SetOptionalItem(RegisteredUserFields.rufRegistrationData, vFields)
          .SetOptionalItem(RegisteredUserFields.rufSecurityQuestion, vFields)
          .SetOptionalItem(RegisteredUserFields.rufSecurityAnswer, vFields)
          .SetOptionalItem(RegisteredUserFields.rufLastUpdatedOn, vFields)
          .SetOptionalItem(RegisteredUserFields.rufValidFrom, vFields)
          .SetOptionalItem(RegisteredUserFields.rufValidTo, vFields)
          .SetOptionalItem(RegisteredUserFields.rufLoginAttempts, vFields)
          .SetOptionalItem(RegisteredUserFields.rufLockedOut, vFields)
          .SetOptionalItem(RegisteredUserFields.rufPasswordExpiryDate, vFields)
          .SetOptionalItem(RegisteredUserFields.rufAmendedBy, vFields)
          .SetOptionalItem(RegisteredUserFields.rufAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Delete(pEnv As CDBEnvironment)
      mvClassFields.Delete(pEnv.Connection)
    End Sub

    Public Sub Save(Optional ByRef pAudit As Boolean = False, Optional ByRef pAmendedBy As String = "", Optional ByRef pJournalNumber As Integer = 0)
      SetValid(RegisteredUserFields.rufAll)
      'mvClassFields.Save(mvEnv, mvExisting, "", pAudit)

      Dim vRetainPasswordHistoryControlNo As Integer = IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRetainRegUserPasswords))

      If vRetainPasswordHistoryControlNo > 1 AndAlso Existing Then
        Dim vTransaction As Boolean
        Try
          If Not mvEnv.Connection.InTransaction Then
            vTransaction = True
            mvEnv.Connection.StartTransaction()
          End If

          'Control value - 2 because we will be inserting the newest password from registered_users into portal_password_history
          DeletePasswordHistory(mvEnv, PasswordHistoryRecord.UserName, vRetainPasswordHistoryControlNo - 2)

          If Not String.IsNullOrEmpty(PasswordHistoryRecord.UserName) Then
            PasswordHistoryRecord.Save()
          End If

          mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit, pJournalNumber) 'overriding new one

          If vTransaction Then
            mvEnv.Connection.CommitTransaction()
          End If

        Catch vEx As Exception
          If vTransaction Then
            mvEnv.Connection.RollbackTransaction()
          End If
        End Try
      Else
        mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit, pJournalNumber) 'overriding new one
      End If
    End Sub

    Public Sub Create(ByRef pParams As CDBParameters, ByVal pContact As Contact, Optional ByRef pResgistrationData As String = "")

      Update(mvEnv, pParams)
      mvClassFields.Item(RegisteredUserFields.rufRegistrationData).Value = pResgistrationData
      If Not pContact Is Nothing Then
        mvClassFields.Item(RegisteredUserFields.rufContactNumber).Value = CStr(pContact.ContactNumber)
        RecordLogOn()
      Else
        mvClassFields.Item(RegisteredUserFields.rufLogOnCount).Value = CStr(0) 'Not logged on just created user record
        Save(False, mvEnv.User.UserID)
      End If
    End Sub

    ''' <summary>
    ''' Checks if the registered user's password matches the one passed in and that the user is not locked out
    ''' </summary>
    ''' <param name="pPassword"></param>
    ''' <returns>True - if the passwords match and the user is not locked out</returns>
    ''' <remarks></remarks>
    Public Function LogOn(ByVal pPassword As String, ByVal pPasswordKnown As Boolean) As Boolean
      Dim vValid As Boolean = False
      Dim vPassword As String = pPassword
      If Not mvEnv.GetConfigOption("portal_password_case_sensitive", True) Then
        vPassword = vPassword.ToUpper()
      End If
      'BR19442 encrypt and compare to database 
      If PasswordHash.PasswordHash.ValidatePassword(vPassword, mvClassFields.Item(RegisteredUserFields.rufPassword).ByteValue) Then
        vValid = True
      End If
      If vValid Then
        If IsLockedOut Then
          'Raise as invalid login
          vValid = False
        Else
          RecordLogOn()
        End If
      End If
      If Not vValid Then
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLoginLockout) Then
          'Failed login, increment attempts
          mvClassFields.Item(RegisteredUserFields.rufLoginAttempts).IntegerValue = LoginAttempts + 1
          'if maximum login attempts is configured then check and set locked out date
          If IntegerValue(mvEnv.GetConfig("portal_maximum_login_attempts")) > 0 Then
            If LoginAttempts >= IntegerValue(mvEnv.GetConfig("portal_maximum_login_attempts")) Then
              mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value = TodaysDateAndTime()
            End If
          End If
          Save()
          If pPasswordKnown Then
            Throw New Exception("Invalid password for user name already in use")
          End If
        End If
      End If
    End Function

    Public Sub RecordLogOn()
      mvClassFields.Item(RegisteredUserFields.rufLogOnCount).Value = CStr(LogOnCount + 1)
      mvClassFields.Item(RegisteredUserFields.rufLastLoggedOn).Value = TodaysDate()
      'Clear the attempts and locked out timestamp if the login attempt is successful
      mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value = String.Empty
      mvClassFields.Item(RegisteredUserFields.rufLoginAttempts).IntegerValue = 0
      Save()
    End Sub

    Public ReadOnly Property Contact() As Contact
      Get
        If mvContact Is Nothing Then
          mvContact = New Contact(mvEnv)
          mvContact.Init(ContactNumber)
        End If
        Contact = mvContact
      End Get
    End Property

    Public ReadOnly Property PasswordHistoryRecord() As PortalPasswordHistory
      Get
        If mvPasswordHistoryRecord Is Nothing Then
          mvPasswordHistoryRecord = New PortalPasswordHistory(mvEnv)
          mvPasswordHistoryRecord.Init()
        End If
        Return mvPasswordHistoryRecord
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property LoginAttempts As Integer
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufLoginAttempts).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LockedOut As String
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value
      End Get
    End Property

    ''' <summary>
    ''' Returns 'True' if the configured number of minutes has not elapsed after the user has exhausted the maximum login attempts
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsLockedOut As Boolean
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value.Length > 0 AndAlso (Now - CDate(mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value)).TotalMinutes < IntegerValue(mvEnv.GetConfig("portal_lockout_duration", "30"))
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(RegisteredUserFields.rufContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(RegisteredUserFields.rufCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property EmailAddress() As String
      Get
        EmailAddress = mvClassFields.Item(RegisteredUserFields.rufEmailAddress).Value
      End Get
    End Property

    Public ReadOnly Property LastLoggedOn() As String
      Get
        LastLoggedOn = mvClassFields.Item(RegisteredUserFields.rufLastLoggedOn).Value
      End Get
    End Property

    Public ReadOnly Property LogOnCount() As Integer
      Get
        LogOnCount = mvClassFields.Item(RegisteredUserFields.rufLogOnCount).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Password() As Byte()
      Get
        Password = mvClassFields.Item(RegisteredUserFields.rufPassword).ByteValue
      End Get
    End Property

    Public ReadOnly Property UserName() As String
      Get
        UserName = mvClassFields.Item(RegisteredUserFields.rufUserName).Value
      End Get
    End Property

    Public ReadOnly Property RegistrationData() As String
      Get
        RegistrationData = mvClassFields.Item(RegisteredUserFields.rufRegistrationData).Value
      End Get
    End Property

    Public ReadOnly Property SecurityQuestion() As String
      Get
        SecurityQuestion = mvClassFields.Item(RegisteredUserFields.rufSecurityQuestion).Value
      End Get
    End Property

    Public ReadOnly Property SecurityAnswer() As String
      Get
        SecurityAnswer = mvClassFields.Item(RegisteredUserFields.rufSecurityAnswer).Value
      End Get
    End Property

    Public ReadOnly Property LastUpdatedOn() As String
      Get
        LastUpdatedOn = mvClassFields.Item(RegisteredUserFields.rufLastUpdatedOn).Value
      End Get
    End Property

    Public ReadOnly Property ValidFrom() As String
      Get
        ValidFrom = mvClassFields.Item(RegisteredUserFields.rufValidFrom).Value
      End Get
    End Property

    Public ReadOnly Property ValidTo() As String
      Get
        ValidTo = mvClassFields.Item(RegisteredUserFields.rufValidTo).Value
      End Get
    End Property

    Public ReadOnly Property PasswordExpiryDate() As String
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufPasswordExpiryDate).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields.Item(RegisteredUserFields.rufAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property IsUserValid() As Boolean
      Get
        Dim vValidUser As Boolean = True
        If ValidFrom.Length > 0 AndAlso Date.Parse(ValidFrom) > Date.Parse(TodaysDate) Then vValidUser = False
        If vValidUser AndAlso ValidTo.Length > 0 AndAlso Date.Parse(ValidTo) < Date.Parse(TodaysDate) Then vValidUser = False
        If vValidUser AndAlso IsDate(PasswordExpiryDate) AndAlso CDate(PasswordExpiryDate) <= Date.Today Then vValidUser = False
        Return vValidUser
      End Get
    End Property

    Public Sub Update(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      If pParams.HasValue("UserName") Then mvClassFields.Item(RegisteredUserFields.rufUserName).Value = pParams("UserName").Value
      If pParams.HasValue("Password") Then
        'If the password has been passed in then it could be a create, a password update or a reset password
        Dim vPasswordExpiry As String = mvEnv.GetConfig("portal_password_duration")
        'If the duration config is set and no expiry date is entered
        If IntegerValue(vPasswordExpiry) > 0 AndAlso pParams.ContainsKey("PasswordExpiryDate") = False Then
          'Then update the expiry date
          mvClassFields(RegisteredUserFields.rufPasswordExpiryDate).Value = DateAdd(DateInterval.Month, IntegerValue(vPasswordExpiry), Date.Today).ToString(CAREDateFormat)
        End If

        Dim vPassword(48) As Byte
        If mvEnv.GetConfigOption("portal_password_case_sensitive", True) Then
          'Encrypt Password BR19442
          vPassword = PasswordHash.PasswordHash.CreateHash(pParams("Password").Value)
        Else
          vPassword = PasswordHash.PasswordHash.CreateHash(pParams("Password").Value.ToUpper)
        End If
        ''If existing then we are updating and not a new registered user
        ''Updating registered_users table so need to 
        ''compare to registered_user password
        If Existing Then
          Dim vPasswordString As String
          If mvEnv.GetConfigOption("portal_password_case_sensitive", True) Then
            'Encrypt Password BR19442
            vPasswordString = pParams("Password").Value
          Else
            vPasswordString = pParams("Password").Value.ToUpper
          End If

          'Only on update so know it exists so first compare to current password in registereg_users
          If PasswordHash.PasswordHash.ValidatePassword(vPasswordString, mvClassFields.Item(RegisteredUserFields.rufPassword).ByteValue) Then
            RaiseError(DataAccessErrors.daePasswordPreviouslyUsedInHistory)
          Else
            'Now check if the password is the same as the passwords in password history
            'First get config value for number of passwords strored in portal_password_history
            Dim vRetainPasswordHistoryControlNo As Integer
            vRetainPasswordHistoryControlNo = IntegerValue(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRetainRegUserPasswords))
            If vRetainPasswordHistoryControlNo > 1 Then
              Dim vPortalPasswordHistory As New PortalPasswordHistory(mvEnv)
              'Initial check of password history already done in registered_user table
              Dim vNumberPasswordHistoryToCheck As Integer = vRetainPasswordHistoryControlNo - 1
              Dim vPassWordExistsInHistoryTable As Boolean = False
              'Loop through the latest cdbControlRetainRegUserPasswords-1 records in the table to check if it already exist 
              Dim vUserName As String = If(pParams.Exists("UserName"), pParams("UserName").Value.ToString, pParams.ParameterExists("OldUserName").Value.ToString)
              vPassWordExistsInHistoryTable = CheckPasswordHistory(pEnv, vUserName, vPasswordString, vNumberPasswordHistoryToCheck)
                If vPassWordExistsInHistoryTable Then
                  RaiseError(DataAccessErrors.daePasswordPreviouslyUsedInHistory)
                End If

                Dim vParams = New CDBParameters()

                PasswordHistoryRecord.SetUserNameAndPassword(mvClassFields.Item(RegisteredUserFields.rufUserName).Value, mvClassFields.Item(RegisteredUserFields.rufPassword).ByteValue)

              End If
            End If

        End If
        Dim vParamName As String
        If Existing Then
          vParamName = String.Empty 'Updating registered_users table
        Else
          vParamName = "Password"   'Inserting into registered_users table
        End If
        mvClassFields.Item(RegisteredUserFields.rufPassword).DBParam = mvEnv.Connection.GetBinaryDBParameter("EncryptedPassword", vPassword, vParamName)
        mvClassFields.Item(RegisteredUserFields.rufPassword).Value = Convert.ToBase64String(vPassword)
      End If
      If pParams.HasValue("EmailAddress") Then mvClassFields.Item(RegisteredUserFields.rufEmailAddress).Value = pParams("EmailAddress").Value
      If pParams.Exists("SecurityQuestion") Then mvClassFields.Item(RegisteredUserFields.rufSecurityQuestion).Value = pParams("SecurityQuestion").Value
      If pParams.Exists("SecurityAnswer") Then mvClassFields.Item(RegisteredUserFields.rufSecurityAnswer).Value = pParams("SecurityAnswer").Value
      'BR19535
      mvClassFields.Item(RegisteredUserFields.rufLastUpdatedOn).Value = TodaysDateAndTime()
      If pParams.Exists("ValidFrom") Then mvClassFields.Item(RegisteredUserFields.rufValidFrom).Value = pParams("ValidFrom").Value
      If pParams.Exists("ValidTo") Then mvClassFields.Item(RegisteredUserFields.rufValidTo).Value = pParams("ValidTo").Value
      If pParams.Exists("LoginAttempts") Then mvClassFields.Item(RegisteredUserFields.rufLoginAttempts).Value = pParams("LoginAttempts").Value
      If pParams.Exists("LockedOut") Then mvClassFields.Item(RegisteredUserFields.rufLockedOut).Value = pParams("LockedOut").Value
      If pParams.Exists("PasswordExpiryDate") Then mvClassFields.Item(RegisteredUserFields.rufPasswordExpiryDate).Value = pParams("PasswordExpiryDate").Value
      mvClassFields.Item(RegisteredUserFields.rufAmendedBy).Value = mvEnv.User.UserID
      mvClassFields.Item(RegisteredUserFields.rufAmendedOn).Value = TodaysDateAndTime()
    End Sub

    Public Function CheckPasswordHistory(ByVal pEnv As CDBEnvironment, ByVal pUserName As String, ByVal pPassword As String, ByVal pNoHistoryToCheck As Integer) As Boolean

      Dim vPasswordHistory As New PortalPasswordHistory(pEnv)
      vPasswordHistory.Init()
      Dim vWhereFields As New CDBFields
      vWhereFields.Clear()
      vWhereFields.Add("pph.user_name", pUserName)
      Dim vSQL As New SQLStatement(pEnv.Connection, vPasswordHistory.GetRecordSetFields(), "portal_password_history pph", vWhereFields, "amended_on DESC")
      vSQL.MaxRows = pNoHistoryToCheck
      Dim vPasswordExistInHistoryTable As Boolean = False
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      While vRS.Fetch
        vPasswordHistory = New PortalPasswordHistory(pEnv)
        vPasswordHistory.InitFromRecordSet(vRS)
        If PasswordHash.PasswordHash.ValidatePassword(pPassword, vPasswordHistory.Password) Then
          vPasswordExistInHistoryTable = True
        End If
      End While
      vRS.CloseRecordSet()

      Return vPasswordExistInHistoryTable
    End Function

    Public Sub DeletePasswordHistory(ByVal pEnv As CDBEnvironment, ByVal pUserName As String, ByVal pNumHistoryToRetain As Integer)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("pph.user_name", pUserName)

      If pNumHistoryToRetain > 0 Then
        Dim vPasswordHistory As New PortalPasswordHistory(pEnv)
        vPasswordHistory.Init()

        Dim vSubSQL As New SQLStatement(pEnv.Connection, "pph.created_on", "portal_password_history pph", vWhereFields, "created_on DESC")
        vSubSQL.MaxRows = pNumHistoryToRetain
        Dim vSQL As New SQLStatement(pEnv.Connection, "MIN(pph.created_on)", String.Format("({0}) pph", vSubSQL.SQL), New CDBFields)
        Dim vLastRetainDate As String = vSQL.GetValue()

        vWhereFields = New CDBFields
        If Not String.IsNullOrEmpty(vLastRetainDate) AndAlso IsDate(vLastRetainDate) Then
          'Delete all by date > lastretaindate
          vWhereFields.Add("user_name", pUserName)
          vWhereFields.Add("created_on", CDBField.FieldTypes.cftTime, vLastRetainDate, CDBField.FieldWhereOperators.fwoLessThan)

          mvEnv.Connection.DeleteRecords("portal_password_history", vWhereFields, False)
        End If
      Else
        'Delete all records
        mvEnv.Connection.DeleteRecords("portal_password_history", vWhereFields, False)
      End If
    End Sub

  End Class
End Namespace
