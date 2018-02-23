Namespace Access

  Partial Public Class LicenseKey
    Private mvEnv As CDBEnvironment

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
        Case LicenceKeyModules.SmartClientCRM
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
            Case LicenceKeyModules.SmartClientCRM
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

    Public Function GetInitialisationKey() As String
      Return GenerateKey("", mvEnv.ClientCode, {1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1})
    End Function

  End Class

End Namespace