Namespace Access

  Partial Public Class CDBUser

    Public Enum OwnershipMessageTypes
      omtReadAccessIfSet = 0
      omtReadAccess
      omtAccessLevelDesc
    End Enum

    Public Enum AccessControlItems
      aciNewCreditCustomer
      aciRemoveCreditStopCode
      aciRunReports
      aciSelectionManager
      aciMemberMailing
      aciStandingOrderMailing
      aciDirectDebitMailing
      aciPayerMailing
      aciSubscriptionsMailing
      aciEventBookingsMailing
      aciEventAttendeesMailing
      aciEventPersonnelMailing
      aciEventVenueConfirmation
      aciOptionalMailingHistory
      aciContactDelete
      aciScheduleTasks
      aciToolbarMaintenance
      aciContactChangeStatus
      aciContactChangeDepartment
      aciContactChangeSource
      aciContactChangeOwnershipDetails
      aciEventSponsorsMailing
      ' BR11756
      aciServiceControlRestriction
    End Enum

    Private mvAddressNumber As Integer

    Public Function OwnershipAccessLevel(ByVal pNumber As Integer, ByVal pEntityGroup As EntityGroup, Optional ByRef pMessage As String = "", Optional ByRef pMessageType As OwnershipMessageTypes = OwnershipMessageTypes.omtReadAccess) As CDBEnvironment.OwnershipAccessLevelTypes
      'Return value is the ownership_access_level
      Dim vPrnUser As PrincipalUser
      Dim vRS As CDBRecordSet
      Dim vLevel As String = ""
      Dim vMessage As String
      Dim vSQL As String

      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vSQL = "SELECT read_access_text, view_access_text, ogu.ownership_access_level, ownership_access_level_desc, department_desc, ownership_group_desc"
        If pEntityGroup.EntityGroupType = EntityGroup.EntityGroupTypes.egtContact Then
          vSQL = vSQL & " FROM contacts c, ownership_groups og, departments d, ownership_group_users ogu, ownership_access_levels owl"
          vSQL = vSQL & " WHERE c.contact_number = " & pNumber & " AND og.ownership_group = c.ownership_group"
        ElseIf pEntityGroup.EntityGroupType = EntityGroup.EntityGroupTypes.egtOrganisation Then
          vSQL = vSQL & " FROM organisations o, ownership_groups og, departments d, ownership_group_users ogu, ownership_access_levels owl"
          vSQL = vSQL & " WHERE o.organisation_number = " & pNumber & " AND og.ownership_group = o.ownership_group"
        End If
        vSQL = vSQL & " AND d.department = og.principal_department AND ogu.ownership_group = og.ownership_group"
        vSQL = vSQL & " AND ogu.logname = '" & Logname & "' AND ogu.valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (TodaysDate()))
        vSQL = vSQL & " AND (ogu.valid_to IS NULL OR ogu.valid_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
        vSQL = vSQL & " AND owl.ownership_access_level = ogu.ownership_access_level"

        If pMessageType = OwnershipMessageTypes.omtAccessLevelDesc Then
          vMessage = "%access_level"
        Else
          vMessage = "You do not have access to this %1. Please contact your Database Administrator."
        End If
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        If vRS.Fetch Then
          vLevel = vRS.Fields("ownership_access_level").Value
          Select Case vLevel
            Case "B" 'Browse (1)
              If Len(vRS.Fields("view_access_text").Value) > 0 Then vMessage = vRS.Fields("view_access_text").Value
            Case "R" 'Read (2)
              If pMessageType = OwnershipMessageTypes.omtAccessLevelDesc Then
                'Nothing
              ElseIf pMessageType = OwnershipMessageTypes.omtReadAccess Then
                vMessage = vRS.Fields("read_access_text").Value
              ElseIf Len(vRS.Fields("read_access_text").Value) > 0 Then
                vMessage = vRS.Fields("read_access_text").Value
              End If
            Case "W" 'Write (3)
              If pMessageType <> OwnershipMessageTypes.omtAccessLevelDesc Then vMessage = ""
          End Select
          If vMessage.Length > 0 Then
            vMessage = Replace(vMessage, "%ownership_group", vRS.Fields("ownership_group_desc").Value)
            vMessage = Replace(vMessage, "%principal_department", vRS.Fields("department_desc").Value)
            vMessage = Replace(vMessage, "%access_level", vRS.Fields("ownership_access_level_desc").Value)
          End If
        End If
        vRS.CloseRecordSet()
        If InStr(1, vMessage, "%1") > 0 Then
          vMessage = Replace(vMessage, "%1", pEntityGroup.Name)
        End If
        If InStr(1, vMessage, "%principal_user") > 0 Then
          vPrnUser = New PrincipalUser
          vPrnUser.Init(mvEnv, pNumber)
          If vPrnUser.PrincipalUserName.Length > 0 Then
            vMessage = Replace(vMessage, "%principal_user", vPrnUser.PrincipalUserName)
          Else
            vMessage = Replace(vMessage, "%principal_user", "not allocated")
          End If
        End If
        pMessage = vMessage
        Return CDBEnvironment.GetOwnershipAccessLevel(vLevel)
      Else
        Return CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite 'default setting
      End If
    End Function

    Public Function HasItemAccessRights(ByVal pItem As AccessControlItems) As Boolean
      Dim vItem As String = ""
      Select Case pItem
        Case AccessControlItems.aciNewCreditCustomer
          vItem = "CDCSNC"
        Case AccessControlItems.aciRemoveCreditStopCode
          vItem = "CDCSRS"
        Case AccessControlItems.aciRunReports
          vItem = "SMFLRE"
        Case AccessControlItems.aciSelectionManager
          vItem = "CDTMSM"
        Case AccessControlItems.aciMemberMailing
          vItem = "SMMMME"
        Case AccessControlItems.aciStandingOrderMailing
          vItem = "SMMMSO"
        Case AccessControlItems.aciDirectDebitMailing
          vItem = "SMMMDD"
        Case AccessControlItems.aciPayerMailing
          vItem = "SMMMPA"
        Case AccessControlItems.aciSubscriptionsMailing
          vItem = "SMMMSU"
        Case AccessControlItems.aciEventBookingsMailing
          vItem = "SMMMEB"
        Case AccessControlItems.aciEventAttendeesMailing
          vItem = "SMMMEA"
        Case AccessControlItems.aciEventPersonnelMailing
          vItem = "SMMMEP"
        Case AccessControlItems.aciEventVenueConfirmation
          vItem = "CDEVVC"
        Case AccessControlItems.aciOptionalMailingHistory
          vItem = "GMMHCM"
        Case AccessControlItems.aciContactDelete
          vItem = "CDCMDE"
        Case AccessControlItems.aciScheduleTasks
          vItem = "GETMST"
        Case AccessControlItems.aciToolbarMaintenance
          vItem = "CDGETM"
        Case AccessControlItems.aciContactChangeStatus
          vItem = "CDCMCS"
        Case AccessControlItems.aciContactChangeDepartment
          vItem = "CDCMCD"
        Case AccessControlItems.aciContactChangeSource
          vItem = "CDCMCR"
        Case AccessControlItems.aciContactChangeOwnershipDetails
          vItem = "CDCMCO"
        Case AccessControlItems.aciEventSponsorsMailing
          vItem = "SMMMES"
          ' BR11756
        Case AccessControlItems.aciServiceControlRestriction
          vItem = "CDGESR"
      End Select
      If vItem.Length > 0 Then Return HasAccessRights(vItem)
    End Function

    Public ReadOnly Property AddressNumber() As Integer
      Get
        Dim vContact As New Contact(mvEnv)
        If mvAddressNumber = 0 Then
          If ContactNumber > 0 Then
            vContact.Init(ContactNumber)
            mvAddressNumber = vContact.AddressNumber
          End If
        End If
        Return mvAddressNumber
      End Get
    End Property

    ''' <summary>Build the Ownership SQL.</summary>
    ''' <param name="pAlias">Table alias to use.  E.g. c for contacts, o for organisations.</param>
    ''' <param name="pTables">Only add the required table names to the SQL.</param>
    ''' <param name="pPrnDept">Restrict the ownership_groups principal department to this value. Only used when <paramref name="pAccessLevelOnly"/> is True.</param>
    ''' <param name="pAccessLevelOnly">Only check access levels.</param>
    ''' <returns></returns>
    Public Function OwnershipSelect(ByVal pAlias As String, ByVal pTables As Boolean, Optional ByVal pPrnDept As String = "", Optional ByVal pAccessLevelOnly As Boolean = False) As String
      Dim vSQL As String = ""
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        If pTables Then
          If pAccessLevelOnly Then
            vSQL = ", ownership_group_users ogu"
          Else
            vSQL = " ,ownership_groups og, departments d, ownership_group_users ogu, ownership_access_levels oal"
          End If
        Else
          If pAccessLevelOnly Then
            vSQL = " AND " & pAlias & ".ownership_group = ogu.ownership_group"
          Else
            vSQL = " AND " & pAlias & ".ownership_group = og.ownership_group AND og.principal_department = d.department"
            If pPrnDept.Length > 0 Then vSQL = vSQL & " AND og.principal_department = '" & pPrnDept & "'"
            vSQL = vSQL & " AND og.ownership_group = ogu.ownership_group"
          End If
          vSQL = vSQL & " AND ogu.logname = '" & Logname & "'"
          vSQL = vSQL & " AND ogu.valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDate)
          vSQL = vSQL & " AND (ogu.valid_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, TodaysDate)
          vSQL = vSQL & " OR ogu.valid_to IS NULL)"
          If Not pAccessLevelOnly Then vSQL = vSQL & " AND ogu.ownership_access_level = oal.ownership_access_level"
        End If
      End If
      Return vSQL
    End Function

    ''' <summary>Build the Ownership SQL.</summary>
    ''' <param name="pAlias">Table alias to use.  E.g. c for contacts, o for organisations.</param>
    ''' <param name="pPrnDept">Restrict the ownership_groups principal department to this value. Only used when <paramref name="pAccessLevelOnly"/> is True.</param>
    ''' <param name="pAccessLevelOnly">Only check access levels.</param>
    ''' <param name="pAnsiJoins"></param>
    ''' <param name="pWhereFields"></param>
    Public Sub OwnershipSelect(ByVal pAlias As String, ByVal pPrnDept As String, ByVal pAccessLevelOnly As Boolean, ByRef pAnsiJoins As AnsiJoins, ByRef pWhereFields As CDBFields)
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        If pAnsiJoins Is Nothing Then pAnsiJoins = New AnsiJoins()
        If pWhereFields Is Nothing Then pWhereFields = New CDBFields()
        If String.IsNullOrWhiteSpace(pAlias) Then pAlias = String.Empty

        If pAccessLevelOnly Then
          pAnsiJoins.Add("ownership_group_users ogu", pAlias & ".ownership_group", "ogu.ownership_group")
        Else
          pAnsiJoins.Add("ownership_groups og", pAlias & ".ownership_group", "og.ownership_group")
          pAnsiJoins.Add("departments d", "og.principal_department", "d.department")
          pAnsiJoins.Add("ownership_group_users ogu", "og.ownership_group", "ogu.ownership_group")
          pAnsiJoins.Add("ownership_access_levels oal", "ogu.ownership_access_level", "oal.ownership_access_level")
          If Not String.IsNullOrWhiteSpace(pPrnDept) Then pWhereFields.Add("og.principal_department", pPrnDept)
        End If
        pWhereFields.Add("ogu.logname", Me.Logname)
        pWhereFields.Add("ogu.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThanEqual)
        pWhereFields.Add("ogu.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        pWhereFields.Add("ogu.valid_to#2", CDBField.FieldTypes.cftDate, String.Empty, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
    End Sub
  End Class

End Namespace
