Namespace Access

  Public Class ContactRole
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ContactRoleFields
      AllFields = 0
      ContactNumber
      OrganisationNumber
      Role
      ValidFrom
      ValidTo
      IsActive
      ContactRoleNumber
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("organisation_number", CDBField.FieldTypes.cftLong)
        .Add("role")
        .Add("valid_from", CDBField.FieldTypes.cftDate)
        .Add("valid_to", CDBField.FieldTypes.cftDate)
        .Add("is_active")
        .Add("contact_role_number", CDBField.FieldTypes.cftLong)

        .SetControlNumberField(ContactRoleFields.ContactRoleNumber, "RO")
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataContactRoleNumber) Then
          .Item(ContactRoleFields.ContactRoleNumber).PrimaryKey = True
          .Item(ContactRoleFields.ContactNumber).PrefixRequired = True
          .Item(ContactRoleFields.OrganisationNumber).PrefixRequired = True
          .Item(ContactRoleFields.Role).PrefixRequired = True
        Else
          .Item(ContactRoleFields.ContactRoleNumber).InDatabase = False
          .Item(ContactRoleFields.ContactNumber).PrimaryKey = True
          .Item(ContactRoleFields.OrganisationNumber).PrimaryKey = True
          .Item(ContactRoleFields.Role).PrimaryKey = True
        End If
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cr"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_roles"
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
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ContactRoleFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        Return mvClassFields(ContactRoleFields.OrganisationNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Role() As String
      Get
        Return mvClassFields(ContactRoleFields.Role).Value
      End Get
    End Property
    Public ReadOnly Property ValidFrom() As String
      Get
        Return mvClassFields(ContactRoleFields.ValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property ValidTo() As String
      Get
        Return mvClassFields(ContactRoleFields.ValidTo).Value
      End Get
    End Property
    Public ReadOnly Property IsActive() As String
      Get
        Return mvClassFields(ContactRoleFields.IsActive).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactRoleFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactRoleFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property ContactRoleNumber() As Integer
      Get
        Return mvClassFields(ContactRoleFields.ContactRoleNumber).IntegerValue
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"
    Private mvRoleDesc As String = ""
    Private mvToBeDeleted As Boolean

    Public Function CreateInstance(pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ContactRole(pEnv)
    End Function

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvRoleDesc = ""
      mvToBeDeleted = False
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      SetCurrent()
    End Sub

    Private Sub SetCurrent()
      If IsDate(ValidFrom) OrElse IsDate(ValidTo) Then
        Dim vActive As Boolean = True
        If IsDate(ValidFrom) AndAlso Date.Parse(ValidFrom) > Today Then vActive = False
        If IsDate(ValidTo) AndAlso Date.Parse(ValidTo) < Today Then vActive = False
        mvClassFields(ContactRoleFields.IsActive).Bool = vActive
      Else
        If mvClassFields(ContactRoleFields.IsActive).Value = "" Then mvClassFields(ContactRoleFields.IsActive).Value = "Y"
      End If
    End Sub

    Public Sub InitFromPosition(ByVal pCP As ContactPosition, ByVal pRole As String, Optional ByRef pValidFrom As String = "", Optional ByRef pValidTo As String = "")
      Init()
      mvClassFields(ContactRoleFields.ContactNumber).IntegerValue = pCP.ContactNumber
      mvClassFields(ContactRoleFields.OrganisationNumber).IntegerValue = pCP.OrganisationNumber
      mvClassFields(ContactRoleFields.Role).Value = pRole

      Dim vGotDates As Boolean = False
      If IsDate(pValidFrom) Then
        If IsDate(pCP.Started) AndAlso Date.Parse(pCP.Started) > Date.Parse(pValidFrom) Then pValidFrom = pCP.Started
        vGotDates = True
      Else
        pValidFrom = pCP.Started
      End If
      If IsDate(pValidTo) Then
        If IsDate(pCP.Finished) AndAlso Date.Parse(pCP.Finished) < Date.Parse(pValidTo) Then pValidTo = pCP.Finished
        vGotDates = True
      Else
        pValidTo = pCP.Finished
      End If
      If IsDate(pValidFrom) AndAlso IsDate(pValidTo) AndAlso Date.Parse(pValidFrom) > Date.Parse(pValidTo) AndAlso Date.Parse(pCP.Started) < Date.Parse(pValidTo) Then
        'the valid from date is later than the valid to date which is incorrect so use position date for valid from
        pValidFrom = pCP.Started
      End If

      If IsDate(pValidFrom) AndAlso IsDate(pCP.Finished) Then
        'If Role starts after Position finished then Role start & end will be Position finish
        If Date.Parse(pValidFrom).CompareTo(Date.Parse(pCP.Finished)) > 0 Then
          pValidFrom = pCP.Finished
          pValidTo = pCP.Started
        End If
      End If

      If IsDate(pValidTo) AndAlso IsDate(pCP.Started) Then
        'If Role finishes before Position starts then Role start & end will be Position start
        If Date.Parse(pValidTo).CompareTo(Date.Parse(pCP.Started)) < 0 Then
          pValidFrom = pCP.Started
          pValidTo = pCP.Finished
        End If
      End If

      'After all this make absolutely sure Role starts before it finishes
      If IsDate(pValidFrom) AndAlso IsDate(pValidTo) Then
        'If Role starts after it ends then start will be end
        If Date.Parse(pValidFrom).CompareTo(Date.Parse(pValidTo)) > 0 Then pValidFrom = pValidTo
      End If

      If Not vGotDates Then mvClassFields(ContactRoleFields.IsActive).Bool = pCP.Current
      Update(pValidFrom, pValidTo)
    End Sub

    Public Overloads Sub Init(ByVal pContactNumber As Integer, ByVal pOrganisationNumber As Integer, Optional ByVal pRole As String = "")
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      'BR15654 added this init to find out of the role record exists for data import
      Init()
      With vWhereFields
        If pContactNumber > 0 Then
          .Add(TableAlias & "." & mvClassFields.Item(ContactRoleFields.ContactNumber).Name, pContactNumber)
        End If
        If pOrganisationNumber > 0 Then
          .Add(TableAlias & "." & mvClassFields.Item(ContactRoleFields.OrganisationNumber).Name, pOrganisationNumber)
        End If
        If pRole.Length > 0 Then
          .Add(TableAlias & "." & mvClassFields.Item(ContactRoleFields.Role).Name, pRole)
        End If
      End With
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields(), DatabaseTableName & " " & TableAlias, vWhereFields, "")
      vRecordSet = vSQL.GetRecordSet
      If vRecordSet.Fetch() Then InitFromRecordSet(vRecordSet)
      vRecordSet.CloseRecordSet()
    End Sub

    Public Overloads Sub Update(ByVal pValidFrom As String, ByVal pValidTo As String)
      Update(pValidFrom, pValidTo, "")
    End Sub
    Public Overloads Sub Update(ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pRole As String)
      If pRole.Length > 0 Then mvClassFields(ContactRoleFields.Role).Value = pRole
      mvClassFields(ContactRoleFields.ValidFrom).Value = pValidFrom
      mvClassFields(ContactRoleFields.ValidTo).Value = pValidTo
      SetCurrent()
    End Sub

    Public Sub SetInactive()
      If IsDate(ValidFrom) OrElse IsDate(ValidTo) Then
        SetCurrent()
      Else
        mvClassFields(ContactRoleFields.IsActive).Bool = False
      End If
    End Sub

    Friend ReadOnly Property WillUpdate() As Boolean
      Get
        SetValid()
        Return mvClassFields.FieldsChanged
      End Get
    End Property

    Public ReadOnly Property RoleDesc() As String
      Get
        If mvRoleDesc.Length = 0 AndAlso Role.Length > 0 Then mvRoleDesc = mvEnv.GetDescription("roles", "role", Role)
        Return mvRoleDesc
      End Get
    End Property

    Public Property ToBeDeleted() As Boolean
      Get
        Return mvToBeDeleted
      End Get
      Set(ByVal Value As Boolean)
        mvToBeDeleted = Value
      End Set
    End Property

    Public Overrides Sub AddDeleteCheckItems()
      AddDeleteCheckItem("contact_position_timesheet", "contact_role_number", "Timesheet")
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateCreateParameters(pParameterList)
      If AmendedBy.Length > 0 AndAlso AmendedOn.Length > 0 Then mvOverrideAmended = True
    End Sub

    Public Sub SetActive(Optional ByVal pAmendedBy As String = "", Optional ByVal pAmendedOn As String = "")
      mvClassFields(ContactRoleFields.IsActive).Value = "Y"
      mvClassFields(ContactRoleFields.ValidTo).Value = ""
      mvClassFields(ContactRoleFields.AmendedBy).Value = pAmendedBy
      mvClassFields(ContactRoleFields.AmendedOn).Value = pAmendedOn
      If pAmendedBy.Length > 0 AndAlso pAmendedOn.Length > 0 Then mvOverrideAmended = True
    End Sub

    Public Function ValidateDates(ByVal pValidFrom As String, ByVal pValidTo As String) As Boolean
      Return ValidateDates(pValidFrom, pValidTo, False)
    End Function
    Public Function ValidateDates(ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pThrowErrors As Boolean) As Boolean
      Dim vValid As Boolean = True
      Dim vEndDate As String = ""
      Dim vStartDate As String = ""
      If IsDate(pValidFrom) OrElse IsDate(pValidTo) Then
        Dim vContact As New Contact(mvEnv)
        vContact.Init(ContactNumber)
        Dim vCheckDates As Boolean
        Dim vCount As Integer
        For Each vPosition As ContactPosition In vContact.GetPositions(0, OrganisationNumber)
          'Loop through and compare dates with Role when a break in the dates appears
          vCount = vCount + 1
          If vCount = 1 Then
            'First Position
            vStartDate = vPosition.Started
            vEndDate = vPosition.Finished
          Else
            'Next Position
            If IsDate(vPosition.Started) Then
              If IsDate(vEndDate) Then
                If DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, Date.Parse(vEndDate)) = Date.Parse(vPosition.Started) Then
                  vEndDate = vPosition.Finished
                Else
                  vCheckDates = True
                End If
              End If
            End If
          End If

          If vCheckDates Then
            vValid = CheckDates(pValidFrom, pValidTo, vStartDate, vEndDate)
            If vValid Then
              If (vCount - 1) = 1 Then
                If IsDate(pValidTo) AndAlso IsDate(vStartDate) Then
                  If Date.Parse(pValidTo) < Date.Parse(vStartDate) Then
                    'Role finishes before first Position starts
                    vValid = False
                  End If
                End If
              End If
              If vValid Then
                'Check to see if Role comes completely between the two Positions
                If IsDate(pValidFrom) AndAlso IsDate(pValidTo) AndAlso IsDate(vEndDate) AndAlso IsDate(vPosition.Started) Then
                  If (Date.Parse(pValidFrom) > Date.Parse(vEndDate)) AndAlso (Date.Parse(pValidTo) < Date.Parse(vPosition.Started)) Then
                    vValid = False
                  End If
                End If
              End If
            End If
            If vValid Then
              'There was a gap in the dates that was valid, so reset the dates
              vStartDate = vPosition.Started
              vEndDate = vPosition.Finished
            End If
          End If

          If vValid = False Then Exit For
        Next vPosition

        If vValid Then vValid = CheckDates(pValidFrom, pValidTo, vStartDate, vEndDate)
        If vValid Then
          If vCount = 1 Then
            If IsDate(pValidTo) AndAlso IsDate(vStartDate) Then
              If Date.Parse(pValidTo) < Date.Parse(vStartDate) Then
                'Role finishes before first Position starts
                vValid = False
              End If
            End If
          End If
          If vValid Then
            If IsDate(vEndDate) AndAlso IsDate(pValidFrom) Then
              If Date.Parse(pValidFrom) > Date.Parse(vEndDate) Then
                'Role starts after last Position finishes
                vValid = False
              End If
            End If
          End If
        End If
      End If

      If vValid = False AndAlso pThrowErrors = True Then
        If IsDate(vStartDate) OrElse IsDate(vEndDate) Then
          RaiseError(DataAccessErrors.daeRoleDatesExceedStatedPositionDates, If(IsDate(vStartDate), CDate(vStartDate).ToString(CAREDateFormat), "<NULL>"), If(IsDate(vEndDate), CDate(vEndDate).ToString(CAREDateFormat), "<NULL>"))
        Else
          RaiseError(DataAccessErrors.daeRoleDatesExceedPositionDates)
        End If
      End If

      If vValid Then
        'Check Timesheets
        vValid = CheckTimesheetDates(pValidFrom, pValidTo, pThrowErrors)
      End If

      Return vValid

    End Function

    Private Function CheckDates(ByVal pRoleFrom As String, ByVal pRoleTo As String, ByVal pPosStart As String, ByVal pPosEnd As String) As Boolean
      Dim vValid As Boolean = True
      If IsDate(pRoleFrom) AndAlso IsDate(pPosStart) Then
        'Both have Start dates
        If Date.Parse(pPosStart) > Date.Parse(pRoleFrom) Then
          'Position starts after Role
          If IsDate(pRoleTo) Then
            If Date.Parse(pRoleTo) >= Date.Parse(pPosStart) Then vValid = False 'Role ends after Position starts
          Else
            'Role has no end date
            vValid = False
          End If
        Else
          'Position starts on/before Role
          If IsDate(pPosEnd) Then
            If Date.Parse(pPosEnd) >= Date.Parse(pRoleFrom) Then
              'Position ends after Role starts
              If IsDate(pRoleTo) Then
                If Date.Parse(pRoleTo) > Date.Parse(pPosEnd) Then vValid = False 'Role ends after Position ends
              Else
                'Role has no end date
                vValid = False
              End If
            End If
          End If
        End If
      End If

      If IsDate(pPosStart) AndAlso Not IsDate(pRoleFrom) Then
        'Position has start but Role does not
        If IsDate(pPosEnd) Then vValid = False 'Position has an end date
      End If

      If vValid AndAlso Not (IsDate(pPosStart) AndAlso IsDate(pRoleFrom)) Then
        'Neither have start dates
        If IsDate(pRoleTo) Then
          If IsDate(pPosEnd) Then
            'Both have end dates
            If Date.Parse(pRoleTo) > Date.Parse(pPosEnd) Then vValid = False 'Role ends after Position
          End If
        Else
          If IsDate(pPosEnd) Then vValid = False 'Role ends after Position
        End If
      End If

      Return vValid

    End Function

    Private Function CheckTimesheetDates(ByVal pRoleFrom As String, ByVal pRoleTo As String, ByVal pThrowErrors As Boolean) As Boolean
      Dim vValidFrom As Nullable(Of Date) = Nothing
      Dim vValidTo As Nullable(Of Date) = Nothing
      Dim vIsValid As Boolean = True

      If IsDate(pRoleFrom) Then vValidFrom = Date.Parse(pRoleFrom)
      If IsDate(pRoleTo) Then vValidTo = Date.Parse(pRoleTo)

      Dim vTimesheetDate As Date
      If vValidFrom.HasValue OrElse vValidTo.HasValue Then
        'Role has dates so check Timesheet
        Dim vWhereFields As New CDBFields(New CDBField("contact_role_number", ContactRoleNumber))
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "timesheet_number, contact_position_number, contact_role_number, timesheet_date", "contact_position_timesheet cpt", vWhereFields)
        Dim vDT As DataTable = mvEnv.Connection.GetDataTable(vSQLStatement)
        If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
          For Each vRow As DataRow In vDT.Rows
            vTimesheetDate = Date.Parse(vRow.Item("timesheet_date").ToString)
            If vValidFrom.HasValue AndAlso vValidFrom.Value.CompareTo(vTimesheetDate) > 0 Then vIsValid = False 'Role starts after Timesheet date
            If vValidTo.HasValue AndAlso vValidTo.Value.CompareTo(vTimesheetDate) < 0 Then vIsValid = False 'Role end after Timesheet date
            If vIsValid = False Then Exit For
          Next
        End If
      Else
        'Role has not dates so Timesheet date will be valid
        vIsValid = True
      End If

      If vIsValid = False And pThrowErrors = True Then
        RaiseError(DataAccessErrors.daeRoleDatesInvalidForTimesheet, vTimesheetDate.ToString(CAREDateFormat))
      End If

      Return vIsValid

    End Function

#End Region

  End Class
End Namespace
