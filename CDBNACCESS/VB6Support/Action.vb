Imports System.Linq

Namespace Access


  Partial Public Class Action

    Public Enum ActionerManagerSettings
      amsAsDefined = 1
      amsAsCreator
      amsAsk
    End Enum

    Public Enum ActionsScheduleTypes
      astGivenDate
      astSoonest
      astToday
      astTomorrow
      astThisWeek
      astNextWeek
      astNextMonth
    End Enum

    Private Enum SetTimeTypes
      sttNoSeconds
      sttStartofDay
      sttEndOfDay
    End Enum

    Private Const AMS_DEFINED As String = "N"
    Private Const AMS_CREATOR As String = "C"
    Private Const AMS_ASK As String = "A"

    Private mvDependancies As Collection
    Private mvSubjects As Collection

    Protected Overrides Sub ClearFields()
      mvDependancies = Nothing
      mvSubjects = Nothing
    End Sub

    Public Function GetActionLinkObjectCode(ByVal pContact As Contact) As IActionLink.ActionLinkObjectTypes
      Dim vLinkCode As IActionLink.ActionLinkObjectTypes
      Select Case pContact.ContactType
        Case Contact.ContactTypes.ctcOrganisation
          vLinkCode = IActionLink.ActionLinkObjectTypes.alotOrganisation
        Case Else
          vLinkCode = IActionLink.ActionLinkObjectTypes.alotContact
      End Select
      Return vLinkCode
    End Function

    Friend ReadOnly Property PriorActions() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vPriorAction As Integer

        If mvDependancies Is Nothing Then
          mvDependancies = New Collection
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT prior_action FROM action_dependencies WHERE action_number = " & ActionNumber)
          While vRecordSet.Fetch()
            vPriorAction = vRecordSet.Fields(1).LongValue
            mvDependancies.Add(vPriorAction)
          End While
          vRecordSet.CloseRecordSet()
        End If
        Return mvDependancies
      End Get
    End Property

    Public Sub SetPriorActions(ByVal pPriorActionList As String)
      If mvExisting Then
        'Delete any existing prior actions
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("action_number", ActionNumber)
        mvEnv.Connection.DeleteRecords("action_dependencies", vWhereFields, False)
      End If
      Dim vInsertFields As New CDBFields
      vInsertFields.Add("action_number", ActionNumber)
      vInsertFields.Add("prior_action", ActionNumber)
      vInsertFields.AddAmendedOnBy(mvEnv.User.UserID)
      Dim vPriorActions() As String = pPriorActionList.Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      For Each vItem As String In vPriorActions
        vInsertFields(2).Value = vItem
        mvEnv.Connection.InsertRecord("action_dependencies", vInsertFields)
      Next
    End Sub

    Public Sub SetCancelled()
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vStatusChanged As Boolean

      vStatusChanged = ActionStatus <> ActionStatuses.astCancelled
      If vStatusChanged Then
        mvEnv.Connection.StartTransaction()
        vAppointment.Init()
        vAppointment.SetEntryStatus(ContactAppointment.ContactAppointmentTypes.catAction, ActionNumber, ContactAppointment.ContactAppointmentTimeStatuses.catsFree)
        mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astCancelled)
        Save(mvEnv.User.UserID, True)
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Friend ReadOnly Property Subjects() As Collection
      Get
        If mvSubjects Is Nothing Then
          mvSubjects = New Collection
          Dim vActionSubject As New ActionSubject(mvEnv)
          vActionSubject.Init()

          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vActionSubject.GetRecordSetFields, "action_subjects acs", New CDBFields(New CDBField("action_number", ActionNumber)))
          Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
          While vRS.Fetch
            vActionSubject = New ActionSubject(mvEnv)
            vActionSubject.InitFromRecordSet(vRS)
            mvSubjects.Add(vActionSubject)
          End While
          vRS.CloseRecordSet()
        End If
        Return mvSubjects
      End Get
    End Property

    Friend ReadOnly Property ActionerCount() As Integer
      Get
        Dim vLink As ActionLink
        Dim vCount As Integer
        For Each vLink In Links
          If vLink.LinkType = IActionLink.ActionLinkTypes.altActioner Then vCount += 1
        Next vLink
        Return vCount
      End Get
    End Property

    Friend ReadOnly Property ManagerCount() As Integer
      Get
        Dim vLink As ActionLink
        Dim vCount As Integer
        For Each vLink In Links
          If vLink.LinkType = IActionLink.ActionLinkTypes.altManager Then vCount += 1
        Next vLink
        Return vCount
      End Get
    End Property

    ''' <summary>Create a new <see cref="Action">Action</see> from the specified Proforma (Template) <see cref="Action">Action</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProForma">The Proforma (Template) <see cref="Action">Action</see> used to create this <see cref="Action">Action</see>.</param>
    ''' <param name="pMasterActionNumber">The master Action Number to be used.  The first <see cref="Action">Action</see> in the <see cref="ActionSet">ActionSet</see> is the master Action.</param>
    ''' <param name="pActionNumber">The Action number to be assigned to this <see cref="Action">Action</see>.</param>
    ''' <param name="pRelatedType">The type of object for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    ''' <param name="pActionerType">The type of object for an Actioner link to be created for each Action.</param>
    ''' <param name="pActionerNumber">The number of the Actioner to be created for each Action.</param>
    ''' <param name="pRelatedDocument">The number of a Document to be linked to each Action.</param>
    ''' <param name="pRelatedExamCentreId">The number of an Exam Centre to be linked to each Action.</param>
    Friend Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProForma As Action, ByVal pMasterActionNumber As Integer, ByVal pActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer, ByVal pActionerType As IActionLink.ActionLinkObjectTypes, ByVal pActionerNumber As Integer, ByVal pRelatedDocument As Integer, ByVal pRelatedExamCentreId As Integer)
      CreateFromProForma(pEnv, pProForma, pMasterActionNumber, pActionerNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, pRelatedDocument, pRelatedExamCentreId, Date.Today, True)
    End Sub
    ''' <summary>Create a new <see cref="Action">Action</see> from the specified Proforma (Template) <see cref="Action">Action</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProForma">The Proforma (Template) <see cref="Action">Action</see> used to create this <see cref="Action">Action</see>.</param>
    ''' <param name="pMasterActionNumber">The master Action Number to be used.  The first <see cref="Action">Action</see> in the <see cref="ActionSet">ActionSet</see> is the master Action.</param>
    ''' <param name="pActionNumber">The Action number to be assigned to this <see cref="Action">Action</see>.</param>
    ''' <param name="pRelatedType">The type of object for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    ''' <param name="pActionerType">The type of object for an Actioner link to be created for each Action.</param>
    ''' <param name="pActionerNumber">The number of the Actioner to be created for each Action.</param>
    ''' <param name="pRelatedDocument">The number of a Document to be linked to each Action.</param>
    ''' <param name="pRelatedExamCentreId">The number of an Exam Centre to be linked to each Action.</param>
    ''' <param name="pProcessingDate">The base date to use for calculating the new Action dates.</param>
    ''' <param name="pCreateAction">True to save the new <see cref="Action">Action</see> and related data in the database, otherwise False to set the Action data without saving.</param>
    Friend Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProForma As Action, ByVal pMasterActionNumber As Integer, ByVal pActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer, ByVal pActionerType As IActionLink.ActionLinkObjectTypes, ByVal pActionerNumber As Integer, ByVal pRelatedDocument As Integer, ByVal pRelatedExamCentreId As Integer, ByVal pProcessingDate As Date, ByVal pCreateAction As Boolean)
      mvEnv = pEnv
      InitClassFields()
      SetDefaults()

      'First set all the fields in this new Action to the values for the Proforma (Template) Action
      Me.Clone(pProForma, pActionNumber)

      'Now set any fields that need to be different
      Dim vRepeatRequired As Boolean = False
      Dim vDays As Integer = 0
      Dim vMonths As Integer = 0
      With pProForma
        mvClassFields.Item(ActionFields.MasterAction).IntegerValue = pMasterActionNumber
        mvClassFields.Item(ActionFields.ActionTemplateNumber).IntegerValue = .ActionNumber
        mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astInactive)
        If .RepeatDelay > 0 Then
          mvClassFields.Item(ActionFields.CompletedOn).Value = pProcessingDate.AddDays(.RepeatDelay).ToString(CAREDateFormat)
        Else
          mvClassFields.Item(ActionFields.CompletedOn).Value = TodaysDate()
        End If

        Dim vCreatedOn As String = .CreatedOn
        Dim vDeadline As String = .Deadline
        Dim vScheduledOn As String = .ScheduledOn
        Dim vCompletedOn As String = .CompletedOn

        'Set the deadline as an offset from the created on date of the proforma
        If vDeadline.Length > 0 Then
          If DeadlineDays.HasValue = True OrElse DeadlineMonths.HasValue Then
            vDays = 0
            If DeadlineMonths.HasValue = True AndAlso DeadlineMonths.Value <> 0 Then vMonths = DeadlineMonths.Value
            If DeadlineDays.HasValue = True AndAlso DeadlineDays.Value <> 0 Then vDays = DeadlineDays.Value
          Else
            vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vCreatedOn), CDate(vDeadline)))
          End If
          vDeadline = CalculateNewDate(pProcessingDate, vDays, vMonths, UseNegativeOffsets).ToString(CAREDateFormat) & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay)
        End If
        If vDeadline.Length > 0 Then vDeadline = CDate(vDeadline).ToString(CAREDateTimeFormat)
        mvClassFields.Item(ActionFields.Deadline).Value = vDeadline

        If .DelayedActivation Then 'If delayed
          If .PriorActions.Count() > 0 Then
            'Delayed from dependant - leave delayed flag set - update date to reflect offset
          Else
            'Delayed from creation - clear delayed cos we now know the date
            mvClassFields.Item(ActionFields.DelayedActivation).Bool = False
          End If
          If vScheduledOn.Length > 0 Then
            vDays = 0
            vMonths = 0
            If DelayDays.HasValue = True OrElse DelayMonths.HasValue = True Then
              If DelayDays.HasValue Then vDays = DelayDays.Value
              If DelayMonths.HasValue Then vMonths = DelayMonths.Value
            Else
              vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vCreatedOn), CDate(vScheduledOn)))
            End If
            vScheduledOn = CalculateNewDate(pProcessingDate, vDays, vMonths, UseNegativeOffsets).ToString(CAREDateFormat) & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay)
          End If
          If vScheduledOn.Length > 0 Then vScheduledOn = CDate(vScheduledOn).ToString(CAREDateTimeFormat)
          mvClassFields.Item(ActionFields.ScheduledOn).Value = vScheduledOn
        Else
          'Not delayed activation
          If .PriorActions.Count() > 0 Then
            mvClassFields.Item(ActionFields.ScheduledOn).Value = String.Empty 'Dependant - set scheduled to null
          Else
            mvClassFields.Item(ActionFields.ScheduledOn).Value = .ActivationDate
            If mvClassFields.Item(ActionFields.ScheduledOn).Value.Length = 0 Then
              mvClassFields.Item(ActionFields.RepeatCount).Value = String.Empty 'Immediate - repeat count null
              mvClassFields.Item(ActionFields.ScheduledOn).Value = String.Empty 'Immediate - set scheduled to null
              mvClassFields.Item(ActionFields.CompletedOn).Value = String.Empty 'Immediate - set completed to null
              mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astDefined)
              'Check if the repeat date is set (different to created on)
              vDays = 0
              vMonths = 0
              If IsDate(vCompletedOn) Then
                If RepeatDays.HasValue = True OrElse RepeatMonths.HasValue = True Then
                  If RepeatDays.HasValue Then vDays = RepeatDays.Value
                  If RepeatMonths.HasValue Then vMonths = RepeatMonths.Value
                Else
                  vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vCreatedOn), CDate(vCompletedOn)))
                End If
                If vDays > 0 OrElse vMonths > 0 Then vRepeatRequired = True
              End If
            End If
          End If
        End If

        'Ensure Actioner & Manager settings are either set or set to the default values
        ActionerSetting = .ActionerSetting
        ManagerSetting = .ManagerSetting

        'Ensure these get set to zero if they are null
        If mvClassFields.Item(ActionFields.DurationDays).Value.Length = 0 Then mvClassFields.Item(ActionFields.DurationDays).Value = "0"
        If mvClassFields.Item(ActionFields.DurationHours).Value.Length = 0 Then mvClassFields.Item(ActionFields.DurationHours).Value = "0"
        If mvClassFields.Item(ActionFields.DurationMinutes).Value.Length = 0 Then mvClassFields.Item(ActionFields.DurationMinutes).Value = "0"
        If mvClassFields.Item(ActionFields.RepeatCount).Value.Length = 0 Then mvClassFields.Item(ActionFields.RepeatCount).Value = "0"

        'If these fields have not changed then set them to null
        If mvClassFields.Item(ActionFields.CompletedOn).ValueChanged = False Then mvClassFields.Item(ActionFields.CompletedOn).Value = String.Empty
        If mvClassFields.Item(ActionFields.RepeatCount).ValueChanged = False Then mvClassFields.Item(ActionFields.RepeatCount).Value = String.Empty
        If mvClassFields.Item(ActionFields.ScheduledOn).ValueChanged = False Then mvClassFields.Item(ActionFields.ScheduledOn).Value = String.Empty
      End With

      Dim vTrans As Boolean = False
      If pEnv.Connection.InTransaction = False Then
        'Ensure we do all this inside a transaction
        pEnv.Connection.StartTransaction()
        vTrans = True
      End If

      'Save this new Action
      SetValid()
      If pCreateAction Then Save(mvEnv.User.UserID, True) 'If we are creating provisional Actions, don't save

      '.. and any Subjects & Links
      If pCreateAction Then
        'If we are creating provisional Actions, don't save
        With pProForma
          CopyActionSubjects(.Subjects, ActionNumber)
          CopyActionLinks(.Links)
        End With
      End If

      'Set Actioner and Manager
      If pCreateAction Then
        'If we are creating provisional Actions, don't save
        If ActionerCount = 0 AndAlso ActionerSetting = ActionerManagerSettings.amsAsCreator Then AddLink(IActionLink.ActionLinkObjectTypes.alotContact, IActionLink.ActionLinkTypes.altActioner, mvEnv.User.ContactNumber)
        If ManagerCount = 0 AndAlso ManagerSetting = ActionerManagerSettings.amsAsCreator Then AddLink(IActionLink.ActionLinkObjectTypes.alotContact, IActionLink.ActionLinkTypes.altManager, mvEnv.User.ContactNumber)
      End If

      'Add other Links
      If pCreateAction Then
        'If we are creating provisional Actions, don't save
        If pRelatedNumber > 0 Then AddLink(pRelatedType, IActionLink.ActionLinkTypes.altRelated, pRelatedNumber)
        If pActionerNumber > 0 Then AddLink(pActionerType, IActionLink.ActionLinkTypes.altActioner, pActionerNumber)
        If pRelatedDocument > 0 Then AddLink(IActionLink.ActionLinkObjectTypes.alotDocument, IActionLink.ActionLinkTypes.altRelated, pRelatedDocument)
        If pRelatedExamCentreId > 0 Then AddLink(IActionLink.ActionLinkObjectTypes.alotExamCentre, IActionLink.ActionLinkTypes.altRelated, pRelatedExamCentreId)
      End If

      'Now handle any actions with Repeat required here
      If pCreateAction = True AndAlso vRepeatRequired = True Then
        Dim vRepeatActionNumber As Integer = mvEnv.GetControlNumber("AC")
        Dim vRepeatAction As New Action(mvEnv)
        vRepeatAction.CloneForRepeatAction(Me, vRepeatActionNumber, pProcessingDate, vDays, vMonths)
        vRepeatAction.Save(mvEnv.User.UserID, True)

        CopyActionSubjects(pProForma.Subjects, vRepeatActionNumber)
        vRepeatAction.CopyActionLinks(pProForma.Links)
      End If

      If vTrans Then pEnv.Connection.CommitTransaction()

    End Sub

    Public Sub AddLink(ByVal pLinkObjectType As IActionLink.ActionLinkObjectTypes, ByVal pLinkType As IActionLink.ActionLinkTypes, ByVal pNumber As Integer)
      AddLink(pLinkObjectType, pLinkType, pNumber, 0)
    End Sub

    Public Sub AddLink(ByVal pLinkObjectType As IActionLink.ActionLinkObjectTypes, ByVal pLinkType As IActionLink.ActionLinkTypes, ByVal pNumber As Integer, ByVal pAdditionalNumber As Integer)
      Dim vLink As IActionLink = Nothing
      Select Case pLinkObjectType
        Case IActionLink.ActionLinkObjectTypes.alotExamCentre
          vLink = New ExamCentreAction(mvEnv)
        Case IActionLink.ActionLinkObjectTypes.alotWorkstream
          vLink = New WorkstreamActionLink(mvEnv)
        Case IActionLink.ActionLinkObjectTypes.alotContactPosition
          vLink = New ContactPositionAction(mvEnv)
        Case Else
          vLink = New ActionLink(mvEnv)
      End Select

      Dim vNotified As String = ""
      Dim vContact As Contact
      Dim vType As JournalTypes

      'Ignore setting if the contact number is unknown
      If pNumber > 0 Then
        'Notify actioners and managers if active action and then link is not to the user
        If ActionStatus = ActionStatuses.astDefined And
           pLinkObjectType = IActionLink.ActionLinkObjectTypes.alotContact And
           pLinkType <> IActionLink.ActionLinkTypes.altRelated And
           pNumber <> mvEnv.User.ContactNumber Then
          vNotified = "N"
        End If
        If vLink.GetType() Is GetType(ExamCentreAction) Then
          DirectCast(vLink, ExamCentreAction).InitFromParams(mvEnv, pLinkObjectType, ActionNumber, pNumber, pLinkType)
        ElseIf vLink.GetType() Is GetType(WorkstreamActionLink) Then
          DirectCast(vLink, WorkstreamActionLink).InitFromParams(mvEnv, pLinkObjectType, ActionNumber, pNumber, pLinkType)
        ElseIf vLink.GetType() Is GetType(ContactPositionAction) Then
          DirectCast(vLink, ContactPositionAction).InitFromParams(mvEnv, pLinkObjectType, ActionNumber, pNumber, pLinkType)
        Else
          DirectCast(vLink, ActionLink).InitFromParams(mvEnv, pLinkObjectType, ActionNumber, pNumber, pLinkType, vNotified, pAdditionalNumber)
        End If
        vLink.Save(mvEnv.User.UserID, True)
        If Not mvLinks Is Nothing Then mvLinks.Add(vLink)
        If ActionStatus = ActionStatuses.astDefined And (pLinkObjectType = IActionLink.ActionLinkObjectTypes.alotContact Or pLinkObjectType = IActionLink.ActionLinkObjectTypes.alotOrganisation) Then
          vContact = New Contact(mvEnv)
          vContact.Init(pNumber)
          Select Case pLinkType
            Case IActionLink.ActionLinkTypes.altActioner
              vType = JournalTypes.jnlActionActioner
            Case IActionLink.ActionLinkTypes.altManager
              vType = JournalTypes.jnlActionManager
            Case IActionLink.ActionLinkTypes.altRelated
              vType = JournalTypes.jnlActionRelated
          End Select
          mvEnv.AddJournalRecord(vType, JournalOperations.jnlActive, vContact.ContactNumber, vContact.AddressNumber, ActionNumber)
        End If
      End If
    End Sub

    Public ReadOnly Property RepeatDelay() As Integer
      Get 'Returns delay in days
        If ActionStatus = ActionStatuses.astProForma Then
          Dim vCreatedOn As String = CreatedOn
          Dim vCompletedOn As String = CompletedOn
          If IsDate(vCreatedOn) And IsDate(vCompletedOn) Then
            Return CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vCreatedOn), CDate(vCompletedOn)))
          End If
        End If
      End Get
    End Property

    Public Property ActivationDate() As String
      Get
        If ActionStatus = ActionStatuses.astProForma And PriorActions.Count() = 0 Then
          Dim vCreatedOn As String = CreatedOn
          Dim vScheduledOn As String = ScheduledOn
          If IsDate(vCreatedOn) And IsDate(vScheduledOn) Then
            Dim vDays As Integer = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vCreatedOn), CDate(vScheduledOn)))
            If DelayedActivation Then
              Return CDate(TodaysDate() & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay)).AddDays(vDays).ToString(CAREDateTimeFormat)
            Else
              'Check if the date is set (different to created on)
              If vDays <> 0 Then
                'This is a specified date - make sure it is in the future
                Dim vScheduledTime As String = CDate(vScheduledOn).ToString("HH:mm:ss")
                vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, Today, CDate(CDate(vScheduledOn).ToString("dd/MM/" & Year(Today)))))
                If vDays < 0 Then
                  'It is in the past so bump the year
                  vScheduledOn = CDate(vScheduledOn).ToString("dd/MM/" & Year(Today) + 1)
                ElseIf vDays > 0 Then
                  'Future so set the year
                  vScheduledOn = CDate(vScheduledOn).ToString("dd/MM/" & Year(Today))
                Else
                  'It's today - vDays = 0 so will get set below
                  vScheduledOn = CDate(vScheduledOn).ToString(CAREDateFormat)
                End If
                Return vScheduledOn & " " & vScheduledTime 'On date
              Else
                Return ""
              End If
            End If
          Else
            Return ""
          End If
        Else
          Return ""
        End If
      End Get
      Set(ByVal Value As String)
        If IsDate(Value) And ActionStatus = ActionStatuses.astProForma And PriorActions.Count() = 0 Then
          Dim vScheduledTime As String
          If IsDate(ScheduledOn) Then
            vScheduledTime = CDate(ScheduledOn).ToString("HH:mm:ss")
          Else
            vScheduledTime = CStr(mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))
          End If
          mvClassFields.Item(ActionFields.ScheduledOn).Value = CDate(Value).ToString(CAREDateFormat) & " " & vScheduledTime
        End If
      End Set
    End Property

    Public Sub InitNumberOnly(ByVal pEnv As CDBEnvironment, ByVal pMasterNumber As Integer, ByVal pActionNumber As Integer)
      mvEnv = pEnv
      Init()
      mvClassFields.Item(ActionFields.MasterAction).IntegerValue = pMasterNumber
      mvClassFields.Item(ActionFields.ActionNumber).IntegerValue = pActionNumber
    End Sub

    Friend Sub CopyActionLinks(ByVal pLinks As List(Of IActionLink))
      For Each vLink As IActionLink In pLinks
        AddLink((vLink.ObjectLinkType), vLink.LinkType, vLink.LinkedItemId)
      Next vLink
    End Sub

    Private Sub CopyActionSubjects(ByRef pSubjects As Collection, ByRef pActionNumber As Integer)
      Dim vNewSubject As ActionSubject

      For Each vActionSubject As ActionSubject In pSubjects
        vNewSubject = New ActionSubject(mvEnv)
        vNewSubject.Init()
        vNewSubject.CloneForNewAction(vActionSubject, pActionNumber)
        vNewSubject.Save(mvEnv.User.UserID, True)
      Next
    End Sub
    Public Sub EmailActioner()
      Dim vLink As ActionLink
      Dim vEmailJob As New EmailJob(mvEnv)
      Dim vEmailBody As String
      Dim vRecordSet As CDBRecordSet

      vEmailBody = mvEnv.GetConfig("email_overdue_text", "")

      For Each vLink In Links
        If vLink.ObjectLinkType <> IActionLink.ActionLinkObjectTypes.alotDocument Then

          If vLink.LinkType = IActionLink.ActionLinkTypes.altActioner Then
            'get email address from actioner

            Dim vWhereFields As New CDBFields
            vWhereFields.Add("contact_number", vLink.ContactNumber)
            vWhereFields.Add("address_number", vLink.AddressNumber)
            vWhereFields.Add("device", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEmailDevice))

            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & mvEnv.Connection.DBSpecialCol("communications", "number") & " FROM communications WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
            While vRecordSet.Fetch

              vEmailJob.Init()
              vEmailJob.SendEmail("Overdue Action " + ActionDesc, vEmailBody, vRecordSet.Fields("number").Value, vLink.ContactNumber.ToString)
            End While

            vRecordSet.CloseRecordSet()

          End If

        End If

      Next vLink

    End Sub
    Public Sub Activate(ByVal pDependant As Boolean)
      'Activate a currently inactive action
      Dim vDays As Integer
      Dim vDeadlineDays As Integer
      Dim vJournalType As JournalTypes
      Dim vLink As ActionLink
      Dim vRepeatAction As Action

      mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astInactive) 'Current status
      'Check the number of days to set up the deadline for
      If Deadline.Length > 0 Then
        vDeadlineDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(CreatedOn), CDate(Deadline)))
        Deadline = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDeadlineDays, CDate(Today & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))))
      End If
      If pDependant And ScheduledOn.Length > 0 Then 'If this action was dependant then check for delay required
        vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(CreatedOn), CDate(ScheduledOn)))
        If vDays > 0 Then
          'Delay was required so set the new activate date and leave the status as inactive
          mvClassFields.Item(ActionFields.ScheduledOn).Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDays, CDate(Today & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))).ToString(CAREDateTimeFormat)
        Else
          mvClassFields.Item(ActionFields.ScheduledOn).Value = "" 'Activate the action with no delay
          mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astDefined)
        End If
      Else
        mvClassFields.Item(ActionFields.ScheduledOn).Value = "" 'Activate the action with no delay
        mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astDefined)
      End If
      mvClassFields.Item(ActionFields.DelayedActivation).Bool = False
      mvClassFields.Item(ActionFields.CompletedOn).Value = ""
      mvClassFields.Save(mvEnv, mvExisting, "automatic", False)

      If ActionStatus = ActionStatuses.astDefined Then
        'Set any actioner or manager to be notified
        If mvEnv.JournalActive Then
          'Add a journal for each actioner or manager (Contact or Organisation)
          For Each vLink In Links
            If vLink.ObjectLinkType <> IActionLink.ActionLinkObjectTypes.alotDocument Then
              Select Case vLink.LinkType
                Case IActionLink.ActionLinkTypes.altActioner
                  vJournalType = JournalTypes.jnlActionActioner
                Case IActionLink.ActionLinkTypes.altManager
                  vJournalType = JournalTypes.jnlActionManager
                Case IActionLink.ActionLinkTypes.altRelated
                  vJournalType = JournalTypes.jnlActionRelated
              End Select
              mvEnv.AddJournalRecord(vJournalType, JournalOperations.jnlActive, vLink.ContactNumber, vLink.AddressNumber, ActionNumber)
            End If
          Next vLink
        End If
        For Each vLink In Links
          If vLink.ObjectLinkType = IActionLink.ActionLinkObjectTypes.alotContact And ((vLink.LinkType = IActionLink.ActionLinkTypes.altActioner) Or (vLink.LinkType = IActionLink.ActionLinkTypes.altManager)) Then
            If vLink.Notified Then
              vLink.Notified = False
              vLink.Save()
            End If
          End If
        Next vLink
      End If
      'Handle if we are activating a repeating action
      If IsDate(CompletedOn) Then
        vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(CreatedOn), CDate(CompletedOn)))
        If vDays > 0 Then
          If RepeatCount > 1 Then
            vRepeatAction = New Action(mvEnv)
            vRepeatAction.Init()
            vRepeatAction.SetupRepeatAction(Me, vDays)
          End If
        End If
      End If
    End Sub

    Public Sub UpdateFromOutlook(ByVal pDesc As String, ByVal pText As String, ByVal pScheduledOn As String)
      mvClassFields.Item(ActionFields.ActionDesc).Value = pDesc
      mvClassFields.Item(ActionFields.ActionText).Value = pText
      mvClassFields.Item(ActionFields.ScheduledOn).Value = pScheduledOn
    End Sub

    Public Sub SetCompleted(ByRef pNewDate As String)
      Dim vRS As CDBRecordSet
      Dim vFound As Boolean
      Dim vWhereFields As New CDBFields
      Dim vStatusChanged As Boolean
      Dim vJournalType As JournalTypes

      If pNewDate.Length > 0 Then
        vStatusChanged = ActionStatus <> ActionStatuses.astCompleted
        mvClassFields.Item(ActionFields.CompletedOn).Value = CDate(pNewDate).ToString(CAREDateTimeFormat)
        If mvClassFields.Item(ActionFields.ActionStatus).Value <> GetActionStatusCode(ActionStatuses.astProForma) Then mvClassFields.Item(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astCompleted)
        If mvEnv.JournalActive And vStatusChanged Then
          vRS = mvEnv.Connection.GetRecordSet("SELECT ca.contact_number, address_number, ca.type FROM contact_actions ca, contacts c WHERE ca.action_number = " & ActionNumber & " AND ca.contact_number = c.contact_number")
          While vRS.Fetch
            Select Case vRS.Fields(3).Value
              Case "A"
                vJournalType = JournalTypes.jnlActionActioner
              Case "M"
                vJournalType = JournalTypes.jnlActionManager
              Case "R"
                vJournalType = JournalTypes.jnlActionRelated
            End Select
            mvEnv.AddJournalRecord(vJournalType, JournalOperations.jnlComplete, vRS.Fields(1).LongValue, vRS.Fields(2).LongValue, ActionNumber)
          End While
          vRS.CloseRecordSet()
          'Add journal entry for all actioners and managers (Organisations)
          vRS = mvEnv.Connection.GetRecordSet("SELECT oa.organisation_number, address_number, oa.type FROM organisation_actions oa, contacts c WHERE oa.action_number = " & ActionNumber & " AND oa.organisation_number = c.contact_number")
          While vRS.Fetch
            Select Case vRS.Fields(3).Value
              Case "A"
                vJournalType = JournalTypes.jnlActionActioner
              Case "M"
                vJournalType = JournalTypes.jnlActionManager
              Case "R"
                vJournalType = JournalTypes.jnlActionRelated
            End Select
            mvEnv.AddJournalRecord(vJournalType, JournalOperations.jnlComplete, vRS.Fields(1).LongValue, vRS.Fields(2).LongValue, ActionNumber)
          End While
          vRS.CloseRecordSet()
        End If
        If vStatusChanged Then
          'Find any inactive actions that are dependent on this action
          Dim vAction As New Action(mvEnv)
          vAction.Init()
          vRS = mvEnv.Connection.GetRecordSet("SELECT " & vAction.GetRecordSetFields() & " FROM action_dependencies ad, actions ac WHERE prior_action = " & ActionNumber & " AND ad.action_number = ac.action_number AND action_status = '" & GetActionStatusCode(ActionStatuses.astInactive) & "'")
          While vRS.Fetch()
            vAction.InitFromRecordSet(vRS)
            vAction.Activate(True)
            vFound = True
          End While
          vRS.CloseRecordSet()
          'Remove any dependencies on this action as it is now complete
          'All actions dependant on this one are now defined or set as activate on date
          If vFound Then
            vWhereFields.Add("prior_action", CDBField.FieldTypes.cftLong, ActionNumber)
            mvEnv.Connection.DeleteRecords("action_dependencies", vWhereFields)
          End If
        End If
      Else
        mvClassFields.Item(ActionFields.CompletedOn).Value = ""
      End If
    End Sub

    Private Sub SetDeadline(ByVal pDeadlineDate As String)
      Dim vOverdue As Boolean

      mvClassFields(ActionFields.Deadline).Value = pDeadlineDate
      If IsDate(pDeadlineDate) Then
        If CDate(pDeadlineDate) < Today Then vOverdue = True
      End If
      If vOverdue Then
        If ActionStatus = ActionStatuses.astDefined Or ActionStatus = ActionStatuses.astScheduled Then mvClassFields(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astOverdue)
      Else
        If ActionStatus = ActionStatuses.astOverdue Then
          If IsDate(ScheduledOn) Then
            mvClassFields(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astScheduled)
          Else
            mvClassFields(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astDefined)
          End If
        End If
      End If
    End Sub

    Public Sub ResetAppointments(ByRef pCheck As Boolean)
      Dim vActionLink As ActionLink
      Dim vAppointment As New ContactAppointment(mvEnv)

      If pCheck And ScheduledOn.Length > 0 Then
        For Each vActionLink In Links
          CheckAppointment(vActionLink)
        Next vActionLink
      End If
      vAppointment.Init()
      vAppointment.ClearEntries(ContactAppointment.ContactAppointmentTypes.catAction, ActionNumber)
      If ScheduledOn.Length > 0 Then
        For Each vActionLink In From l As IActionLink In Links
                                Where l.GetType() Is GetType(ActionLink)
                                Select DirectCast(l, ActionLink)
          AddAppointment(vActionLink)
        Next vActionLink
      End If
    End Sub

    Public Sub CheckAppointment(ByRef pActionLink As ActionLink)
      Dim vAllowDoubleBooking As Boolean = mvEnv.GetConfigOption("ca_overlap_apt", False)
      Dim vAppointment As New ContactAppointment(mvEnv)
      If ScheduledOn.Length > 0 And pActionLink.LinkType = IActionLink.ActionLinkTypes.altActioner Then
        vAppointment.Init()
        If DurationDays = 0 And DurationHours = 0 And DurationMinutes = 0 Then SetDuration(0, 0, 30)
        If Not vAllowDoubleBooking Then
          'BR21013 If Calendar items are allowed to overlap, there is not need to check for conflicts
          vAppointment.CheckCalendarConflict(pActionLink.ContactNumber, ScheduledOn, AddDuration(CDate(ScheduledOn)).ToString(CAREDateTimeFormat), ContactAppointment.ContactAppointmentTypes.catAction, ActionNumber, False)
        End If
      End If
    End Sub

    Public Sub AddAppointment(ByRef pActionLink As ActionLink)
      Dim vAppointment As New ContactAppointment(mvEnv)
      If ScheduledOn.Length > 0 Then
        If pActionLink.LinkType = IActionLink.ActionLinkTypes.altActioner Then
          vAppointment.Init()
          vAppointment.Create(DirectCast(pActionLink, ActionLink).ContactNumber, ScheduledOn, AddDuration(CDate(ScheduledOn)).ToString(CAREDateTimeFormat), ContactAppointment.ContactAppointmentTypes.catAction, ActionDesc, ActionNumber)
          vAppointment.Save()
        ElseIf pActionLink.LinkType = IActionLink.ActionLinkTypes.altRelated Then
          If mvEnv.GetConfigOption("ac_related_appointments", False) Then
            vAppointment.Init()
            vAppointment.Create(DirectCast(pActionLink, ActionLink).ContactNumber, ScheduledOn, AddDuration(CDate(ScheduledOn)).ToString(CAREDateTimeFormat), ContactAppointment.ContactAppointmentTypes.catAction, ActionDesc, ActionNumber, ContactAppointment.ContactAppointmentTimeStatuses.catsFree)
            vAppointment.Save()
          End If
        End If
      End If
    End Sub

    Friend Sub SetupRepeatAction(ByRef pAction As Action, ByVal pDays As Integer)
      mvClassFields(ActionFields.ActionNumber).IntegerValue = mvEnv.GetControlNumber("AC")
      mvClassFields(ActionFields.CreatedBy).Value = pAction.CreatedBy
      mvClassFields(ActionFields.MasterAction).IntegerValue = pAction.MasterAction
      mvClassFields(ActionFields.ActionLevel).IntegerValue = pAction.ActionLevel
      mvClassFields(ActionFields.SequenceNumber).IntegerValue = pAction.SequenceNumber
      mvClassFields(ActionFields.ActionDesc).Value = pAction.ActionDesc
      mvClassFields(ActionFields.ActionText).Value = pAction.ActionText
      mvClassFields(ActionFields.ActionPriority).Value = pAction.ActionPriority
      mvClassFields(ActionFields.ActionStatus).Value = GetActionStatusCode(ActionStatuses.astInactive)
      mvClassFields(ActionFields.DocumentClass).Value = pAction.DocumentClass
      mvClassFields(ActionFields.DurationDays).IntegerValue = pAction.DurationDays
      mvClassFields(ActionFields.DurationHours).IntegerValue = pAction.DurationHours
      mvClassFields(ActionFields.DurationMinutes).IntegerValue = pAction.DurationMinutes
      mvClassFields(ActionFields.DelayedActivation).Bool = False
      mvClassFields(ActionFields.CompletedOn).Value = ""
      mvClassFields(ActionFields.RepeatCount).Value = ""
      If pAction.RepeatCount <> 1 Then 'More repeats
        mvClassFields(ActionFields.CompletedOn).Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, pDays, CDate(Today & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))).ToString(CAREDateTimeFormat)
        mvClassFields(ActionFields.RepeatCount).IntegerValue = pAction.RepeatCount - 1
      End If
      'Scheduled should be today + completed_on - created_on (the repeat delay)
      mvClassFields(ActionFields.ScheduledOn).Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, pDays, CDate(Today & " " & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))).ToString(CAREDateTimeFormat)
      mvClassFields(ActionFields.Deadline).Value = pAction.Deadline
      ActionerSetting = pAction.ActionerSetting
      ManagerSetting = pAction.ManagerSetting
      mvClassFields.Save(mvEnv, mvExisting, "automatic", False)
      CopyActionSubjects((pAction.Subjects), ActionNumber)
      CopyActionLinks((pAction.Links))
    End Sub

    Private Sub ResetLevels(ByVal pNewActionLevel As Integer)
      'Level of current Action has changed; reset level of any Actions under this Action
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vOrderClause As String
      Dim vWhereClause As String
      Dim vContinue As Boolean

      vWhereFields.Add("master_action", MasterAction, CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("action_number", ActionNumber, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("sequence_number", SequenceNumber, CDBField.FieldWhereOperators.fwoGreaterThan)
      vWhereClause = mvEnv.Connection.WhereClause(vWhereFields)
      vOrderClause = " ORDER BY sequence_number"
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT action_number,action_level FROM actions WHERE " & vWhereClause & vOrderClause)
      vWhereFields.Clear()
      vWhereFields.Add("action_number", CDBField.FieldTypes.cftLong)
      vWhereFields.Add("master_action", CDBField.FieldTypes.cftLong, mvClassFields(ActionFields.MasterAction).LongValue)
      vUpdateFields.Add("action_level", CDBField.FieldTypes.cftLong)
      vContinue = True
      While vRecordSet.Fetch And vContinue = True
        If vRecordSet.Fields("action_level").LongValue <= mvClassFields(ActionFields.ActionLevel).LongValue Then
          vContinue = False
        Else
          vWhereFields(1).Value = vRecordSet.Fields(1).Value
          If pNewActionLevel > mvClassFields(ActionFields.ActionLevel).LongValue Then
            vUpdateFields(1).Value = CStr(vRecordSet.Fields(2).LongValue + 1)
          Else
            vUpdateFields(1).Value = CStr(vRecordSet.Fields(2).LongValue - 1)
          End If
          mvEnv.Connection.UpdateRecords("actions", vUpdateFields, vWhereFields)
        End If
      End While
      vRecordSet.CloseRecordSet()
      mvClassFields(ActionFields.ActionLevel).Value = CStr(pNewActionLevel)
    End Sub

    Private Sub ResetSequences(ByVal pNewSequenceNumber As Integer)
      'Sequence of current Action has changed; reset sequence for all other Actions under this Master Action
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vOrderClause As String = ""
      Dim vWhereClause As String
      Dim vFirstRecord As Boolean

      vWhereFields.Add("master_action", MasterAction, CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add("action_number", ActionNumber, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereClause = mvEnv.Connection.WhereClause(vWhereFields)

      If mvClassFields(ActionFields.SequenceNumber).LongValue = 0 Then
        'New Action; does not yet have a sequence
        vWhereClause = vWhereClause & " AND sequence_number >= " & pNewSequenceNumber
        vOrderClause = " ORDER BY sequence_number DESC"
      ElseIf pNewSequenceNumber < mvClassFields(ActionFields.SequenceNumber).LongValue Then
        'Current Action is moving up list
        vWhereClause = vWhereClause & " AND sequence_number >= " & pNewSequenceNumber
        vWhereClause = vWhereClause & " AND sequence_number < " & mvClassFields(ActionFields.SequenceNumber).LongValue
        vOrderClause = " ORDER BY sequence_number DESC"
      ElseIf pNewSequenceNumber > mvClassFields(ActionFields.SequenceNumber).LongValue Then
        'Current Action is moving down list
        vWhereClause = vWhereClause & " AND sequence_number > " & mvClassFields(ActionFields.SequenceNumber).LongValue
        vWhereClause = vWhereClause & " AND sequence_number <= " & pNewSequenceNumber
        vOrderClause = " ORDER BY sequence_number"
      End If
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT action_number,sequence_number,action_level FROM actions WHERE " & vWhereClause & vOrderClause)
      vWhereFields.Clear()
      vWhereFields.Add("action_number", CDBField.FieldTypes.cftLong)
      vWhereFields.Add("master_action", CDBField.FieldTypes.cftLong, mvClassFields(ActionFields.MasterAction).LongValue)
      vUpdateFields.Add("sequence_number", CDBField.FieldTypes.cftLong)
      vUpdateFields.Add("action_level", CDBField.FieldTypes.cftLong)
      vFirstRecord = True
      While vRecordSet.Fetch
        vWhereFields(1).Value = CStr(vRecordSet.Fields(1).LongValue)
        If mvClassFields(ActionFields.SequenceNumber).LongValue = 0 Or pNewSequenceNumber < mvClassFields(ActionFields.SequenceNumber).LongValue Then
          vUpdateFields(1).Value = CStr(vRecordSet.Fields(2).LongValue + 10)
        Else
          vUpdateFields(1).Value = CStr(vRecordSet.Fields(2).LongValue - 10)
        End If
        If mvExisting = True And vFirstRecord = True Then
          'Swap level
          vUpdateFields(2).Value = mvClassFields(ActionFields.ActionLevel).Value
          mvClassFields(ActionFields.ActionLevel).Value = CStr(vRecordSet.Fields(3).LongValue)
          vFirstRecord = False
        Else
          vUpdateFields(2).Value = CStr(vRecordSet.Fields(3).LongValue)
        End If
        mvEnv.Connection.UpdateRecords("actions", vUpdateFields, vWhereFields)
      End While
      vRecordSet.CloseRecordSet()
      mvClassFields(ActionFields.SequenceNumber).Value = CStr(pNewSequenceNumber)
    End Sub

    Public Function GetPossibleScheduleDate(ByVal pType As ActionsScheduleTypes, ByVal pIgnoreWorkingDay As Boolean, ByVal pIgnoreWeekend As Boolean, Optional ByVal pDate As Date = #12:00:00 AM#, Optional ByVal pEarliestDate As String = "", Optional ByVal pContactNumber As Integer = 0) As String
      Dim vActionerNumbers As String = ""
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vRecordSet As CDBRecordSet
      Dim vPossibleStart As Date
      Dim vPossibleEnd As Date
      Dim vMaxPossibleEnd As Date
      Dim vMins As Integer
      Dim vDay As Integer
      Dim vDate As Date
      Dim vCannotSchedule As Boolean
      Dim vWhereFields As New CDBFields
      Dim vMaxDay As Integer 'Day of Week
      Dim vNoDays As Integer 'Number of Days to increment the date
      Dim vColl As New Collection
      Dim vEarliestDate As Date

      'Look for the first available schedule time for the action
      mvIgnoreWorkingDay = pIgnoreWorkingDay
      mvIgnoreWeekend = pIgnoreWeekend
      If Len(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfLunch)) = 0 Then mvIgnoreLunch = True

      vMins = Minute(Now)
      vMins = ((Int(vMins \ 15) + 1) * 15) - vMins '15 minute slots

      Select Case pType
        Case ActionsScheduleTypes.astGivenDate
          vPossibleStart = pDate
          vMaxPossibleEnd = AddDuration(vPossibleStart)
        Case ActionsScheduleTypes.astSoonest
          vPossibleStart = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, vMins, Now), SetTimeTypes.sttNoSeconds)
          vMaxPossibleEnd = SetTimeOnDate(vPossibleStart.AddYears(1), SetTimeTypes.sttEndOfDay)
        Case ActionsScheduleTypes.astToday
          vPossibleStart = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, vMins, Now), SetTimeTypes.sttNoSeconds)
          vMaxPossibleEnd = SetTimeOnDate(vPossibleStart, SetTimeTypes.sttEndOfDay)
        Case ActionsScheduleTypes.astTomorrow
          vPossibleStart = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, Today), SetTimeTypes.sttStartofDay)
          vMaxPossibleEnd = SetTimeOnDate(vPossibleStart, SetTimeTypes.sttEndOfDay)
        Case ActionsScheduleTypes.astThisWeek
          vPossibleStart = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, vMins, Now), SetTimeTypes.sttNoSeconds)
          vDay = Weekday(Now)
          vMaxDay = If(mvIgnoreWeekend = True, FirstDayOfWeek.Saturday, FirstDayOfWeek.Friday) 'vbSunday = 1 .... vbSaturday = 7
          If vDay <> vMaxDay Then
            'Sets vDay to the number of days to vMaxDay
            vDay = vMaxDay - vDay 'Make sure max is Saturday / Friday
          Else
            vDay = 0
          End If
          vMaxPossibleEnd = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDay, vPossibleStart), SetTimeTypes.sttEndOfDay)
        Case ActionsScheduleTypes.astNextWeek
          vDay = Weekday(Now)
          If mvIgnoreWeekend Then
            vMaxDay = FirstDayOfWeek.Saturday
            vNoDays = 8
          Else
            vMaxDay = FirstDayOfWeek.Sunday
            vNoDays = 9
          End If
          'Set vDay to the number of days to the day after vMaxDay
          If vDay <> vMaxDay Then vDay = vNoDays - vDay 'Make sure day is next Sunday / Monday
          vPossibleStart = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDay, Today), SetTimeTypes.sttStartofDay)
          vMaxPossibleEnd = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 4, vPossibleStart), SetTimeTypes.sttEndOfDay) 'make sure max is friday
        Case ActionsScheduleTypes.astNextMonth
          vDate = DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, Today)
          vPossibleStart = SetTimeOnDate(DateSerial(Year(vDate), Month(vDate), 1), SetTimeTypes.sttStartofDay)
          AdjustForWeekend(vPossibleStart)
          'To find end add another month - set to first of the month and subtract one day
          vDate = DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, vPossibleStart)
          vDate = DateSerial(Year(vDate), Month(vDate), 1)
          vMaxPossibleEnd = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, vDate), SetTimeTypes.sttEndOfDay)
      End Select

      If Len(pEarliestDate) > 0 Then
        vEarliestDate = CDate(pEarliestDate)
        If vEarliestDate > vPossibleStart Then vPossibleStart = vEarliestDate
      End If

      'Now test to check that there is enough time between start and end dates
      If AddDuration(vPossibleStart) > vMaxPossibleEnd Then
        vCannotSchedule = True
      Else
        'First get a list of all the actioners
        Dim vLink As ActionLink
        For Each vLink In Links
          If vLink.LinkType = IActionLink.ActionLinkTypes.altActioner Then
            If vActionerNumbers.Length > 0 Then vActionerNumbers = vActionerNumbers & ","
            vActionerNumbers = vActionerNumbers & vLink.ContactNumber
          End If
        Next vLink

        'Now get a list of all their appointments during the possible period
        vPossibleEnd = AddDuration(vPossibleStart)
        If vActionerNumbers.Length > 0 AndAlso mvEnv.GetConfigOption("ca_overlap_apt", True) = False Then
          vAppointment.Init()
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, vActionerNumbers, CDBField.FieldWhereOperators.fwoIn)
          vWhereFields.Add("start_date", CDBField.FieldTypes.cftTime, vMaxPossibleEnd.ToString(CAREDateTimeFormat), CDBField.FieldWhereOperators.fwoLessThan)
          vWhereFields.Add("end_date", CDBField.FieldTypes.cftTime, vPossibleStart.ToString(CAREDateTimeFormat), CDBField.FieldWhereOperators.fwoGreaterThan)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAppointment.GetRecordSetFields() & " FROM contact_appointments ca WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY start_date")
          While vRecordSet.Fetch()
            vAppointment.InitFromRecordSet(vRecordSet)
            vColl.Add(vAppointment)
            vAppointment = New ContactAppointment(mvEnv)
            vAppointment.Init()
          End While
          vRecordSet.CloseRecordSet()
        End If
        'Get external calendar appointments
        For Each vAppointment In vColl
          With vAppointment
            If .AppointmentType <> ContactAppointment.ContactAppointmentTypes.catAction Or ((.AppointmentType = ContactAppointment.ContactAppointmentTypes.catAction) And (.UniqueId <> ActionNumber)) Then
              If CDate(.StartDate) <= vPossibleStart And CDate(.EndDate) > vPossibleStart Then
                vPossibleStart = CDate(.EndDate)
                AdjustForWorkDay(vPossibleStart)
                vPossibleEnd = AddDuration(vPossibleStart)
              ElseIf CDate(.StartDate) < vPossibleEnd Then
                vPossibleStart = CDate(.EndDate)
                AdjustForWorkDay(vPossibleStart)
                vPossibleEnd = AddDuration(vPossibleStart)
              End If
            End If
          End With
        Next vAppointment
        If vPossibleEnd > vMaxPossibleEnd Then vCannotSchedule = True
        End If
        If Not vCannotSchedule Then
          Return vPossibleStart.ToString(CAREDateTimeFormat)
        Else
          Return ""
        End If
    End Function

    Private Function SetTimeOnDate(ByRef pDate As Date, ByRef pType As SetTimeTypes) As Date
      Dim vTime As Date
      Dim vDate As Date
      Dim vWorkTime As String

      vDate = DateValue(CStr(pDate))
      vTime = TimeValue(CStr(pDate))

      Select Case pType
        Case SetTimeTypes.sttNoSeconds
          SetTimeOnDate = CDate(vDate & " " & vTime.ToString("HH:mm"))
        Case SetTimeTypes.sttStartofDay
          vWorkTime = If(mvIgnoreWorkingDay = True, "00:00", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay))
          SetTimeOnDate = CDate(vDate & " " & vWorkTime)
        Case SetTimeTypes.sttEndOfDay
          vWorkTime = If(mvIgnoreWorkingDay = True, "23:59", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEndOfDay))
          SetTimeOnDate = CDate(vDate & " " & vWorkTime)
      End Select
    End Function

    Private Sub AdjustForWorkDay(ByRef pDate As Date)
      Dim vCurrentTime As Date

      If mvIgnoreWorkingDay = False Then
        vCurrentTime = TimeValue(CStr(pDate))
        If vCurrentTime < TimeValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay)) Then
          pDate = SetTimeOnDate(pDate, SetTimeTypes.sttStartofDay)
        ElseIf vCurrentTime >= TimeValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEndOfDay)) Then
          pDate = SetTimeOnDate(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, pDate), SetTimeTypes.sttStartofDay)
        End If
      End If
      AdjustForWeekend(pDate)
    End Sub

  End Class
End Namespace
