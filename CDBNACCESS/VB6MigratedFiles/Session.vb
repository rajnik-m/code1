

Namespace Access
  Public Class EventSession

    Public Enum SessionRecordSetTypes 'These are bit values
      ssrtAll = &HFFFFS
      'ADD additional recordset types here
      ssrtNumber = 1
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SessionFields
      sfAll = 0
      sfEventNumber
      sfSessionNumber
      sfSessionDesc
      sfSessionType
      sfSubject
      sfSkillLevel
      sfStartDate
      sfEndDate
      sfStartTime
      sfEndTime
      sfLocation
      sfMinimumAttendees
      sfMaximumAttendees
      sfTargetAttendees
      sfNumberInterested
      sfNumberOfAttendees
      sfNumberOnWaitingList
      sfMaximumOnWaitingList
      sfNotes
      sfVenueBookingNumber
      sfCpdApprovalStatus
      sfCpdDateApproved
      sfCpdAwardingBody
      sfCpdCategory
      sfCpdYear
      sfCpdPoints
      sfCpdNotes
      sfExternalAppointmentID
      sfAmendedBy
      sfAmendedOn
      sfSessionLongDesc
    End Enum

    Public Enum CPDApprovalStatusTypes
      castNone
      castApproved
      castWaitingApproval
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvSelectedPersonnel As Contacts
    Private mvNewStartDate As String
    Private mvNewEndDate As String
    Private mvVenueBooking As EventVenueBooking
    Private mvAppointmentDesc As String
    Private mvActivities As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "sessions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftInteger)
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("session_desc")
          .Add("session_type")
          .Add("subject")
          .Add("skill_level")
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("end_date", CDBField.FieldTypes.cftDate)
          .Add("start_time")
          .Add("end_time")
          .Add("location")
          .Add("minimum_attendees", CDBField.FieldTypes.cftInteger)
          .Add("maximum_attendees", CDBField.FieldTypes.cftInteger)
          .Add("target_attendees", CDBField.FieldTypes.cftInteger)
          .Add("number_interested", CDBField.FieldTypes.cftInteger)
          .Add("number_of_attendees", CDBField.FieldTypes.cftInteger)
          .Add("number_on_waiting_list", CDBField.FieldTypes.cftInteger)
          .Add("maximum_on_waiting_list", CDBField.FieldTypes.cftInteger)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("venue_booking_number", CDBField.FieldTypes.cftLong)
          .Add("cpd_approval_status")
          .Add("cpd_date_approved", CDBField.FieldTypes.cftDate)
          .Add("cpd_awarding_body")
          .Add("cpd_category")
          .Add("cpd_year", CDBField.FieldTypes.cftInteger)
          .Add("cpd_points", CDBField.FieldTypes.cftLong)
          .Add("cpd_notes")
          .Add("external_appointment_id")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("long_description")
        End With
        'Although cpd_notes and location are actually memos we don't set them like that as they will fail in oracle
        'since in oracle there is only one long field and that is the notes field

        mvClassFields.Item(SessionFields.sfSessionNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(SessionFields.sfEventNumber).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfSessionNumber).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfStartDate).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfAmendedBy).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfAmendedOn).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfCpdCategory).PrefixRequired = True
        mvClassFields.Item(SessionFields.sfSessionLongDesc).PrefixRequired = True

        mvClassFields.Item(SessionFields.sfSessionLongDesc).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataEventLongDescription)
        mvClassFields.Item(SessionFields.sfExternalAppointmentID).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbOutlookIntegration)
        'UPGRADE_NOTE: Object mvVenueBooking may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mvVenueBooking = Nothing
      Else
        mvClassFields.ClearItems()
      End If
      'UPGRADE_NOTE: Object mvActivities may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvActivities = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(SessionFields.sfNumberOfAttendees).Value = CStr(0)
      mvClassFields.Item(SessionFields.sfNumberOnWaitingList).Value = CStr(0)
      mvClassFields.Item(SessionFields.sfNumberInterested).Value = CStr(0)
      mvClassFields.Item(SessionFields.sfStartDate).Value = TodaysDate()
      mvClassFields.Item(SessionFields.sfEndDate).Value = TodaysDate()
      mvClassFields.Item(SessionFields.sfStartTime).Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStartOfDay)
      mvClassFields.Item(SessionFields.sfEndTime).Value = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEndOfDay)
    End Sub

    Private Sub SetValid(ByRef pField As SessionFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SessionFields.sfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SessionFields.sfAmendedBy).Value = mvEnv.User.Logname

      If mvClassFields.Item(SessionFields.sfVenueBookingNumber).InDatabase And mvClassFields.Item(SessionFields.sfVenueBookingNumber).IntegerValue = 0 And Not (mvVenueBooking Is Nothing) Then
        mvClassFields.Item(SessionFields.sfVenueBookingNumber).Value = CStr(mvVenueBooking.VenueBookingNumber)
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SessionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SessionRecordSetTypes.ssrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "s")
      Else
        vFields = "session_number"
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSessionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSessionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SessionRecordSetTypes.ssrtAll) & " FROM sessions s WHERE session_number = " & pSessionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SessionRecordSetTypes.ssrtAll)
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

    Friend Sub InitFromSession(ByRef pNewEvent As CDBEvent, ByRef pBaseDate As String, ByRef pNewBaseDate As String, ByRef pSession As EventSession, ByRef pWeekdaysOnly As Boolean)
      Dim vDays As Integer
      Dim vNewStartDate As String
      Dim vNewEndDate As String
      Dim vNoWeekdays As Integer

      With pSession
        mvClassFields.Item(SessionFields.sfSessionNumber).Value = CStr(pNewEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfSessionNumber))   'CStr(.SessionNumber)    'CStr(pNewEvent.BaseItemNumber + (.SessionNumber Mod 10000))
        mvClassFields.Item(SessionFields.sfEventNumber).Value = CStr(pNewEvent.EventNumber)
        'If CDbl(mvClassFields.Item(SessionFields.sfSessionNumber).Value) = pNewEvent.LowestSessionNumber Then '   pNewEvent.BaseItemNumber Then
        If pSession.SessionType = "0" Then
          mvClassFields.Item(SessionFields.sfSessionDesc).Value = String.Format(ProjectText.String26125, Left(pNewEvent.EventDesc, 42))
        Else
          mvClassFields.Item(SessionFields.sfSessionDesc).Value = .SessionDesc
        End If
        mvClassFields.Item(SessionFields.sfSessionLongDesc).Value = .SessionLongDesc
        mvClassFields.Item(SessionFields.sfSessionType).Value = .SessionType
        mvClassFields.Item(SessionFields.sfSubject).Value = .Subject
        mvClassFields.Item(SessionFields.sfSkillLevel).Value = .SkillLevel
        mvClassFields.Item(SessionFields.sfLocation).Value = .Location
        mvClassFields.Item(SessionFields.sfMinimumAttendees).Value = CStr(.MinimumAttendees)
        mvClassFields.Item(SessionFields.sfMaximumAttendees).Value = CStr(.MaximumAttendees)
        mvClassFields.Item(SessionFields.sfMaximumOnWaitingList).Value = CStr(.MaximumOnWaitingList)
        mvClassFields.Item(SessionFields.sfTargetAttendees).Value = CStr(.TargetAttendees)
        mvClassFields.Item(SessionFields.sfNotes).Value = .Notes
        mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value = GetApprovalStatusCode(.CpdApprovalStatus)
        If mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value = "N" Then mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value = ""
        mvClassFields.Item(SessionFields.sfCpdDateApproved).Value = .CPDDateApproved
        mvClassFields.Item(SessionFields.sfCpdAwardingBody).Value = .CpdAwardingBody
        mvClassFields.Item(SessionFields.sfCpdCategory).Value = .CPDCategory
        mvClassFields.Item(SessionFields.sfCpdYear).Value = .CPDYear
        mvClassFields.Item(SessionFields.sfCpdPoints).Value = .CPDPoints
        mvClassFields.Item(SessionFields.sfCpdNotes).Value = .CPDNotes

        vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pBaseDate), CDate(.StartDate))) 'Find no of days from base date to session start
        vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDays, CDate(pNewBaseDate))) 'Add to new base date

        If pWeekdaysOnly Then vNewStartDate = CStr(NextWeekDay(CDate(vNewStartDate)))
        mvClassFields.Item(SessionFields.sfStartDate).Value = vNewStartDate
        mvClassFields.Item(SessionFields.sfStartTime).Value = .StartTime

        If pWeekdaysOnly Then
          vNoWeekdays = WeekDaysDiff(CDate(.StartDate), CDate(.EndDate))  'Find no of weekdays from start to end of session
          vNewEndDate = CStr(AddWeekdays(CDate(vNewStartDate), vNoWeekdays)) 'Add to new start to set the end date
        Else
          vDays = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(.StartDate), CDate(.EndDate))) 'Find no of days from start to end of session
          vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, vDays, CDate(vNewStartDate))) 'Add to new start to set the end date
        End If
        mvClassFields.Item(SessionFields.sfEndDate).Value = vNewEndDate
        mvClassFields.Item(SessionFields.sfEndTime).Value = .EndTime
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SessionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SessionFields.sfSessionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SessionRecordSetTypes.ssrtAll) = SessionRecordSetTypes.ssrtAll Then
          .SetItem(SessionFields.sfEventNumber, vFields)
          .SetItem(SessionFields.sfSessionDesc, vFields)
          .SetItem(SessionFields.sfSessionType, vFields)
          .SetItem(SessionFields.sfSubject, vFields)
          .SetItem(SessionFields.sfSkillLevel, vFields)
          .SetItem(SessionFields.sfStartDate, vFields)
          .SetItem(SessionFields.sfEndDate, vFields)
          .SetItem(SessionFields.sfStartTime, vFields)
          .SetItem(SessionFields.sfEndTime, vFields)
          .SetItem(SessionFields.sfLocation, vFields)
          .SetItem(SessionFields.sfMinimumAttendees, vFields)
          .SetItem(SessionFields.sfMaximumAttendees, vFields)
          .SetItem(SessionFields.sfTargetAttendees, vFields)
          .SetItem(SessionFields.sfNumberInterested, vFields)
          .SetItem(SessionFields.sfNumberOfAttendees, vFields)
          .SetItem(SessionFields.sfNumberOnWaitingList, vFields)
          .SetItem(SessionFields.sfMaximumOnWaitingList, vFields)
          .SetOptionalItem(SessionFields.sfVenueBookingNumber, vFields)
          .SetOptionalItem(SessionFields.sfCpdApprovalStatus, vFields)
          .SetOptionalItem(SessionFields.sfCpdDateApproved, vFields)
          .SetOptionalItem(SessionFields.sfCpdAwardingBody, vFields)
          .SetOptionalItem(SessionFields.sfCpdCategory, vFields)
          .SetOptionalItem(SessionFields.sfCpdYear, vFields)
          .SetOptionalItem(SessionFields.sfCpdPoints, vFields)
          .SetOptionalItem(SessionFields.sfCpdNotes, vFields)
          .SetOptionalItem(SessionFields.sfExternalAppointmentID, vFields)
          .SetItem(SessionFields.sfNotes, vFields)
          .SetItem(SessionFields.sfAmendedBy, vFields)
          .SetItem(SessionFields.sfAmendedOn, vFields)
          .SetOptionalItem(SessionFields.sfSessionLongDesc, vFields)
        End If
      End With
    End Sub

    Public Function CanDelete(ByRef pMessage As String) As Boolean
      Dim vCanDelete As Boolean = False

      Dim vWhereFields As New CDBFields
      vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong, SessionNumber)
      If mvEnv.Connection.GetCount("session_bookings", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27201) 'Bookings have been made for this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("option_sessions", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27202) 'Booking Options have been setup for this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("event_resources", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27203) 'Event Resources have been setup for this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("event_personnel", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27218) 'Event Personnel records refer to this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("session_activities", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27219) 'Session Activities records refer to this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("session_tests", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27229) 'Session Tests records refer to this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("session_test_results", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27230) 'Session Test Results records refer to this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("session_candidate_numbers", vWhereFields) > 0 Then
        pMessage = (ProjectText.String27231) 'Session Candidate Numbers records refer to this Session; Session cannot be deleted
      ElseIf mvEnv.Connection.GetCount("session_cpd", vWhereFields) > 0 Then
        pMessage = ProjectText.CannotDeleteSessionCPDExists   'Session CPD records refer to this Session; Session cannot be deleted
      Else
        vCanDelete = True
      End If

      Return vCanDelete
    End Function

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong, SessionNumber)
      mvEnv.Connection.StartTransaction()
      mvEnv.Connection.DeleteRecords("event_personnel", vWhereFields, False)
      mvEnv.Connection.DeleteRecords("event_resources", vWhereFields, False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vInTransaction As Boolean
      Dim vWhereFields As New CDBFields
      Dim vEventPersonnel As New EventPersonnel(mvEnv)
      Dim vRecordSet As CDBRecordSet

      SetValid(SessionFields.sfAll)
      vInTransaction = mvEnv.Connection.InTransaction
      If mvExisting Then
        'If we have changed the dates of an existing session we need to reschedule the personnel
        If mvClassFields.Item(SessionFields.sfStartDate).ValueChanged Or mvClassFields.Item(SessionFields.sfStartTime).ValueChanged Or mvClassFields.Item(SessionFields.sfEndDate).ValueChanged Or mvClassFields.Item(SessionFields.sfEndTime).ValueChanged Then
          vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong, SessionNumber)
          vEventPersonnel.Init()
          If Not vInTransaction Then mvEnv.Connection.StartTransaction()
          'Find all the personnel for the session
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vEventPersonnel.GetRecordSetFields() & " FROM event_personnel ep WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
          While vRecordSet.Fetch() = True
            vEventPersonnel.InitFromRecordSet(vRecordSet)
            vEventPersonnel.Reschedule(StartDate, StartTime, EndDate, EndTime)
            vEventPersonnel.AppointmentDescription = mvAppointmentDesc
            vEventPersonnel.Save()
          End While
          vRecordSet.CloseRecordSet()
        End If
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      If Not vInTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SetLocationFromVenue(ByRef pVenue As String)
      mvClassFields.Item(SessionFields.sfLocation).Value = mvEnv.Connection.GetValue("SELECT location FROM venues WHERE venue = '" & pVenue & "'")
    End Sub

    Public Sub SetValuesFromEvent(ByVal pEvent As CDBEvent)
      If Not mvExisting Then
        mvClassFields.Item(SessionFields.sfEventNumber).Value = CStr(pEvent.EventNumber)
        If mvClassFields.Item(SessionFields.sfSessionType).Value.Length > 0 Then
          mvClassFields.Item(SessionFields.sfSessionNumber).Value = CStr(pEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfSessionNumber))
        Else
          mvClassFields.Item(SessionFields.sfSessionNumber).Value = CStr(pEvent.LowestSessionNumber)   'CStr(pEvent.BaseItemNumber)
          mvClassFields.Item(SessionFields.sfSessionType).Value = "0"
          mvClassFields.Item(SessionFields.sfStartDate).Value = pEvent.StartDate
        End If
      Else
        If SessionNumber = pEvent.LowestSessionNumber Then
          mvClassFields.Item(SessionFields.sfStartDate).Value = pEvent.StartDate
        End If
      End If
      If SessionNumber = pEvent.LowestSessionNumber Then   'pEvent.BaseItemNumber Then
        mvAppointmentDesc = pEvent.EventDesc
      Else
        mvAppointmentDesc = pEvent.EventDesc & " : " & SessionDesc
      End If
    End Sub

    Public Sub AddStandardResources()
      Dim vAttrList As String
      Dim vSelectList As String

      vAttrList = "session_number, product, rate, copy_to, despatch_to, issue_basis, amended_on, amended_by, allocated"
      vSelectList = "session_number,sr.product, sr.rate, copy_to, despatch_to, issue_basis, " & mvEnv.Connection.SQLLiteral("", Now) & ", '" & mvEnv.User.Logname & "' , 'N' "
      mvEnv.Connection.ExecuteSQL("INSERT INTO event_resources (" & vAttrList & ") SELECT " & vSelectList & " FROM sessions s, standard_resources sr WHERE s.session_number = " & SessionNumber & " AND sr.subject = s.subject AND sr.skill_level = s.skill_level")
    End Sub

    Public ReadOnly Property Activities() As Collection
      Get
        Dim vSessionActivity As New EventSessionActivity
        Dim vRecordSet As CDBRecordSet

        If mvActivities Is Nothing Then
          mvActivities = New Collection
          vSessionActivity.Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vSessionActivity.GetRecordSetFields(EventSessionActivity.EventSessionActivityRecordSetTypes.esartAll) & " FROM session_activities WHERE session_number = " & SessionNumber & " ORDER BY activity")
          While vRecordSet.Fetch() = True
            vSessionActivity = New EventSessionActivity
            vSessionActivity.InitFromRecordSet(mvEnv, vRecordSet, EventSessionActivity.EventSessionActivityRecordSetTypes.esartAll)
            mvActivities.Add(vSessionActivity)
          End While
          vRecordSet.CloseRecordSet()
        End If
        Activities = mvActivities
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(SessionFields.sfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SessionFields.sfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CpdApprovalStatus() As CPDApprovalStatusTypes
      Get
        Select Case mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value
          Case "A"
            CpdApprovalStatus = CPDApprovalStatusTypes.castApproved
          Case "W"
            CpdApprovalStatus = CPDApprovalStatusTypes.castWaitingApproval
          Case Else
            CpdApprovalStatus = CPDApprovalStatusTypes.castNone
        End Select
      End Get
    End Property

    Public ReadOnly Property CpdApprovalStatusCode() As String
      Get
        Return mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value
      End Get
    End Property

    Public ReadOnly Property CpdAwardingBody() As String
      Get
        CpdAwardingBody = mvClassFields.Item(SessionFields.sfCpdAwardingBody).Value
      End Get
    End Property

    Public ReadOnly Property CPDCategory() As String
      Get
        CPDCategory = mvClassFields.Item(SessionFields.sfCpdCategory).Value
      End Get
    End Property

    Public ReadOnly Property CPDDateApproved() As String
      Get
        CPDDateApproved = mvClassFields.Item(SessionFields.sfCpdDateApproved).Value
      End Get
    End Property

    Public ReadOnly Property CPDNotes() As String
      Get
        CPDNotes = mvClassFields.Item(SessionFields.sfCpdNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property CPDPoints() As String
      Get
        CPDPoints = mvClassFields.Item(SessionFields.sfCpdPoints).Value
      End Get
    End Property

    Public ReadOnly Property CPDYear() As String
      Get
        CPDYear = mvClassFields.Item(SessionFields.sfCpdYear).Value
      End Get
    End Property

    Public ReadOnly Property EndDate() As String
      Get
        EndDate = mvClassFields.Item(SessionFields.sfEndDate).Value
      End Get
    End Property

    Public ReadOnly Property EndTime() As String
      Get
        EndTime = mvClassFields.Item(SessionFields.sfEndTime).Value
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(SessionFields.sfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Location() As String
      Get
        Location = mvClassFields.Item(SessionFields.sfLocation).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property MaximumAttendees() As Integer
      Get
        MaximumAttendees = mvClassFields.Item(SessionFields.sfMaximumAttendees).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MaximumOnWaitingList() As Integer
      Get
        MaximumOnWaitingList = mvClassFields.Item(SessionFields.sfMaximumOnWaitingList).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MinimumAttendees() As Integer
      Get
        MinimumAttendees = mvClassFields.Item(SessionFields.sfMinimumAttendees).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(SessionFields.sfNotes).MultiLineValue
      End Get
    End Property

    Public Property NumberInterested() As Integer
      Get
        NumberInterested = mvClassFields.Item(SessionFields.sfNumberInterested).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionFields.sfNumberInterested).Value = CStr(Value)
      End Set
    End Property

    Public Property NumberOfAttendees() As Integer
      Get
        NumberOfAttendees = mvClassFields.Item(SessionFields.sfNumberOfAttendees).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionFields.sfNumberOfAttendees).Value = CStr(Value)
      End Set
    End Property

    Public Property NumberOnWaitingList() As Integer
      Get
        NumberOnWaitingList = mvClassFields.Item(SessionFields.sfNumberOnWaitingList).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionFields.sfNumberOnWaitingList).Value = CStr(Value)
      End Set
    End Property

    Public ReadOnly Property SessionDesc() As String
      Get
        SessionDesc = mvClassFields.Item(SessionFields.sfSessionDesc).Value
      End Get
    End Property

    Public Property SessionLongDesc() As String
      Get
        SessionLongDesc = mvClassFields.Item(SessionFields.sfSessionLongDesc).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SessionFields.sfSessionLongDesc).Value = Value
      End Set
    End Property

    Public ReadOnly Property SessionNumber() As Integer
      Get
        SessionNumber = mvClassFields.Item(SessionFields.sfSessionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SessionType() As String
      Get
        SessionType = mvClassFields.Item(SessionFields.sfSessionType).Value
      End Get
    End Property

    Public ReadOnly Property SkillLevel() As String
      Get
        SkillLevel = mvClassFields.Item(SessionFields.sfSkillLevel).Value
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(SessionFields.sfStartDate).Value
      End Get
    End Property

    Public ReadOnly Property StartTime() As String
      Get
        StartTime = mvClassFields.Item(SessionFields.sfStartTime).Value
      End Get
    End Property

    Public ReadOnly Property Subject() As String
      Get
        Subject = mvClassFields.Item(SessionFields.sfSubject).Value
      End Get
    End Property

    Public ReadOnly Property TargetAttendees() As Integer
      Get
        TargetAttendees = mvClassFields.Item(SessionFields.sfTargetAttendees).IntegerValue
      End Get
    End Property

    Public Property VenueBookingNumber() As String
      Get
        VenueBookingNumber = mvClassFields.Item(SessionFields.sfVenueBookingNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(SessionFields.sfVenueBookingNumber).Value = Value
      End Set
    End Property

    Public ReadOnly Property SelectedPersonnelContacts() As Contacts
      Get
        If mvSelectedPersonnel Is Nothing Then mvSelectedPersonnel = New Contacts(mvEnv)
        SelectedPersonnelContacts = mvSelectedPersonnel
      End Get
    End Property

    Public Property NewStartDate() As String
      Get
        NewStartDate = mvNewStartDate
      End Get
      Set(ByVal Value As String)
        mvNewStartDate = Value
      End Set
    End Property
    Public Property NewEndDate() As String
      Get
        NewEndDate = mvNewEndDate
      End Get
      Set(ByVal Value As String)
        mvNewEndDate = Value
      End Set
    End Property

    Public ReadOnly Property VenueBooking() As EventVenueBooking
      Get
        If mvVenueBooking Is Nothing Then
          mvVenueBooking = New EventVenueBooking
          'Since we are using the venue booking class we must assume the
          'attribute is in the database
          mvClassFields.Item(SessionFields.sfVenueBookingNumber).InDatabase = True
        End If
        VenueBooking = mvVenueBooking
      End Get
    End Property

    Public ReadOnly Property BaseSessionType() As String
      Get
        BaseSessionType = "0"
      End Get
    End Property

    Public ReadOnly Property ExternalAppointmentId() As String
      Get
        Return mvClassFields(SessionFields.sfExternalAppointmentID).Value
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pEvent As CDBEvent, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(SessionFields.sfEventNumber).Value = pParams("EventNumber").Value
        'If pEvent.Existing Then
        .Item(SessionFields.sfSessionNumber).Value = CStr(pEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfSessionNumber))
        'Else
        '.Item(SessionFields.sfSessionNumber).Value = CStr(pEvent.BaseItemNumber)
        'End If
        Update(pEvent, pParams)
      End With
    End Sub

    Public Function PersonnelAvailable(ByVal pNewStart As String, ByVal pNewEnd As String) As Boolean
      Dim vConflict As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vAppointment As New ContactAppointment(mvEnv)

      vAppointment.Init()
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT contact_number FROM event_personnel WHERE session_number = " & SessionNumber)
      While vRecordSet.Fetch() = True And vConflict = False
        vAppointment.CheckCalendarConflict(vRecordSet.Fields(1).IntegerValue, pNewStart, pNewEnd, ContactAppointment.ContactAppointmentTypes.catPersonnel, SessionNumber, False)
      End While
      vRecordSet.CloseRecordSet()
    End Function

    Public Sub Update(ByRef pEvent As CDBEvent, ByRef pParams As CDBParameters)
      Dim vWeekDay As Integer

      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("SessionDesc") Then .Item(SessionFields.sfSessionDesc).Value = pParams("SessionDesc").Value
        If pParams.Exists("SessionType") Then .Item(SessionFields.sfSessionType).Value = pParams("SessionType").Value
        If pParams.Exists("Subject") Then .Item(SessionFields.sfSubject).Value = pParams("Subject").Value
        If pParams.Exists("SkillLevel") Then .Item(SessionFields.sfSkillLevel).Value = pParams("SkillLevel").Value
        If pParams.Exists("StartDate") Then .Item(SessionFields.sfStartDate).Value = pParams("StartDate").Value
        If pParams.Exists("EndDate") Then .Item(SessionFields.sfEndDate).Value = pParams("EndDate").Value
        If pParams.Exists("StartTime") Then .Item(SessionFields.sfStartTime).Value = pParams("StartTime").Value
        If pParams.Exists("EndTime") Then .Item(SessionFields.sfEndTime).Value = pParams("EndTime").Value
        If pParams.Exists("Location") Then .Item(SessionFields.sfLocation).Value = pParams("Location").Value
        If pParams.Exists("TargetAttendees") Then .Item(SessionFields.sfTargetAttendees).Value = pParams("TargetAttendees").Value
        If pParams.Exists("MinimumAttendees") Then .Item(SessionFields.sfMinimumAttendees).Value = pParams("MinimumAttendees").Value
        If pParams.Exists("MaximumAttendees") Then .Item(SessionFields.sfMaximumAttendees).Value = pParams("MaximumAttendees").Value
        If pParams.Exists("MaximumOnWaitingList") Then .Item(SessionFields.sfMaximumOnWaitingList).Value = pParams("MaximumOnWaitingList").Value
        If pParams.Exists("Notes") Then .Item(SessionFields.sfNotes).Value = pParams("Notes").Value
        If pParams.Exists("CpdApprovalStatus") Then .Item(SessionFields.sfCpdApprovalStatus).Value = pParams("CpdApprovalStatus").Value
        If pParams.Exists("CpdDateApproved") Then .Item(SessionFields.sfCpdDateApproved).Value = pParams("CpdDateApproved").Value
        If pParams.Exists("CpdAwardingBody") Then .Item(SessionFields.sfCpdAwardingBody).Value = pParams("CpdAwardingBody").Value
        If pParams.Exists("CpdCategory") Then .Item(SessionFields.sfCpdCategory).Value = pParams("CpdCategory").Value
        If pParams.Exists("CpdYear") Then .Item(SessionFields.sfCpdYear).Value = pParams("CpdYear").Value
        If pParams.Exists("CpdPoints") Then .Item(SessionFields.sfCpdPoints).Value = pParams("CpdPoints").Value
        If pParams.Exists("CpdNotes") Then .Item(SessionFields.sfCpdNotes).Value = pParams("CpdNotes").Value
        If pParams.Exists("LongDescription") Then .Item(SessionFields.sfSessionLongDesc).Value = pParams("LongDescription").Value
        If pParams.Exists("VenueBookingNumber") Then .Item(SessionFields.sfVenueBookingNumber).Value = pParams("VenueBookingNumber").Value 'Value could be null if we have not chosen a venue
        If pParams.Exists("ExternalAppointmentID") Then .Item(SessionFields.sfExternalAppointmentID).Value = pParams("ExternalAppointmentID").Value

        If SessionNumber = pEvent.LowestSessionNumber Then   'pEvent.BaseItemNumber Then
          If pEvent.External And MaximumOnWaitingList > 0 Then
            RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26126)) 'Waiting List places cannot be assigned to External Events
          ElseIf Len(pEvent.BookingsClose) > 0 Then
            If CDate(pEvent.BookingsClose) > CDate(EndDate) Then
              RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26113)) 'Bookings Close date cannot be after end date of event
            End If
          End If
          If pEvent.Template And mvEnv.GetConfigOption("opt_we_eve_from_template") = False Then
            vWeekDay = Weekday(CDate(StartDate))
            If vWeekDay = FirstDayOfWeek.Saturday Or vWeekDay = FirstDayOfWeek.Sunday Then
              RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26127)) 'Template Event Start must be a Weekday
            End If
            vWeekDay = Weekday(CDate(EndDate))
            If vWeekDay = FirstDayOfWeek.Saturday Or vWeekDay = FirstDayOfWeek.Sunday Then
              RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26128)) 'Template Event End must be a Weekday
            End If
          End If
          mvClassFields.Item(SessionFields.sfSessionDesc).Value = String.Format(ProjectText.String26125, Left(pEvent.EventDesc, 42))
        End If
        If MaximumAttendees < MinimumAttendees Then
          RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26108)) 'Maximum Attendees cannot be less than Minimum
        ElseIf MaximumAttendees < NumberOfAttendees Then
          RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26109)) 'Maximum Attendees cannot be less than current bookings count
        ElseIf MaximumOnWaitingList < NumberOnWaitingList Then
          RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26110)) 'Maximum number on waiting List cannot be less than current number on waiting list
        End If
        If CDate(StartDate) > CDate(EndDate) Then
          RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26111)) 'Start date cannot be after end date
        ElseIf CDate(StartDate) = CDate(EndDate) Then
          If IsDate(StartTime) And IsDate(EndTime) Then
            If CDate(StartTime) > CDate(EndTime) Then
              RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String26112)) 'Start time cannot be after end time
            End If
          End If
        End If
        If pEvent.Existing Then
          If CDate(StartDate) < CDate(pEvent.StartDate) Then
            RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String27207)) 'Start date cannot be prior to event start date
          ElseIf CDate(StartDate) = CDate(pEvent.StartDate) Then
            If CDate(StartTime) < CDate(pEvent.BaseSession.StartTime) Then RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String27208)) 'Start time cannot be prior to event start time
          End If
          If CDate(EndDate) > CDate(pEvent.BaseSession.EndDate) Then
            RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String27209)) 'End date cannot be after event end date
          ElseIf CDate(EndDate) = CDate(pEvent.BaseSession.EndDate) Then
            If CDate(EndTime) > CDate(pEvent.BaseSession.EndTime) Then RaiseError(DataAccessErrors.daeEventParameterError, (ProjectText.String27210)) 'End time cannot be after event end time
          End If
          If (SessionNumber <> pEvent.BaseSession.SessionNumber) And mvExisting Then
            If mvClassFields(SessionFields.sfStartDate).ValueChanged Or mvClassFields(SessionFields.sfStartTime).ValueChanged Or mvClassFields(SessionFields.sfEndDate).ValueChanged Or mvClassFields(SessionFields.sfEndTime).ValueChanged Then
              If pEvent.MoveSessionDates = False Then RaiseError(DataAccessErrors.daeEventParameterError, "Move session dates is not set for this event")
              If NumberOfAttendees + NumberInterested + NumberOnWaitingList > 0 Then RaiseError(DataAccessErrors.daeBookingsMadeCannotChange)
              PersonnelAvailable(StartDate & " " & StartTime, EndDate & " " & EndTime) 'Check personnel availability
            End If
          End If
        End If
      End With
    End Sub

    Public Sub SetNewDates()
      mvClassFields.Item(SessionFields.sfStartDate).Value = mvNewStartDate
      mvClassFields.Item(SessionFields.sfEndDate).Value = mvNewEndDate
    End Sub

    Private Function GetApprovalStatusCode(ByVal pType As CPDApprovalStatusTypes) As String
      Select Case pType
        Case CPDApprovalStatusTypes.castApproved
          GetApprovalStatusCode = "A"
        Case CPDApprovalStatusTypes.castWaitingApproval
          GetApprovalStatusCode = "W"
        Case Else
          GetApprovalStatusCode = "N"
      End Select
    End Function

    Public Sub ClearCPDData()
      mvClassFields.Item(SessionFields.sfCpdApprovalStatus).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdDateApproved).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdAwardingBody).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdCategory).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdYear).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdPoints).Value = String.Empty
      mvClassFields.Item(SessionFields.sfCpdNotes).Value = String.Empty
    End Sub
  End Class
End Namespace
