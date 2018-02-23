Namespace Access

  Public Class ContactAppointment

    Public Enum ContactAppointmentOrigin
      caoInternal = 0 'Care appointment
      caoExternal 'External appointment Oracle Calendar or Outlook
      caoBoth 'Both Care and external
    End Enum

    Private mvOrigin As ContactAppointmentOrigin
    Private mvExternalID As String

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If AppointmentTimeStatus = ContactAppointmentTimeStatuses.catsNone Then AppointmentTimeStatus = ContactAppointmentTimeStatuses.catsBusy
    End Sub

    Public Sub ModifyCalendarEntry(ByVal pContactNumber As Integer, ByVal pStart As String, ByVal pEnd As String, ByVal pUniqueID As Integer, ByVal pDesc As String)
      mvClassFields.Item(ContactAppointmentFields.ContactNumber).IntegerValue = pContactNumber
      mvClassFields.Item(ContactAppointmentFields.StartDate).Value = pStart
      mvClassFields.Item(ContactAppointmentFields.EndDate).Value = pEnd
      mvClassFields.Item(ContactAppointmentFields.Description).Value = TruncateString(pDesc, 125)
      mvClassFields.Item(ContactAppointmentFields.UniqueId).IntegerValue = pUniqueID
      Save()
    End Sub

    Public Function CheckCalendarConflict(ByVal pContactNumber As Integer, ByVal pStart As String, ByVal pEnd As String, ByVal pAppointmentType As ContactAppointmentTypes, ByVal pUniqueID As Integer, ByVal pCheckExistingEntry As Boolean, Optional ByVal pConvertInterestedBooking As Boolean = False, Optional ByVal pDoubleBook As Boolean = False) As Boolean
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vConflict As Boolean
      Dim vContact As New Contact(mvEnv)
      Dim vAppointment As New ContactAppointment(mvEnv)
      Dim vExistingCA As ContactAppointment
      Dim vSessionNumber As Integer

      If CDate(pEnd) <= CDate(pStart) Then
        RaiseError(DataAccessErrors.daeInvalidDateRange)
      Else
        vContact.Init()
        vWhereFields.Add("cc.contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        vWhereFields.Add("end_date", CDBField.FieldTypes.cftTime, pStart, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        vWhereFields.Add("start_date", CDBField.FieldTypes.cftTime, pEnd, CDBField.FieldWhereOperators.fwoLessThanEqual)
        vWhereFields.Add("c.contact_number", CDBField.FieldTypes.cftLong, "cc.contact_number")
        vWhereFields.Add("time_status", CDBField.FieldTypes.cftCharacter, "B")
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT record_type, unique_id, description, start_date, end_date, " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName) & " FROM contact_appointments cc, contacts c WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
        Dim vName As String = ""
        While vRecordSet.Fetch() And vConflict = False
          If pAppointmentType = ContactAppointmentTypes.catPersonnelTask Then
            'In this case, pUniqueID supplied = Session Number
            Select Case vRecordSet.Fields(1).Value
              Case GetAppointmentTypeCode(ContactAppointmentTypes.catPersonnel)
                'OK to overlap if for same Event
                If CInt(vRecordSet.Fields(2).LongValue / 10000) <> CInt(pUniqueID / 10000) Then
                  vConflict = True
                End If
              Case GetAppointmentTypeCode(ContactAppointmentTypes.catPersonnelTask)
                'OK to overlap if for same Event
                vSessionNumber = CInt(Val(mvEnv.Connection.GetValue("SELECT session_number FROM event_personnel_tasks ept, event_personnel ep WHERE ept.event_personnel_task_number = " & vRecordSet.Fields(2).LongValue & " AND ep.event_personnel_number = ept.event_personnel_number")))
                If vSessionNumber > 0 Then
                  If CInt(vSessionNumber / 10000) <> CInt(pUniqueID / 10000) Then
                    vConflict = True
                  End If
                End If
              Case Else
                vConflict = True
            End Select
          Else
            If vRecordSet.Fields(1).Value <> GetAppointmentTypeCode(pAppointmentType) Then 'If a different type then definately a conflict
              If pAppointmentType = ContactAppointmentTypes.catPersonnel And vRecordSet.Fields(1).Value = GetAppointmentTypeCode(ContactAppointmentTypes.catPersonnelTask) Then
                'Personnel/Personnel Task appointments - OK to conflict if same Event
                vSessionNumber = CInt(Val(mvEnv.Connection.GetValue("SELECT session_number FROM event_personnel_tasks ept, event_personnel ep WHERE ept.event_personnel_task_number = " & vRecordSet.Fields(2).LongValue & " AND ep.event_personnel_number = ept.event_personnel_number")))
                If vSessionNumber > 0 Then
                  If CInt(vSessionNumber / 10000) <> CInt(pUniqueID / 10000) Then
                    vConflict = True
                  End If
                End If
              Else
                vConflict = True
              End If
            Else
              If vRecordSet.Fields(2).LongValue <> pUniqueID Then 'Same type of appointment but different number
                If pAppointmentType = ContactAppointmentTypes.catEvent Then
                  If CInt(vRecordSet.Fields(2).LongValue / 10000) = CInt(pUniqueID / 10000) Then 'Same event? Could be overlapping sessions
                    vConflict = False
                  Else
                    vConflict = True
                  End If
                ElseIf pAppointmentType = ContactAppointmentTypes.catPersonnel Then
                  If pUniqueID Mod 10000 = 0 And vRecordSet.Fields(2).LongValue Mod 10000 <> 0 Then 'Just look at remainder
                    'pUniqueID = lowest session number & not an actual session
                    'Reset pUniqueID to recordset value to check for conflict with next record
                    pUniqueID = vRecordSet.Fields(2).LongValue
                  Else
                    vConflict = True
                  End If
                Else
                  vConflict = True
                End If
              Else
                If pCheckExistingEntry Then
                  If pConvertInterestedBooking Then
                    vConflict = False
                    Return False
                  Else
                    vConflict = True 'A record exists for the same unique id - conflict only if specified
                  End If
                Else
                  vExistingCA = New ContactAppointment(mvEnv)
                  vExistingCA.Init()
                  vExistingCA.Create(pContactNumber, CDate(vRecordSet.Fields(4).Value), CDate(vRecordSet.Fields(5).Value), 0, GetAppointmentTypeCode(pAppointmentType), vRecordSet.Fields(3).Value, ContactAppointmentTimeStatuses.catsNone)
                  If pAppointmentType = ContactAppointmentTypes.catOther Then
                    If CDate(vRecordSet.Fields(4).Value) <> CDate(StartDate) Or CDate(vRecordSet.Fields(5).Value) <> CDate(EndDate) Then
                      vConflict = True
                    End If
                  End If
                End If
              End If
            End If
          End If
          If vConflict Then
            vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtName)
            vName = vContact.Name & " : " & vRecordSet.Fields(3).Value & " : " & vRecordSet.Fields(4).Value
          End If
        End While
        vRecordSet.CloseRecordSet()
        If vConflict Then
          'Dont raise an error if the user wants to double book
          If pAppointmentType = ContactAppointmentTypes.catServiceBooking AndAlso pDoubleBook Then Return True
          RaiseError(DataAccessErrors.daeAppointmentConflict, vName)
        End If
      End If
      Return True
    End Function

    Public Sub ClearEntries(ByVal pAppointmentType As ContactAppointmentTypes, ByVal pUniqueID As Integer, Optional ByRef pContactNumber As Integer = 0, Optional ByRef pInternalOnly As Boolean = False)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("unique_id", pUniqueID)
      vWhereFields.Add("record_type", GetAppointmentTypeCode(pAppointmentType))
      If pContactNumber > 0 Then vWhereFields.Add("contact_number", pContactNumber)
      mvEnv.Connection.DeleteRecords(DatabaseTableName, vWhereFields, False)
    End Sub

    Public Overloads Sub Create(ByVal pContactNumber As Integer, ByVal pStart As String, ByVal pEnd As String, ByVal pAppointmentType As ContactAppointmentTypes, ByVal pDesc As String, Optional ByRef pUnique As Integer = 0, Optional ByRef pTimeStatus As ContactAppointmentTimeStatuses = ContactAppointmentTimeStatuses.catsNone, Optional ByRef pExternalID As String = "", Optional ByVal pStatus As EventBooking.EventBookingStatuses = 0, Optional ByVal pOutlookId As String = "")
      'BR11719 - Added optional stsus parameter and code below to deal with

      AppointmentType = pAppointmentType
      mvClassFields.Item(ContactAppointmentFields.ContactNumber).Value = CStr(pContactNumber)
      mvClassFields.Item(ContactAppointmentFields.StartDate).Value = pStart
      mvClassFields.Item(ContactAppointmentFields.EndDate).Value = pEnd
      mvClassFields.Item(ContactAppointmentFields.Description).Value = Left(pDesc, 125)
      If pUnique > 0 Then mvClassFields.Item(ContactAppointmentFields.UniqueId).LongValue = pUnique
      If pTimeStatus <> ContactAppointmentTimeStatuses.catsNone Then AppointmentTimeStatus = pTimeStatus
      If pAppointmentType = ContactAppointmentTypes.catExternal Then
        mvOrigin = ContactAppointmentOrigin.caoExternal
      Else
        mvOrigin = ContactAppointmentOrigin.caoInternal
      End If
      mvExternalID = pExternalID
      ' BR11719
      If (AppointmentType = ContactAppointmentTypes.catEvent) And (mvEnv.GetConfigOption("ev_show_int_only_bkgs_as_free", False)) And (pStatus = EventBooking.EventBookingStatuses.ebsInterested) Then
        AppointmentTimeStatus = ContactAppointmentTimeStatuses.catsFree
      End If
      'BR17931
      If pOutlookId <> "" Then
        mvClassFields.Item(ContactAppointmentFields.OutlookId).Value = pOutlookId
      End If

    End Sub

    Public Sub SetEntryStatus(ByVal pAppointmentType As ContactAppointmentTypes, ByVal pUniqueID As Integer, ByRef pTimeStatus As ContactAppointmentTimeStatuses, Optional ByRef pContactNumber As Integer = 0)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("unique_id", pUniqueID)
      vWhereFields.Add("record_type", GetAppointmentTypeCode(pAppointmentType))
      If pContactNumber > 0 Then vWhereFields.Add("contact_number", pContactNumber)
      AppointmentTimeStatus = pTimeStatus
      Dim vFields As New CDBFields
      vFields.Add("time_status", mvClassFields(ContactAppointmentFields.TimeStatus).Value)
      mvEnv.Connection.UpdateRecords((mvClassFields.DatabaseTableName), vFields, vWhereFields, False)
    End Sub

    Public Overloads Sub Update(ByVal pStartDate As String, ByVal pEndDate As String, ByVal pDescription As String)
      mvClassFields.Item(ContactAppointmentFields.StartDate).Value = pStartDate
      mvClassFields.Item(ContactAppointmentFields.EndDate).Value = pEndDate
      mvClassFields.Item(ContactAppointmentFields.Description).Value = TruncateString(pDescription, 125)
    End Sub

    Public Function GetAppointmentsDT(ByVal pContactNumber As Integer, ByVal pStartDate As Date, ByVal pEndDate As Date) As CDBDataTable
      Dim vDataTable As New CDBDataTable
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      vWhereFields.Add("end_date", CDBField.FieldTypes.cftTime, pStartDate.ToString(CAREDateTimeFormat), CDBField.FieldWhereOperators.fwoGreaterThan)
      vWhereFields.Add("start_date", CDBField.FieldTypes.cftTime, pEndDate.ToString(CAREDateTimeFormat), CDBField.FieldWhereOperators.fwoLessThan)
      Dim vSQL As New SQLStatement(mvEnv.Connection, GetRecordSetFields, "contact_appointments", vWhereFields, "start_date")
      vDataTable.FillFromSQL(mvEnv, vSQL)
      Return vDataTable
    End Function

  End Class
End Namespace


