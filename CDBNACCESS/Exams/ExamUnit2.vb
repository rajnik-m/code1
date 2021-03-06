Namespace Access

  Partial Public Class ExamUnit

#Region "Non AutoGenerated Code"

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If mvClassFields(ExamUnitFields.ExamBaseUnitId).Value = "" Then mvClassFields(ExamUnitFields.ExamBaseUnitId).IntegerValue = 0
    End Sub

    Public Sub ValidateExamBookingUnits(ByVal pExamSession As ExamSession, ByVal pExamCentre As ExamCentre, ByVal pExamUnitLink As ExamUnitLink, ByVal pUnits As String, ByVal pUnitLinks As String, ByVal pContact As Contact, Optional ByVal pBookingDate As String = "")
      'Check if this unit belongs to the correct session
      If Me.ExamSessionId <> pExamSession.ExamSessionId Then RaiseError(DataAccessErrors.daeExamBookingUnitSession, Me.ExamUnitId.ToString, pExamSession.ExamSessionCode)

      If Me.ExamUnitId <> pExamUnitLink.ExamUnitId2 Then RaiseError(DataAccessErrors.daeExamBookingUnitLinkNotForUnit, Me.ExamUnitId.ToString, pExamUnitLink.ExamUnitId2.ToString)

      If Me.ExamSessionId > 0 Then
        'Check the centre is available for this session
        Dim vSessionCentre As New ExamSessionCentre(mvEnv)
        vSessionCentre.Init(pExamSession.ExamSessionId, pExamCentre.ExamCentreId)
        If vSessionCentre.Existing = False Then RaiseError(DataAccessErrors.daeExamBookingInvalidCentre)

        'Check if we are allowed to use this centre - use valid from and to
        Dim vTransactionDate As Date = Date.Today
        If pBookingDate.Length > 0 Then vTransactionDate = CDate(pBookingDate)
        If pExamCentre.ValidFrom.Length > 0 AndAlso vTransactionDate < CDate(pExamCentre.ValidFrom) Then RaiseError(DataAccessErrors.daeExamBookingInvalidCentre, pExamCentre.ExamCentreCode)
        If pExamCentre.ValidTo.Length > 0 AndAlso vTransactionDate > CDate(pExamCentre.ValidTo) Then RaiseError(DataAccessErrors.daeExamBookingInvalidCentre, pExamCentre.ExamCentreCode)
      End If

      'Now check the home and overseas closing dates - This can be overwridded from the client so commented out for now
      'Dim vClosingDate As String
      'If pExamCentre.Overseas Then
      '  vClosingDate = pExamSession.OverseasClosingDate
      'Else
      '  vClosingDate = pExamSession.HomeClosingDate
      'End If
      'If IsDate(vClosingDate) AndAlso vTransactionDate > CDate(vClosingDate) Then

      'End If

      'Need to validate that all the units passed in exist 
      Dim vExamUnits As New StringList(pUnits, ",")
      Dim vExamUnitLinks As New StringList(pUnitLinks, ",")
      Dim vUnitWhereFields As New CDBFields
      vUnitWhereFields.Add("exam_unit_id", pUnits, CDBField.FieldWhereOperators.fwoIn)
      Dim vUnits As List(Of ExamUnit) = Me.GetList(Of ExamUnit)(Me, vUnitWhereFields)
      'Need the LinkId's from ExamUnitLink so find the record and attach it to the ExamUnit so we can get it later
      vUnitWhereFields.Clear()
      vUnitWhereFields.Add("exam_unit_id_2", CDBField.FieldTypes.cftInteger, pUnits, CDBField.FieldWhereOperators.fwoIn)
      vUnitWhereFields.Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger, pUnitLinks, CDBField.FieldWhereOperators.fwoIn)
      Dim vExamUnitLink As New ExamUnitLink(mvEnv)
      vExamUnitLink.Init()
      Dim vUnitLinks As List(Of ExamUnitLink) = vExamUnitLink.GetList(Of ExamUnitLink)(vExamUnitLink, vUnitWhereFields)
      For Each vUnit As ExamUnit In vUnits
        For Each vExamUnitLink In vUnitLinks
          If vUnit.ExamUnitId = vExamUnitLink.ExamUnitId2 Then
            vUnit.UnitLink = vExamUnitLink
            Exit For
          End If
        Next
      Next

      If vExamUnitLinks.Count <> vExamUnits.Count Then RaiseError(DataAccessErrors.daeExamBookingLinksCountMismatch)
      For Each vExamUnit As String In vExamUnits
        'Check that each exam unit has only been provided once in the list (can appear multiple times in a tree now as we can share units)
        If vExamUnits.FindAll(Function(item) item = vExamUnit).Count > 1 Then RaiseError(DataAccessErrors.daeExamBookingUnitMultipleInSessionBooking, vExamUnit)
        Dim vFound As Boolean = False
        For Each vSelectedUnit As ExamUnit In vUnits
          If vSelectedUnit.ExamUnitId = IntegerValue(vExamUnit) Then
            vFound = True
            Exit For
          End If
        Next
        If vFound = False Then RaiseError(DataAccessErrors.daeExamBookingUnitMissing, vExamUnit)
      Next
      'Check the units are for the correct session and are booking allowed
      'And that the contact is allowed to book on the unit in cases where there are exclude or include views
      For Each vSelectedUnit As ExamUnit In vUnits
        'Is the unit for the correct session
        If vSelectedUnit.ExamSessionId <> pExamSession.ExamSessionId Then RaiseError(DataAccessErrors.daeExamBookingUnitSession, vSelectedUnit.ExamUnitId.ToString, pExamSession.ExamSessionCode)
        'Does the unit have allow bookings
        If vSelectedUnit.AllowBookings = False Then RaiseError(DataAccessErrors.daeExamBookingUnitAllowed, vSelectedUnit.ExamUnitId.ToString)
        'Validate the contact against Include/Exclude views for the unit
        Dim vViewWhereFields As New CDBFields
        vViewWhereFields.Add("contact_number", pContact.ContactNumber)
        If vSelectedUnit.ExcludeView.Length > 0 Then
          If mvEnv.Connection.GetCount(vSelectedUnit.ExcludeView, vViewWhereFields) > 0 Then _
            RaiseError(DataAccessErrors.daeExamBookingExcludeViewValidationFailed, vSelectedUnit.ExamUnitCode)
        End If
        If vSelectedUnit.IncludeView.Length > 0 Then
          If mvEnv.Connection.GetCount(vSelectedUnit.IncludeView, vViewWhereFields) < 1 Then _
            RaiseError(DataAccessErrors.daeExamBookingIncludeViewValidationFailed, vSelectedUnit.ExamUnitCode)
        End If
      Next
      'Now get the heirarchy of units for this session including booking and passed details
      Dim vParams As New CDBParameters
      vParams.Add("ContactNumber", CDBField.FieldTypes.cftInteger, pContact.ContactNumber.ToString)
      If pExamSession.ExamSessionId > 0 Then vParams.Add("ExamSessionId", CDBField.FieldTypes.cftInteger, pExamSession.ExamSessionId.ToString)
      vParams.Add("AllowBookings", "Y")
      vParams.Add("ValidateParameters", "Y")
      Dim vDS As New ExamDataSelection(mvEnv, ExamDataSelection.ExamDataSelectionTypes.dstExamStudentBookingUnits, vParams, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      Dim vDT As CDBDataTable = vDS.DataTable

      Dim vExamUnitLinkId As Integer = 0
      Dim vParentUnitLinkId As Integer = 0
      For Each vSelectedUnit As ExamUnit In vUnits
        'For each unit first find it in the list of units for the session that allow booking
        vExamUnitLinkId = vSelectedUnit.UnitLink.ExamUnitLinkId
        vParentUnitLinkId = vSelectedUnit.UnitLink.ParentUnitLinkId

        Dim vRow As CDBDataRow = vDT.FindRow("ExamUnitLinkId", vExamUnitLinkId.ToString)
        If vRow Is Nothing Then RaiseError(DataAccessErrors.daeExamBookingUnitMissing, vExamUnitLinkId.ToString)
        'Next check there is not already a booking in the same session
        If vRow.Item("Booked").StartsWith("Y") Then RaiseError(DataAccessErrors.daeExamBookingExists, vSelectedUnit.ExamUnitId.ToString)
        'Now check the unit has not already been passed
        If vRow.Item("Passed").StartsWith("Y") AndAlso vRow.Item("ScheduleRequired").StartsWith("Y") Then RaiseError(DataAccessErrors.daeExamBookingUnitPassed, vSelectedUnit.ExamUnitCode)

        'Now we need to check that there is either an existing booking for any parent or that the parent is included in the list of units
        Dim vParentUnit As Integer = IntegerValue(vRow.Item("ParentUnitLinkId"))
        If vParentUnit > 0 Then
          If Not vExamUnitLinks.Contains(vParentUnit.ToString) Then
            'The parent is not in the list of items being booked
            Dim vParentRow As CDBDataRow = vDT.FindRow("ExamUnitLinkId", vParentUnit.ToString)
            If vParentRow Is Nothing Then RaiseError(DataAccessErrors.daeExamBookingUnitMissing, vParentUnit.ToString)
            If Not vParentRow.Item("Booked").StartsWith("Y") Then RaiseError(DataAccessErrors.daeExamBookingParentNotBooked, vExamUnitLinkId.ToString, vParentUnit.ToString)
          End If
        End If
      Next

      ' Check for prerequisite exam grades
      'Get list of ALL past exams sat by candidate
      vParams.Clear()
      vParams.Add("ContactNumber", CDBField.FieldTypes.cftInteger, pContact.ContactNumber.ToString)
      vDS = New ExamDataSelection(mvEnv, ExamDataSelection.ExamDataSelectionTypes.dstExamStudentUnitHeader, vParams, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      vDT = vDS.DataTable

      'Get list of prerequisite exam unit grades require to allow booking on unit
      Dim vPreReqFields As String = "eup.exam_prerequisite_unit_id, eg.exam_grade, eg.sequence_number, eu.exam_unit_code, eup.pass_required"
      Dim vPreReqWhereFields As New CDBFields
      Dim vPreReqAnsiJoins As New AnsiJoins
      vPreReqAnsiJoins.AddLeftOuterJoin("exam_grades eg", "eup.minimum_grade", "eg.exam_grade")
      vPreReqAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "eup.exam_prerequisite_unit_id") ' NOTE: eup.exam_prerequisite_unit_id is the zero-session based exam_unit_id

      For Each vSelectedUnit As ExamUnit In vUnits
        vPreReqWhereFields.Clear()
        vPreReqWhereFields.Add("eup.exam_unit_id", vSelectedUnit.ExamUnitId) ' NOTE: eup.exam_unit_id is the session based exam_unit_id
        Dim vPreReqSQL As New SQLStatement(mvEnv.Connection, vPreReqFields, "exam_unit_prerequisites eup", vPreReqWhereFields, "", vPreReqAnsiJoins)
        Dim vPreReqDT As CDBDataTable = New CDBDataTable()
        vPreReqDT.FillFromSQL(mvEnv, vPreReqSQL)

        For Each vPreReqRow As CDBDataRow In vPreReqDT.Rows
          Dim vExistingStudentUnitRow As CDBDataRow = vDT.FindRow("ExamUnitId", vPreReqRow.Item("exam_prerequisite_unit_id").ToString)
          If vExistingStudentUnitRow Is Nothing Then RaiseError(DataAccessErrors.daeExamBookingUnitPrerequisiteMissing, vSelectedUnit.ExamUnitCode, vPreReqRow.Item("exam_unit_code")) ' exam unit not sat

          ' Check candidate's result has not expired
          If IsDate(vExistingStudentUnitRow.Item("Expires")) Then
            If CDate(vExistingStudentUnitRow.Item("Expires")) < Date.Today Then RaiseError(DataAccessErrors.daeExamBookingUnitPrerequisiteExpired, vSelectedUnit.ExamUnitCode, vExistingStudentUnitRow.Item("CurrentGrade"))
          End If

          ' Check for a simple Pass requirement
          If vPreReqRow.Item("pass_required").ToString = "Y" Then
            If vExistingStudentUnitRow.Item("CurrentResult") <> "P" Then RaiseError(DataAccessErrors.daeExamBookingUnitPrerequisitePassRequired, vSelectedUnit.ExamUnitCode, vPreReqRow.Item("exam_unit_code"))
          Else
            ' Validate that grade exists
            If vExistingStudentUnitRow.Item("CurrentGrade").Length < 1 Then RaiseError(DataAccessErrors.daeExamBookingUnitPrerequisiteMissing, vSelectedUnit.ExamUnitCode, vPreReqRow.Item("exam_unit_code")) ' Exam Grade missing so probably booked but not yet sat

            ' Validate grade sequences
            If ((vPreReqRow.Item("sequence_number").Length < 1) Or (IntegerValue(vPreReqRow.Item("sequence_number")) < 0)) Then RaiseError(DataAccessErrors.daeExamGradeSequenceMissing, vExistingStudentUnitRow.Item("ExamUnitCode"), vPreReqRow.Item("exam_grade")) ' Exam Grade sequence missing
            If ((vExistingStudentUnitRow.Item("ExamGradeSequenceNumber").Length < 1) Or (IntegerValue(vExistingStudentUnitRow.Item("ExamGradeSequenceNumber")) < 0)) Then RaiseError(DataAccessErrors.daeExamGradeSequenceMissing, vSelectedUnit.ExamUnitCode, vExistingStudentUnitRow.Item("CurrentGrade")) ' Exam Grade sequence missing

            ' compare required grade's sequence to attained grade's sequence
            If IntegerValue(vPreReqRow.Item("sequence_number").ToString) > IntegerValue(vExistingStudentUnitRow.Item("ExamGradeSequenceNumber")) Then RaiseError(DataAccessErrors.daeExamBookingUnitPrerequisiteMinimumGrade, vSelectedUnit.ExamUnitCode, vPreReqRow.Item("exam_unit_code")) ' exam unit sat but grade is not good enough
          End If
        Next
      Next

    End Sub

    Public Function AddExamBooking(ByVal pExamSession As ExamSession, ByVal pExamCentre As ExamCentre, ByVal pExamUnitLink As ExamUnitLink, ByVal pUnits As String, ByVal pUnitLinks As String, ByVal pContact As Contact, ByVal pAddressNumber As Integer, ByVal pAmount As Double, Optional ByVal pNotes As String = "", Optional ByVal pBookingDate As String = "", Optional ByVal pSalesContactNumber As Integer = 0, Optional ByVal pCourseStartDate As String = "", Optional ByVal pAssessmentLanguage As String = "", Optional ByVal pLines As CollectionList(Of ExamBookingLine) = Nothing, Optional ByVal pStudyMode As String = "") As ExamBooking

      Dim vUnitsList As New StringList(pUnits, ",")
      Dim vLinksList As New StringList(pUnitLinks, ",")
      Dim vExamBooking As ExamBooking = Nothing
      'If the unit type specifies that it requires a schedule then check that a schedule record exists
      Dim vSchedules As New List(Of ExamSchedule)
      Dim vUnitMappings As New List(Of ExamUnitMapping)     'Hold a list of the mappings from the session units to the template units
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      vWhereFields.Add("eu.exam_unit_id", pUnits, CDBField.FieldWhereOperators.fwoIn)
      vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")
      vAnsiJoins.Add("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")
      vWhereFields.Add("eul.exam_unit_link_id", pUnitLinks, CDBField.FieldWhereOperators.fwoIn)
      Dim vUnitSQL As New SQLStatement(mvEnv.Connection, "exam_unit_id,schedule_required,exam_unit_code,exam_base_unit_id,exam_question,exam_unit_link_id,base_unit_link_id", "exam_units eu", vWhereFields, "", vAnsiJoins)
      Dim vRS As CDBRecordSet = vUnitSQL.GetRecordSet
      While vRS.Fetch
        If vRS.Fields(2).Bool Then 'Schedule required
          Dim vSchedule As New ExamSchedule(mvEnv)
          Dim vScheduleFields As New CDBFields
          vScheduleFields.Add("exam_session_id", pExamSession.ExamSessionId)
          vScheduleFields.Add("exam_centre_id", pExamCentre.ExamCentreId)
          vScheduleFields.Add("exam_unit_id", vRS.Fields(1).IntegerValue)
          vSchedule.InitWithPrimaryKey(vScheduleFields)
          'Does the schedule exist for this combination of session, centre and unit
          If vSchedule.Existing = False Then RaiseError(DataAccessErrors.daeNoScheduleDefined, pExamSession.ExamSessionCode, pExamCentre.ExamCentreCode, vRS.Fields(3).Value)
          'If so does it have capacity
          If vSchedule.HasCapacity = False Then RaiseError(DataAccessErrors.daeNoScheduleAtCapacity, pExamSession.ExamSessionCode, pExamCentre.ExamCentreCode, vRS.Fields(3).Value)
          If pExamCentre.HasCapacity(vSchedule) = False Then RaiseError(DataAccessErrors.daeNoCentreCapacity, pExamSession.ExamSessionCode, pExamCentre.ExamCentreCode, vRS.Fields(3).Value)
          vSchedules.Add(vSchedule)
        End If
        Dim vIdx As Integer = vUnitsList.IndexOf(vRS.Fields(1).Value.ToString)
        If vIdx < 0 OrElse vLinksList(vIdx) <> vRS.Fields(6).Value.ToString Then RaiseError(DataAccessErrors.daeExamBookingUnitLinkNotForUnit, vLinksList(vIdx), vRS.Fields(6).ToString)
        Dim vUnitMapping As New ExamUnitMapping(vRS.Fields(1).IntegerValue, vRS.Fields(4).IntegerValue, vRS.Fields(6).IntegerValue, vRS.Fields(7).IntegerValue, vRS.Fields(5).Bool, vRS.Fields(2).Bool)
        vUnitMappings.Add(vUnitMapping)
      End While
      vRS.CloseRecordSet()
      'OK All the units requiring a schedule have one

      'Check booking schedules do not clash with each other
      Dim vClashFoundInCurrentBooking As Boolean = False
      Dim vClashDate As DateTime
      If vSchedules.Count > 1 Then
        For Each vSchedule As ExamSchedule In vSchedules
          For Each vInnerSchedule As ExamSchedule In vSchedules
            If Not Object.ReferenceEquals(vSchedule, vInnerSchedule) Then
              vClashFoundInCurrentBooking = vSchedule.ClashCheck(vInnerSchedule)
              If vClashFoundInCurrentBooking Then
                vClashDate = CDate(vInnerSchedule.StartDate)
                Exit For
              End If
            End If
          Next
          If vClashFoundInCurrentBooking Then Exit For
        Next
      End If
      If vClashFoundInCurrentBooking Then RaiseError(DataAccessErrors.daeScheduleClashInBooking, vClashDate.ToString(CAREDateFormat))

      'Check booking schedules do not clash with existing bookings in the same session
      Dim vBooking As New ExamBooking(mvEnv)
      Dim vBookingWhereFields As New CDBFields
      vBookingWhereFields.Add("contact_number", pContact.ContactNumber)
      vBookingWhereFields.Add("cancellation_reason")
      Dim vBookings As List(Of ExamBooking) = vBooking.GetList(Of ExamBooking)(vBooking, vBookingWhereFields)
      For Each vIdxBooking As ExamBooking In vBookings
        Dim vOldSchedules As List(Of ExamSchedule) = vIdxBooking.GetListOfSchedules()
        For Each vSchedule As ExamSchedule In vSchedules
          For Each vOldSchedule As ExamSchedule In vOldSchedules
            vClashFoundInCurrentBooking = vSchedule.ClashCheck(vOldSchedule)
            If vClashFoundInCurrentBooking Then
              vClashDate = CDate(vOldSchedule.StartDate)
              Exit For
            End If
          Next
          If vClashFoundInCurrentBooking Then Exit For
        Next
      Next
      If vClashFoundInCurrentBooking Then RaiseError(DataAccessErrors.daeScheduleClashExistingBooking, vClashDate.ToString(CAREDateFormat))

      'Check if there is already a exam_student_header for this exam unit
      'If so update it, if not then create it - It has the base unit id in it (from the template)
      Dim vStudentHeader As New ExamStudentHeader(mvEnv)
      Dim vStudentWhereFields As New CDBFields
      Dim vExamBaseUnitLinkId As Integer = 0
      If ExamBaseUnitId > 0 Then
        vStudentWhereFields.Add("exam_unit_id", ExamBaseUnitId)
        'vStudentWhereFields.Add("exam_unit_link_id", vExamBaseUnitLinkId)  You can only have one header per unit, so this is no necessary.
        'If vStudentWhereFields("exam_unit_link_id").IntegerValue = 0 Then RaiseError(DataAccessErrors.daeExamUnitBaseLinkNotFound, Me.ExamUnitDescription)
        vExamBaseUnitLinkId = pExamUnitLink.GetBaseUnitLinkId(mvEnv)
      Else
        vStudentWhereFields.Add("exam_unit_id", ExamUnitId)
        'vStudentWhereFields.Add("exam_unit_link_id", pExamUnitLink.ExamUnitLinkId)   You can only have one header per unit, so this is no necessary.
        vExamBaseUnitLinkId = pExamUnitLink.ExamUnitLinkId
      End If
      vStudentWhereFields.Add("contact_number", pContact.ContactNumber)
      vStudentHeader.InitWithPrimaryKey(vStudentWhereFields)

      Dim vTrans As Boolean = False
      If mvEnv.Connection.InTransaction = False Then

        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      Dim vStudentParams As New CDBParameters
      If vStudentHeader.Existing Then
        If pExamSession.ExamSessionId > 0 Then
          vStudentParams.Add("LastSessionId", pExamSession.ExamSessionId)
        End If
        If vStudentHeader.ExamUnitLinkId <> vExamBaseUnitLinkId Then
          vStudentParams.Add("ExamUnitLinkId").Value = vExamBaseUnitLinkId.ToString
        End If
        'Only update if required.
        If vStudentParams.Count > 0 Then vStudentHeader.Update(vStudentParams)
      Else
        vStudentParams.Add("ContactNumber", pContact.ContactNumber)
        If pExamSession.ExamSessionId > 0 Then
          vStudentParams.Add("ExamUnitId", ExamBaseUnitId)
          vStudentParams.Add("ExamUnitLinkId").Value = vExamBaseUnitLinkId.ToString
          vStudentParams.Add("FirstSessionId", pExamSession.ExamSessionId)
          vStudentParams.Add("LastSessionId", pExamSession.ExamSessionId)
        Else
          vStudentParams.Add("ExamUnitId", ExamUnitId)
          vStudentParams.Add("ExamUnitLinkId").Value = pExamUnitLink.ExamUnitLinkId.ToString
        End If
        vStudentHeader.Create(vStudentParams)
      End If
      vStudentHeader.Save(mvEnv.User.UserID)

      'Create the exam booking record
      Dim vBookingParams As New CDBParameters
      vBookingParams.Add("ContactNumber", pContact.ContactNumber)
      vBookingParams.Add("AddressNumber", pAddressNumber)
      vBookingParams.Add("ExamSessionId", pExamSession.ExamSessionId)
      vBookingParams.Add("ExamCentreId", pExamCentre.ExamCentreId)
      vBookingParams.Add("ExamUnitId", ExamUnitId)
      vBookingParams.Add("ExamUnitLinkId", pExamUnitLink.ExamUnitLinkId)
      vBookingParams.Add("Amount", pAmount)
      vBookingParams.Add("SpecialRequirements", "N")
      vBookingParams.Add("StudyMode", pStudyMode)
      vExamBooking = New ExamBooking(mvEnv)
      vExamBooking.Create(vBookingParams)
      If pLines IsNot Nothing AndAlso pLines.Count > 0 Then
        vExamBooking.SetTransactionInfo(pLines(0).BatchNumber, pLines(0).TransactionNumber)
      End If
      vExamBooking.Save(mvEnv.User.UserID)

      'Create the exam booking unit for each specified unit
      Dim vBookingUnitParams As New CDBParameters
      vBookingUnitParams.Add("ExamBookingId", vExamBooking.ExamBookingId)
      vBookingUnitParams.Add("ExamUnitId")
      vBookingUnitParams.Add("ExamScheduleId")
      vBookingUnitParams.Add("ContactNumber", pContact.ContactNumber)
      vBookingUnitParams.Add("ExamUnitLinkId")
      vBookingUnitParams.Add("AddressNumber", pAddressNumber)
      vBookingUnitParams.Add("AttemptNumber")
      vBookingUnitParams.Add("ExamStudentUnitStatus")
      vBookingUnitParams.Add("CourseStartDate", CDBField.FieldTypes.cftDate, pCourseStartDate)
      vBookingUnitParams.Add("ExamAssessmentLanguage", pAssessmentLanguage)

      Dim vExamBookingUnit As ExamBookingUnit
      For Each vUnitMapping As ExamUnitMapping In vUnitMappings
        If vUnitMapping.ExamQuestion = False Then
          'If it is not a question
          'Check if there is already a exam_student_unit_header for this exam unit
          'If so update it, if not then create it
          Dim vStudentUnitHeader As New ExamStudentUnitHeader(mvEnv)
          Dim vStudentUnitWhereFields As New CDBFields
          vStudentUnitWhereFields.Add("exam_student_header_id", vStudentHeader.ExamStudentHeaderId)
          vStudentUnitWhereFields.Add("exam_unit_id", vUnitMapping.BaseUnitId)
          vExamBaseUnitLinkId = vUnitMapping.BaseUnitLinkId
          vStudentUnitHeader.InitWithPrimaryKey(vStudentUnitWhereFields)
          Dim vStudentUnitParams As New CDBParameters
          If vStudentUnitHeader.Existing Then
            If vUnitMapping.ExamSchedule Then
              vStudentUnitParams.Add("Attempts", vStudentUnitHeader.NumberOfAttempts + 1)
            End If
            If vStudentUnitHeader.ExamUnitLinkId <> vUnitMapping.BaseUnitLinkId Then
              vStudentUnitParams.Add("ExamUnitLinkId", vUnitMapping.BaseUnitLinkId)
            End If
            If vStudentUnitParams.Count > 0 Then vStudentUnitHeader.Update(vStudentUnitParams)
          Else
            vStudentUnitParams.Add("ExamStudentHeaderId", vStudentHeader.ExamStudentHeaderId)
            vStudentUnitParams.Add("ExamUnitId", vUnitMapping.BaseUnitId)
            vStudentUnitParams.Add("ExamUnitLinkId", vUnitMapping.BaseUnitLinkId)
            vStudentUnitParams.Add("ExamBookingId", vBookingUnitParams("ExamBookingId"))
            If vUnitMapping.ExamSchedule Then vStudentUnitParams.Add("Attempts", 1)
            vStudentUnitHeader.Create(vStudentUnitParams)
          End If
          vStudentUnitHeader.Save(mvEnv.User.UserID)
          If vUnitMapping.ExamSchedule Then
            vBookingUnitParams("AttemptNumber").Value = vStudentUnitHeader.Attempts.ToString
            If vStudentUnitHeader.NumberOfAttempts = 1 Then
              vBookingUnitParams("ExamStudentUnitStatus").Value = "F"
            Else
              vBookingUnitParams("ExamStudentUnitStatus").Value = "R"
            End If
          Else
            vBookingUnitParams("AttemptNumber").Value = ""
            vBookingUnitParams("ExamStudentUnitStatus").Value = ""
          End If
        Else
          vBookingUnitParams("AttemptNumber").Value = ""
          vBookingUnitParams("ExamStudentUnitStatus").Value = ""
        End If
        For Each vSchedule As ExamSchedule In vSchedules
          'If we have a schedule record for it then create the link
          If vSchedule.ExamUnitId = vUnitMapping.ExamUnitId Then
            vBookingUnitParams("ExamScheduleId").Value = vSchedule.ExamScheduleId.ToString
            Exit For
          Else
            vBookingUnitParams("ExamScheduleId").Value = ""
          End If
        Next
        vBookingUnitParams("ExamUnitId").Value = vUnitMapping.ExamUnitId.ToString
        vBookingUnitParams("ExamUnitLinkId").Value = vUnitMapping.ExamUnitLinkId.ToString
        vExamBookingUnit = New ExamBookingUnit(mvEnv)
        vExamBookingUnit.Create(vBookingUnitParams)
        If pLines IsNot Nothing Then
          For Each vBookingLine As ExamBookingLine In pLines
            If vBookingLine.ExamUnitId = vExamBookingUnit.ExamUnitId AndAlso vBookingLine.ExamUnitProductId = 0 Then
              'Set the transaction info for the booking unit
              vExamBookingUnit.SetTransactionInfo(vBookingLine.BatchNumber, vBookingLine.TransactionNumber, vBookingLine.TransactionLineNumber)
            End If
          Next
        End If
        vExamBookingUnit.Save(mvEnv.User.UserID)
        If pLines IsNot Nothing Then
          For Each vBookingLine As ExamBookingLine In pLines
            If vBookingLine.ExamUnitId = vExamBookingUnit.ExamUnitId AndAlso vBookingLine.ExamUnitProductId > 0 Then
              'Create the transaction for additional exam products
              Dim vExamTransaction As New ExamBookingTransaction(mvEnv)
              vExamTransaction.SetTransactionInfo(vExamBookingUnit.ExamBookingUnitId, vBookingLine.BatchNumber, vBookingLine.TransactionNumber, vBookingLine.TransactionLineNumber)
            End If
          Next
        End If
        If vExamBookingUnit.ExamScheduleId > 0 Then
          Dim vSql As New SQLStatement(mvEnv.Connection,
                                       "workstream_id",
                                       "workstream_links",
                                       New CDBFields({New CDBField("exam_schedule_id",
                                                                  CDBField.FieldTypes.cftInteger,
                                                                  vExamBookingUnit.ExamScheduleId.ToString)}))
          Using vWorkstreams As DataTable = vSql.GetDataTable
            If vWorkstreams IsNot Nothing Then
              For Each vWorkstreamId As Integer In (From vRow As DataRow In vWorkstreams.AsEnumerable
                                                    Select vRow.Field(Of Integer)("workstream_id"))
                Dim vWorkstream As New Workstream(mvEnv)
                vWorkstream.InitWithPrimaryKey(New CDBFields({New CDBField("workstream_id",
                                                                           CDBField.FieldTypes.cftInteger,
                                                                           vWorkstreamId.ToString)}))
                vWorkstream.AddExamBookingUnit(vExamBookingUnit)
              Next vWorkstreamId
            End If
          End Using
        End If
      Next

      If vTrans Then mvEnv.Connection.CommitTransaction()

      Return vExamBooking
    End Function

    Public Function GetTopLevelExamUnitId() As Integer
      Dim vResult As Integer = Me.ExamUnitId
      Dim vExamUnitId As Integer = GetParentExamUnitId()

      If vExamUnitId > 0 Then
        Dim vExamUnit As New ExamUnit(mvEnv)
        vExamUnit.Init(vExamUnitId)
        vExamUnitId = vExamUnit.GetTopLevelExamUnitId()
        If vExamUnitId > 0 Then
          vResult = vExamUnitId
        End If
      End If

      Return vResult
    End Function

    Public Function GetParentExamUnitId() As Integer
      Dim vExamUnit As New ExamUnit(mvEnv)
      Dim vExamUnitLink As New ExamUnitLink(mvEnv)
      Dim vParams As New CDBFields()
      vParams.Add("exam_unit_id_2", Me.ExamUnitId)
      vExamUnitLink.InitWithPrimaryKey(vParams)
      If vExamUnitLink.Existing Then
        Return vExamUnitLink.ExamUnitId1
      Else
        Return 0
      End If
    End Function

    Private mvExamUnitLink As ExamUnitLink
    Friend Property UnitLink As ExamUnitLink
      Get
        If mvExamUnitLink Is Nothing Then
          mvExamUnitLink = New ExamUnitLink(mvEnv)
          mvExamUnitLink.Init()
        End If
        Return mvExamUnitLink
      End Get
      Set(ByVal pValue As ExamUnitLink)
        If pValue IsNot Nothing AndAlso pValue.ExamUnitId2 = ExamUnitId Then
          mvExamUnitLink = pValue
        Else
          mvExamUnitLink = New ExamUnitLink(mvEnv)
          mvExamUnitLink.Init()
        End If
      End Set
    End Property

    Protected Overrides Function GetFirstDeleteCheckItem() As DeleteCheckItem
      Dim vRtn As DeleteCheckItem = MyBase.GetFirstDeleteCheckItem()

      If vRtn Is Nothing Then
        'Only allow the cascade delete if there are no shared units (i.e. there's only one exam unit link)
        Dim vWhereFields As New CDBFields
        Dim vPrimaryKey As ClassField = mvClassFields.GetUniquePrimaryKey
        vWhereFields.Add("exam_unit_id_2", vPrimaryKey.FieldType, vPrimaryKey.Value)
        If mvEnv.Connection.GetCount("exam_unit_links", vWhereFields) > 1 Then 'There should only be one exam unit link, i.e. the unit cannot be shared.
          vRtn = New DeleteCheckItem("exam_unit_links", "exam_unit_id_2", "a Shared Unit")
        End If
      End If
      Return vRtn
    End Function


    Private Class ExamUnitMapping
      Private mvExamUnitId As Integer
      Private mvExamBaseUnitId As Integer
      Private mvExamQuestion As Boolean
      Private mvExamSchedule As Boolean
      Private mvExamUnitLinkId As Integer
      Private mvBaseUnitLinkId As Integer

      Public Sub New(ByVal pExamUnitId As Integer, ByVal pBaseUnitId As Integer, ByVal pExamUnitLinkId As Integer, ByVal pBaseUnitLinkId As Integer, ByVal pQuestion As Boolean, ByVal pSchedule As Boolean)
        mvExamUnitId = pExamUnitId
        mvExamUnitLinkId = pExamUnitLinkId
        mvExamBaseUnitId = pBaseUnitId
        mvBaseUnitLinkId = pBaseUnitLinkId
        mvExamQuestion = pQuestion
        mvExamSchedule = pSchedule
      End Sub

      Public ReadOnly Property ExamUnitId As Integer
        Get
          Return mvExamUnitId
        End Get
      End Property
      Public ReadOnly Property ExamUnitLinkId As Integer
        Get
          Return mvExamUnitLinkId
        End Get
      End Property
      Public ReadOnly Property BaseUnitId As Integer
        Get
          If mvExamBaseUnitId > 0 Then
            Return mvExamBaseUnitId         'Session based
          Else
            Return mvExamUnitId             'Non Session based
          End If
        End Get
      End Property
      Public ReadOnly Property BaseUnitLinkId As Integer
        Get
          If mvBaseUnitLinkId > 0 Then
            Return mvBaseUnitLinkId         'Session based
          Else
            Return mvExamUnitLinkId             'Non Session based
          End If
        End Get
      End Property

      Public ReadOnly Property ExamQuestion As Boolean
        Get
          Return mvExamQuestion
        End Get
      End Property
      Public ReadOnly Property ExamSchedule As Boolean
        Get
          Return mvExamSchedule
        End Get
      End Property
    End Class

#End Region

#Region "Accreditation Status"

    Public Shared Function IsUnitAccredited(ByVal pEnv As CDBEnvironment, ByVal pDT As CDBDataTable, ByVal pTrader As Boolean) As CDBDataTable
      Return IsUnitAccredited(pEnv, pDT, pTrader, False)
    End Function

    ''' <summary>
    ''' Check If the units are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Environment class </param>
    ''' <param name="pDT">Data table with the exam units</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsUnitAccredited(ByVal pEnv As CDBEnvironment, ByVal pDT As CDBDataTable, ByVal pTrader As Boolean, ByVal pBooking As Boolean) As CDBDataTable
      Dim vDataTable As New CDBDataTable
      If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamUnitAccreditation).Length > 0 AndAlso _
        pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamUnitAccreditation) = "Y" Then
        If pDT IsNot Nothing AndAlso (pDT.Columns.ContainsKey("accreditation_status") Or pDT.Columns.ContainsKey("AccreditationStatus")) Then
          Dim vLinkIdColumn As String = ""
          If pDT.Columns.ContainsKey("exam_unit_link_id") Then
            vLinkIdColumn = "exam_unit_link_id"
          Else
            vLinkIdColumn = "ExamUnitLinkId"
          End If
          If pBooking Then
            Dim vDTable As CDBDataTable = pDT
            Dim vRowNumber As Integer = pDT.Rows.Count - 1
            Do
              If Not CheckUnitAccreditationStatus(pEnv, vDTable.Rows(vRowNumber).IntegerItem(vLinkIdColumn), pTrader) Then
                Dim vString = RemoveChildLinkUnits(pEnv, vDTable.Rows(vRowNumber).IntegerItem(vLinkIdColumn), pDT)
              End If
              If vRowNumber >= pDT.Rows.Count Then vRowNumber = pDT.Rows.Count 'If we have removed rows, vRowNumber may now be greater than the number of rows in the table
              vRowNumber -= 1
            Loop While vRowNumber >= 0
          Else
            For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1
              If Not CheckUnitAccreditationStatus(pEnv, pDT.Rows(vRowNumber).IntegerItem(vLinkIdColumn), pTrader) Then
                pDT.Rows.RemoveAt(vRowNumber)
              End If
            Next
          End If
        End If
      Else
        Return pDT
      End If
      Return pDT
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function RemoveChildLinkUnits(ByVal pEnv As CDBEnvironment, ByVal pExamUnitLink As Integer, ByVal pDt As CDBDataTable) As CDBDataTable
      Dim vResult As New StringBuilder(pExamUnitLink)
      Dim vSqlString As String = "WITH WGS (exam_unit_link_id) AS ( select exam_unit_link_id from exam_unit_links where exam_unit_link_id =" & "'" & pExamUnitLink & "'"
      vSqlString = vSqlString + " union all select eul.exam_unit_link_id from exam_unit_links eul inner join wgs on eul.parent_unit_link_id =  wgs.exam_unit_link_id) select * from wgs;"

      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQLDONOTUSE(pEnv, vSqlString)

      For Each vDataRow As CDBDataRow In vDataTable.Rows
        For vRowNumber As Integer = pDt.Rows.Count - 1 To 0 Step -1
          If CInt(pDt.Rows(vRowNumber).Item("ExamUnitLinkId")) = CInt(vDataRow.Item("exam_unit_link_id")) Then
            pDt.Rows.RemoveAt(vRowNumber)
            Exit For
          End If
        Next
      Next

      Return pDt
      'Return vDataTable
    End Function

    ''' <summary>
    ''' Check if the Units are accredited, then only add them to the lookup when called from trader or results entry screens
    ''' </summary>
    ''' <param name="pEnv">Environment class </param>
    ''' <param name="pUnitLinkId">Unit Link Id to validate</param>
    ''' <param name="pTrader">Trader flag</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CheckUnitAccreditationStatus(ByVal pEnv As CDBEnvironment, ByVal pUnitLinkId As Integer, ByVal pTrader As Boolean) As Boolean
      Dim vAnsiJoin As New AnsiJoins
      Dim vFields As String = "eul.accreditation_status,allow_registration,ignore_accreditation_validity,allow_result_entry,eul.accreditation_valid_from,eul.accreditation_valid_to"
      Dim vWhereClause As New CDBFields
      Dim vResult As Boolean = False

      vWhereClause.Add("eul.exam_unit_link_id", pUnitLinkId)

      If pTrader Then
        vWhereClause.Add("acs.allow_registration", "Y")
      Else
        vWhereClause.Add("acs.allow_result_entry", "Y")
      End If

      vAnsiJoin.Add("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")
      vAnsiJoin.Add("exam_accreditation_statuses acs", "eul.accreditation_status", "acs.accreditation_status")

      Dim vSql As New SQLStatement(pEnv.Connection, vFields, "exam_units eu", vWhereClause, "", vAnsiJoin)
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQL(pEnv, vSql)


      If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then

        Dim vIgnoreBooking As Boolean = If(vDataTable.Rows(0).Item("ignore_accreditation_validity").Length > 0 AndAlso vDataTable.Rows(0).Item("ignore_accreditation_validity") = "Y", True, False)

        Dim vValidFrom As String = If(vDataTable.Columns.ContainsKey("accreditation_valid_from"), vDataTable.Rows(0).Item("accreditation_valid_from"), vDataTable.Rows(0).Item("AccreditationValidFrom"))
        Dim vValidTo As String = If(vDataTable.Columns.ContainsKey("accreditation_valid_to"), vDataTable.Rows(0).Item("accreditation_valid_to"), vDataTable.Rows(0).Item("AccreditationValidTo"))

        'CheckCentreAccreditationStatus If booking is allowed for centers, this should only be checked for trade application
        If pTrader Then
          If IsAccreditationValid(vValidFrom, vValidTo, vIgnoreBooking) Then vResult = True
        Else
          Return True
        End If
      End If
      Return vResult
    End Function
    ''' <summary>
    ''' Validate the Date range specified for Accreditation
    ''' </summary>
    ''' <param name="pValidFrom">Accreditation Valid from</param>
    ''' <param name="pValidTo">Accreditation valid To</param>
    ''' <param name="pIgnoreValidity"> Ignore Validity</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function IsAccreditationValid(ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pIgnoreValidity As Boolean) As Boolean
      Dim vResult As Boolean = False

      'Check if the dates are valid
      If Not pIgnoreValidity Then
        If pValidFrom.Length = 0 AndAlso pValidTo.Length = 0 Then
          vResult = False
        ElseIf pValidFrom.Length > 0 AndAlso CDate(pValidFrom) > Date.Today Then
          vResult = False 'future
        ElseIf pValidTo.Length > 0 AndAlso CDate(pValidTo) < Date.Today Then
          vResult = False 'past 
        Else
          vResult = True
        End If
      Else
        vResult = True
      End If
      Return vResult
    End Function
#End Region

  End Class
End Namespace
