Imports System.Reflection

Namespace Access

  Public Class WorkstreamDataSelection
    Inherits DataSelection

    Public Shadows Enum DataSelectionTypes As Integer
      dstNone = 6000
      dstWorkstreamData
      dstWorkstreamDetails
      dstWorkstreamExamBookingUnits
      dstWorkstreamActions
      dstWorkstreamExamSchedule
      dstWorkstreamDocuments
      dstWorkstreamActionsFromTemplates
      dstWorkstreamCategories
      dstWorkstreamSelectionPages
    End Enum

    Private mvWorkstreamSelectionType As DataSelectionTypes
    Private mvMethodDictionary As Dictionary(Of DataSelectionTypes, MethodInfo)

    Private Property SelectionType As DataSelectionTypes
      Get
        Return mvWorkstreamSelectionType
      End Get
      Set(value As DataSelectionTypes)
        mvWorkstreamSelectionType = value
      End Set
    End Property

    Private Property WorkstreamGroup As String
      Get
        Return mvGroupCode
      End Get
      Set(value As String)
        mvGroupCode = value
      End Set
    End Property

    Private Property MethodDictionary As Dictionary(Of DataSelectionTypes, MethodInfo)
      Get
        Return mvMethodDictionary
      End Get
      Set(value As Dictionary(Of DataSelectionTypes, MethodInfo))
        value = mvMethodDictionary
      End Set
    End Property

    Public Sub New(ByVal pGroupCode As String, ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, pGroupCode)
      LoadMethodDictionary()
    End Sub

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, "")
      LoadMethodDictionary()
    End Sub

    Private Sub LoadMethodDictionary()
      mvMethodDictionary = New Dictionary(Of DataSelectionTypes, MethodInfo)
      Dim vType As Type = Me.GetType()
      Dim vSelectionMethods As MethodInfo() = vType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
      For Each vMethodInfo As MethodInfo In vSelectionMethods
        Dim parms As ParameterInfo() = vMethodInfo.GetParameters()
        'Must only have 1 parameter, the datatable
        If parms IsNot Nothing AndAlso parms.Length = 1 Then
          If parms(0).ParameterType = GetType(CDBDataTable) Then
            Dim atts() As Object = vMethodInfo.GetCustomAttributes(GetType(EnumEquivalentAttribute), False)
            If atts IsNot Nothing AndAlso atts.Length > 0 Then
              For Each att As EnumEquivalentAttribute In atts
                If att.EquivalentValue IsNot Nothing AndAlso TypeOf att.EquivalentValue Is DataSelectionTypes Then
                  mvMethodDictionary.Add(DirectCast(att.EquivalentValue, DataSelectionTypes), vMethodInfo)
                End If
              Next
            End If
          End If
        End If
      Next
    End Sub

    Private Sub ExecuteSelectionMethod(pMethod As MethodInfo, pDataTable As CDBDataTable)
      If pMethod IsNot Nothing Then
        Try
          pMethod.Invoke(Me, New Object() {pDataTable})
        Catch vEx As Exception
          If vEx.GetType Is GetType(TargetInvocationException) Then
            vEx = vEx.InnerException
            Throw
          End If
        End Try
      End If
    End Sub

    Public Overrides Function DataTable() As CDBDataTable
      If mvParameters Is Nothing Then mvParameters = New CDBParameters
      Dim vDataTable As New CDBDataTable
      If mvResultColumns.Length > 0 Then vDataTable.AddColumnsFromList(mvResultColumns)
      If MethodDictionary.ContainsKey(SelectionType) Then
        Dim vMethod As MethodInfo = MethodDictionary(SelectionType)
        If vMethod IsNot Nothing Then
          ExecuteSelectionMethod(vMethod, vDataTable)
        End If
      End If
      Return vDataTable
    End Function

    Private Overloads Sub Init(ByVal pEnv As CDBEnvironment, ByVal pType As WorkstreamDataSelection.DataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsage As DataSelectionUsages, ByVal pGroup As String)
      mvEnv = pEnv
      SelectionType = pType
      mvParameters = pParams
      mvUsage = pUsage
      WorkstreamGroup = pGroup
      mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
      mvDisplayListItems = Nothing
      mvDataSelectionListType = pListType

      Dim vPrimaryList As Boolean = False

      Select Case SelectionType
        Case DataSelectionTypes.dstWorkstreamData
          mvResultColumns = "WorkstreamId,WorkstreamDesc,WorkstreamGroup,WorkstreamGroupDesc,StartDate,EndDate,WorkstreamGroupOutcome,WorkstreamGroupOutcomeDesc,OutcomeDate,Notes,AmendedBy,AmendedOn"
          mvSelectColumns = "WorkstreamId,WorkstreamDesc,WorkstreamGroup,WorkstreamGroupDescription,StartDate,EndDate,WorkstreamGroupOutcome,WorkstreamGroupOutcomeDesc,OutcomeDate,Notes,AmendedBy,AmendedOn"
          mvHeadings = "ID,Description,Group Code,Group Description,Start Date,End Date,Outcome Code,Outcome Description,Outcome Date,Notes,Amended By,Amended On"
          mvRequiredItems = "WorkstreamId,WorkstreamGroup"
          mvCode = "WDA"
        Case DataSelectionTypes.dstWorkstreamDetails
          mvResultColumns = "WorkstreamId,WorkstreamDesc,WorkstreamGroup,WorkstreamGroupDesc,StartDate,EndDate,WorkstreamGroupOutcome,WorkstreamGroupOutcomeDesc,OutcomeDate,Notes,WorkstreamContactNumber,OwnershipGroup,AccessLevel,AmendedBy,AmendedOn"
          mvSelectColumns = "WorkstreamId,WorkstreamDesc,WorkstreamGroup,WorkstreamGroupDesc,StartDate,EndDate,WorkstreamGroupOutcome,WorkstreamGroupOutcomeDesc,OutcomeDate,Notes,WorkstreamContactNumber,AmendedBy,AmendedOn"
          mvHeadings = "ID,Description,Group Code,Group Description,Start Date,End Date,Outcome Code,Outcome Description,Outcome Date,Notes,ContactNumber,Amended By,Amended On"
          mvRequiredItems = "WorkstreamId,WorkstreamGroup,WorkstreamGroupDesc,WorkstreamDesc,OwnershipGroup,AccessLevel"
          mvCode = "WMNT"
        Case DataSelectionTypes.dstWorkstreamExamBookingUnits
          'mvResultsColumns
          mvResultColumns = {"ExamBookingUnitId",
                             "Breadcrumb",
                             "ExamUnitCode",
                             "ExamUnitDescription",
                             "ExamCentreCode",
                             "ExamCentreDescription",
                             "ExamSessionCode",
                             "ExamSessionDescription",
                             "ContactNumber",
                             "LabelName",
                             "ExamCandidateNumber",
                             "AttemptNumber",
                             "ExamStudentUnitStatus",
                             "OriginalMark",
                             "ModeratedMark",
                             "TotalMark",
                             "RawMark",
                             "OriginalGrade",
                             "ModeratedGrade",
                             "TotalGrade",
                             "OriginalResult",
                             "ModeratedResult",
                             "TotalResult",
                             "EntryDate",
                             "ExpiryDate",
                             "DoneDate",
                             "DeskNumber",
                             "BatchNumber",
                             "TransactionNumber",
                             "LineNumber",
                             "ActivityGroup",
                             "CancelledBy",
                             "CancelledOn",
                             "CancellationReason",
                              "CreatedBy",
                              "CreatedOn",
                              "AmendedBy",
                              "AmendedOn"}.AsCommaSeperated
          'mvSelectColumns
          mvSelectColumns = {"ExamBookingUnitId",
                             "Breadcrumb",
                             "ExamUnitCode",
                             "ExamUnitDescription",
                             "ExamCentreCode",
                             "ExamCentreDescription",
                             "ExamSessionCode",
                             "ExamSessionDescription",
                             "ContactNumber",
                             "LabelName",
                             "ExamCandidateNumber",
                             "DetailItems",
                             "AttemptNumber",
                             "OriginalMark",
                             "OriginalGrade",
                             "OriginalResult",
                             "EntryDate",
                             "BatchNumber",
                             "CancelledOn",
                             "DeskNumber",
                             "CreatedBy",
                             "AmendedBy",
                             "NewColumn",
                             "ExamStudentUnitStatus",
                             "ModeratedMark",
                             "ModeratedGrade",
                             "ModeratedResult",
                             "ExpiryDate",
                             "TransactionNumber",
                             "CancelledBy",
                             "ActivityGroup",
                             "CreatedOn",
                             "AmendedOn",
                             "NewColumn2",
                             "RawMark",
                             "TotalMark",
                             "TotalGrade",
                             "TotalResult",
                             "DoneDate",
                             "LineNumber",
                             "CancellationReason"}.AsCommaSeperated
          'mvHeadings
          mvHeadings = {"Id",
                        "Breadcrumb",
                        "Unit Code",
                        "Unit Description",
                        "Centre Code",
                        "Centre Description",
                        "Session Code",
                        "Session Description",
                        "Contact Number",
                        "Name",
                        "Candidate Number",
                        "",
                        "Attempt",
                        "Mark",
                        "Grade",
                        "Result",
                        "Entry Date",
                        "Batch Number",
                        "Cancelled",
                        "Desk Number",
                        "Created By",
                        "Amended By",
                        "",
                        "Status",
                        "Moderated",
                        "Moderated",
                        "Moderated",
                        "Expiry Date",
                        "Transaction",
                        "By",
                        "Activity Group",
                        "Created On",
                        "Amended On",
                        "",
                        "Raw Mark",
                        "Total",
                        "Total",
                        "Total",
                        "Done Date",
                        "Line Number",
                        "Reason"}.AsCommaSeperated
          mvRequiredItems = "ExamBookingUnitId"
          mvCode = "WEBU"

        Case DataSelectionTypes.dstWorkstreamActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,AmendedBy,AmendedOn,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,LinkType,LinkTypeDesc,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText,OutlookId"
          mvSelectColumns = "ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600"
          mvHeadings = DataSelectionText.WorkstreamActions     'Number,Description,Priority,Status,Action Text,,Deadline,Created By,,Scheduled On,Created On,,Completed On
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Workstream Actions"
          mvDisplayTitle = "Workstream Actions"
          mvMaintenanceDesc = "Action"
          mvRequiredItems = "ActionStatus,MasterAction"
          mvCode = "WSAC"
          vPrimaryList = True
        Case DataSelectionTypes.dstWorkstreamExamSchedule
          mvResultColumns = "WorkstreamLinkId,WorkstreamId,ExamScheduleId,ExamSessionId,ExamSessionCode,ExamSessionDescription,ExamSessionYear,ExamSessionMonth,SessionSequenceNumber,ExamCentreId,ExamCentreCode,ExamCentreDescription,ExamUnitId,GradingUnitCode,ExamUnitDescription,ExamScheduleStartDate,ExamScheduleStartTime,ExamScheduleEndTime"
          mvSelectColumns = "ExamSessionCode,ExamSessionDescription,ExamSessionYear,ExamSessionMonth,SessionSequenceNumber,ExamCentreCode,ExamCentreDescription,GradingUnitCode,ExamUnitDescription,ExamScheduleStartDate,ExamScheduleStartTime,ExamScheduleEndTime"
          mvWidths = "800,2000,1500,1500,2000,1200,1600,1200,1200,1600,1200,1200,1600,1400,1400,1400"
          mvHeadings = "Session Code,Session,Session Year,Session Month,Session Sequence,Centre Code,Centre,Exam Unit Code,Exam Unit,Schedule Start Date,Schedule Start Time,Schedule End Time"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Workstream Exam Schedule"
          mvDisplayTitle = "Workstream Exam Schedule"
          mvMaintenanceDesc = "Exam Schedule"
          mvRequiredItems = "WorkstreamLinkId,WorkstreamId,ExamSessionId,ExamCentreId,ExamScheduleId,ExamUnitId"
          mvCode = "WXSC"

        Case DataSelectionTypes.dstWorkstreamDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc,Precis"
          mvResultColumns = mvResultColumns & ",Subject,CallDuration,TotalDuration,SelectionSet"      'Optional Attributes
          mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc,DetailItems,CreatedBy,Source,PackageCode,NewColumn,DepartmentDesc,StandardDocument,NewColumn2,DocumentClassDesc,TheirReference"
          mvHeadings = DataSelectionText.String17987     'Document Number,Dated,In/ Out,Subject,Reference+,Document Type,Topic+,Sub Topic,,Creator,Source,Package,,Department,Standard Document,,Document Class,Their Reference
          mvWidths = "1200,1200,1200,1500,1500,1500,1400,1400,1200,1400,1200,600,1200,1400,1200,1200,1200,1200"
          mvRequiredItems = "Access"
          mvDescription = DataSelectionText.String17988    'Contact Documents                              
          mvCode = "WSDC"
          vPrimaryList = True

        Case DataSelectionTypes.dstWorkstreamActionsFromTemplates
          mvResultColumns = "ActionNumber,MasterAction,ActionLevel,SequenceNumber,ActionDesc,ActionText,ActionPriority,ActionPriorityDesc,ActionStatus,ActionStatusDesc,DocumentClass,"
          mvResultColumns &= "DurationDays,DurationHours,DurationMinutes,Deadline,ScheduledOn,CompletedOn,RepeatCount,ActionTemplateNumber,UseWorkingDays,UseNegativeOffsets,CreatedBy,"
          mvResultColumns &= "CreatedOn,DelayedActivation,ActionerSetting,ManagerSetting,DelayDays,DelayMonths,DeadlineDays,DeadlineMonths,RepeatDays,RepeatMonths,OutlookId,"
          mvResultColumns &= "Department,IsExisting,Delete,AccessRights"
          mvSelectColumns = "Delete,ActionDesc,Deadline,ScheduledOn,CompletedOn,ActionStatus"
          mvWidths = "1200,1200,1200,1200,1200"
          mvHeadings = "Delete?,Description,Deadline,Scheduled,Completed,Status"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvDescription = "Workstream Actions From Templates"
          mvDisplayTitle = "Workstream Actions From Templates"
          mvRequiredItems = mvResultColumns   'This will be used to create Actions so must have all columns.
          mvCode = "WSAT"

        Case DataSelectionTypes.dstWorkstreamCategories
          mvResultColumns = "Activity,ActivityValue,Quantity,ActivityDate,Source,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,WorkstreamId,CategoryId,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,ActivityDate,Quantity,Source,ValidFrom,ValidTo,NoteFlag,Status,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40029     'Activity Desc,Value,Quantity,Source,Valid from,Valid to,Notes?,Amended by,Amended on
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200"
          mvRequiredItems = "CategoryId,WorkstreamId,Activity,ActivityValue,Quantity,ActivityDate,Source,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,WorkstreamId,CategoryId,NoteFlag,Status,StatusOrder"
          mvDescription = "Workstream Categories"
          mvCode = "WXCM"

        Case DataSelectionTypes.dstWorkstreamSelectionPages
          mvSelectColumns = "SelectionHeading,Details,Actions,Categories,Documents," & _
                            "SelectionHeading1,ExamBookingUnits,ExamSchedules"
          mvResultColumns = mvSelectColumns & ",SelectionHeading2"
          mvHeadings = "%Desc,Details,Actions,Activities,Documents," & _
                       "Links,Exam Booking Units,Exam Schedules"
          mvWidths = "300,300,300,300,300,300,300,300"
          mvRequiredItems = String.Empty
          mvDescription = "Workstream Selection Pages"
          mvAvailableUsages = DataSelectionUsages.dsuSmartClient
          mvCode = "WSSP"
      End Select

      Dim vSelectItems() As String = mvSelectColumns.Split(","c)
      Dim vWidths As New StringBuilder
      For vIndex As Integer = 0 To vSelectItems.Length - 1
        If vIndex > 0 Then vWidths.Append(",")
        vWidths.Append("1000")
      Next
      mvWidths = vWidths.ToString

      If vPrimaryList = True AndAlso pListType = DataSelectionListType.dsltEditing AndAlso mvAvailableUsages = DataSelectionUsages.dsuSmartClient Then
        mvResultColumns = mvResultColumns & ",DetailItems,NewColumn,NewColumn2,NewColumn3,Spacer"
      End If

      Select Case pListType
        Case DataSelectionListType.dsltUser
          ReadUserDisplayListItems(mvEnv.User.Department, mvEnv.User.Logname, pGroup, pUsage)
      End Select
    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamDetails)>
    Private Sub GetWorkstreamDetails(ByVal pDataTable As CDBDataTable)
      Dim vFields As String = "workstreams.workstream_id,workstreams.workstream_desc,workstreams.workstream_group,workstream_groups.workstream_group_desc,workstreams.start_date,workstreams.end_date,workstreams.workstream_group_outcome,workstream_group_outcomes.workstream_group_outcome_desc,workstreams.outcome_date,workstreams.notes,workstreams.contact_number,workstreams.ownership_group,COALESCE(record_access.ownership_access_level, 'W') access_level,workstreams.amended_by,workstreams.amended_on"
      Dim vColumnNames As String = "workstream_id,workstream_desc,workstream_group,workstream_group_desc,start_date,end_date,workstream_group_outcome,workstream_group_outcome_desc,outcome_date,notes,contact_number,ownership_group,access_level,amended_by,amended_on"
      Dim vAnsiJoins As New AnsiJoins()
      Dim vLoggedInUser As String = String.Format("'{0}'", mvEnv.User.Logname)
      vAnsiJoins.Add("workstream_groups", "workstream_groups.workstream_group", "workstreams.workstream_group")
      vAnsiJoins.AddLeftOuterJoin("workstream_group_outcomes", "workstream_group_outcomes.workstream_group", "workstreams.workstream_group", "workstream_group_outcomes.workstream_group_outcome", "workstreams.workstream_group_outcome")
      vAnsiJoins.AddLeftOuterJoin("ownership_group_users record_access", "record_access.ownership_group", "workstreams.ownership_group", "record_access.logname", vLoggedInUser)

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("record_access.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("record_access.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)


      If mvParameters.Exists("WorkstreamId") Then
        If WorkstreamGroup.Length > 0 Then vWhereFields.Add("workstreams.workstream_group", WorkstreamGroup)
        vWhereFields.Add("workstreams.workstream_id", mvParameters("WorkstreamId").Value)
      ElseIf WorkstreamGroup.Length > 0 Then
        vWhereFields.Add("workstreams.workstream_group", WorkstreamGroup)
      End If

      If mvParameters.Exists("ExamBookingUnitId") OrElse
         mvParameters.Exists("ExamScheduleId") Then
        vAnsiJoins.Add("workstream_links", "workstream_links.workstream_id", "workstreams.workstream_id")
      End If

      If mvParameters.Exists("ExamBookingUnitId") Then
        vWhereFields.Add("workstream_links.exam_booking_unit_id", mvParameters("ExamBookingUnitId").Value)
      End If

      If mvParameters.Exists("ExamScheduleId") Then
        vWhereFields.Add("workstream_links.exam_schedule_id", mvParameters("ExamScheduleId").Value)
      End If

      If mvParameters.Exists("Keys") Then
        vWhereFields.Add("workstreams.workstream_id", mvParameters("Keys").Value, CDBField.FieldWhereOperators.fwoIn)
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "workstreams", vWhereFields, "workstreams.start_date desc, workstreams.workstream_id desc", vAnsiJoins)
      If mvParameters.Exists("Top") Then
        vSQL.MaxRows = mvParameters("Top").IntegerValue
      End If
      pDataTable.FillFromSQL(mvEnv, vSQL, vColumnNames, "", True)

      If pDataTable.Columns.ContainsKey("AccessLevel") Then
        Dim vAllowedCols As New List(Of String)
        If mvRequiredItems IsNot Nothing AndAlso mvRequiredItems.Length > 0 Then
          vAllowedCols.AddRange(mvRequiredItems.Split(","c))
        End If
        ApplyOwnershipAccessSecurity(pDataTable, "AccessLevel", vAllowedCols)
      End If

    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamExamBookingUnits)>
    Private Sub GetWorkstreamExamBookingUnits(ByVal pDataTable As CDBDataTable)
      Dim vSql As New StringBuilder
      vSql.AppendLine("With CTE(id,")
      vSql.AppendLine("      parent_id, ")
      vSql.AppendLine("      breadcrumb) ")
      vSql.AppendLine("AS   (SELECT exam_unit_link_id AS id, ")
      vSql.AppendLine("             parent_unit_link_id AS parent_id, ")
      vSql.AppendLine("             CAST(eu.exam_unit_code AS varchar(1024)) AS breadcrumb ")
      vSql.AppendLine("      FROM   exam_unit_links eul ")
      vSql.AppendLine("             INNER JOIN exam_units eu ")
      vSql.AppendLine("                        ON eu.exam_unit_id = eul.exam_unit_id_2 ")
      vSql.AppendLine("      WHERE parent_unit_link_id = 0")
      vSql.AppendLine("      UNION ALL")
      vSql.AppendLine("      SELECT eul.exam_unit_link_id AS id, ")
      vSql.AppendLine("             eul.parent_unit_link_id as parent_id, ")
      vSql.AppendLine("             CAST(CTE.breadcrumb + ' / ' + eu.exam_unit_code AS varchar(1024)) AS breadcrumb ")
      vSql.AppendLine("      FROM   exam_unit_links eul ")
      vSql.AppendLine("             INNER JOIN exam_units eu ")
      vSql.AppendLine("                        ON eu.exam_unit_id = eul.exam_unit_id_2 ")
      vSql.AppendLine("             INNER JOIN CTE ")
      vSql.AppendLine("                        ON CTE.id = eul.parent_unit_link_id) ")
      vSql.AppendLine("SELECT ebu.exam_booking_unit_id, ")
      vSql.AppendLine("       CTE.breadcrumb, ")
      vSql.AppendLine("       eu.exam_unit_code, ")
      vSql.AppendLine("       Coalesce(ecu.local_name, eu.exam_unit_description) AS exam_unit_description, ")
      vSql.AppendLine("       ec.exam_centre_code, ")
      vSql.AppendLine("       ec.exam_centre_description, ")
      vSql.AppendLine("       es.exam_session_code, ")
      vSql.AppendLine("       es.exam_session_description, ")
      vSql.AppendLine("       c.contact_number, ")
      vSql.AppendLine("       c.label_name, ")
      vSql.AppendLine("       ebu.exam_candidate_number, ")
      vSql.AppendLine("       ebu.attempt_number, ")
      vSql.AppendLine("       ebu.exam_student_unit_status, ")
      vSql.AppendLine("       ebu.original_mark, ")
      vSql.AppendLine("       ebu.moderated_mark, ")
      vSql.AppendLine("       ebu.total_mark, ")
      vSql.AppendLine("       ebu.raw_mark, ")
      vSql.AppendLine("       ebu.original_grade, ")
      vSql.AppendLine("       ebu.moderated_grade, ")
      vSql.AppendLine("       ebu.total_grade, ")
      vSql.AppendLine("       ebu.original_result, ")
      vSql.AppendLine("       ebu.moderated_result, ")
      vSql.AppendLine("       ebu.total_result, ")
      vSql.AppendLine("       ebu.entry_date, ")
      vSql.AppendLine("       ebu.expiry_date, ")
      vSql.AppendLine("       ebu.done_date, ")
      vSql.AppendLine("       ebu.desk_number, ")
      vSql.AppendLine("       ebu.batch_number, ")
      vSql.AppendLine("       ebu.transaction_number, ")
      vSql.AppendLine("       ebu.line_number, ")
      vSql.AppendLine("       eu.activity_group, ")
      vSql.AppendLine("       eb.cancelled_by, ")
      vSql.AppendLine("       eb.cancelled_on, ")
      vSql.AppendLine("       cr.cancellation_reason_desc, ")
      vSql.AppendLine("       ebu.created_by, ")
      vSql.AppendLine("       ebu.created_on, ")
      vSql.AppendLine("       ebu.amended_by, ")
      vSql.AppendLine("       ebu.amended_on ")
      vSql.AppendLine("FROM   workstream_links wl ")
      vSql.AppendLine("       INNER JOIN exam_booking_units ebu ")
      vSql.AppendLine("                  ON ebu.exam_booking_unit_id = wl.exam_booking_unit_id ")
      vSql.AppendLine("       INNER JOIN CTE ")
      vSql.AppendLine("                  ON CTE.id = ebu.exam_unit_link_id ")
      vSql.AppendLine("       INNER JOIN exam_unit_links eul ")
      vSql.AppendLine("                  ON eul.exam_unit_link_id = ebu.exam_unit_link_id ")
      vSql.AppendLine("       INNER JOIN exam_units eu ")
      vSql.AppendLine("                  ON eu.exam_unit_id = eul.exam_unit_id_2 ")
      vSql.AppendLine("	      INNER JOIN contacts c ")
      vSql.AppendLine("	                 ON c.contact_number = ebu.contact_number ")
      vSql.AppendLine("	      LEFT OUTER JOIN exam_bookings eb ")
      vSql.AppendLine("	                 ON eb.exam_booking_id = ebu.exam_booking_id ")
      vSql.AppendLine("	      LEFT OUTER JOIN cancellation_reasons cr ")
      vSql.AppendLine("	                 ON cr.cancellation_reason = eb.cancellation_reason ")
      vSql.AppendLine("	      LEFT OUTER JOIN exam_centres ec ")
      vSql.AppendLine("	                      ON ec.exam_centre_id = eb.exam_centre_id ")
      vSql.AppendLine("       LEFT OUTER JOIN exam_centre_units ecu ")
      vSql.AppendLine("	                      ON ecu.exam_centre_id = eb.exam_centre_id ")
      vSql.AppendLine("		                    AND ecu.exam_unit_link_id = ebu.exam_unit_link_id ")
      vSql.AppendLine("	      LEFT OUTER JOIN exam_sessions es ")
      vSql.AppendLine("                       ON es.exam_session_id = eb.exam_session_id")
      vSql.AppendFormat("WHERE  wl.workstream_id = {0}", mvParameters("WorkstreamId").Value)
      pDataTable.FillFromSQL(mvEnv, New SQLStatement(mvEnv.Connection, vSql.ToString))
    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamActions)>
    Private Sub GetWorkstreamActions(ByVal pDataTable As CDBDataTable)
      If mvParameters.ContainsKey("WorkstreamId") = False Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterNotFound, "WorkstreamId")
      ElseIf mvParameters.ParameterExists("WorkstreamId").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "WorkstreamId")
      End If
      GetActions(pDataTable)
    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamDocuments)>
    Private Sub GetWorkstreamDocuments(ByVal pDataTable As CDBDataTable)
      If mvParameters.ContainsKey("WorkstreamId") = False Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterNotFound, "WorkstreamId")
      ElseIf mvParameters.ParameterExists("WorkstreamId").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "WorkstreamId")
      End If
      GetDocumentLinks(pDataTable)
    End Sub

    Private Sub GetDocumentLinks(ByVal pDataTable As CDBDataTable)
      'Exams Changes
      Dim vWorkstreamId As Boolean = mvParameters.Exists("WorkstreamId")


      Dim vAttrs As String = "dated,cl.communications_log_number,cl.package,label_name,c.contact_number,document_type_desc,created_by,department_desc,our_reference,direction,their_reference,cl.document_type,cl.document_class,document_class_desc,standard_document,cl.source,recipient,forwarded,archiver,completed,cls.topic,topic_desc,cls.sub_topic,sub_topic_desc,creator_header,department_header,public_header,d.department,creator_header AS access_level,standard_document AS standard_document_desc"
      vAttrs &= ",precis,subject,call_duration,total_duration,selection_set"

      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing) Then vAttrs = vAttrs.Replace(",selection_set", ",")

      Dim vTables As New StringBuilder
      vTables.Append("document_log_links dl, communications_log cl, communications_log_subjects cls, contacts c, document_types dt, document_classes dc, departments d, topics t, sub_topics st")

      Dim vWhereFields As New CDBFields()

      vWhereFields.Add("dl.workstream_id", CDBField.FieldTypes.cftInteger, mvParameters("WorkstreamId").Value)
      vWhereFields.AddJoin("cl.communications_log_number#2", "cls.communications_log_number")
      vWhereFields.Add("primary", "Y").SpecialColumn = True
      vWhereFields.AddJoin("c.contact_number", "cl.contact_number")
      vWhereFields.AddJoin("dt.document_type", "cl.document_type")
      vWhereFields.AddJoin("dc.document_class", "cl.document_class")
      vWhereFields.AddJoin("d.department", "cl.department")
      vWhereFields.AddJoin("t.topic", "cls.topic")
      vWhereFields.AddJoin("st.topic", "t.topic")
      vWhereFields.AddJoin("st.sub_topic", "cls.sub_topic")
      If vWorkstreamId Then vWhereFields.AddJoin("cl.communications_log_number", "dl.communications_log_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, vAttrs, vTables.ToString, vWhereFields, "dated DESC, cl.communications_log_number DESC")
      'If Not (mvDataSelectionListType = DataSelectionTypes.dstContactDocuments Or mvDataSelectionListType = DataSelectionTypes.dstDistinctDocuments Or mvDataSelectionListType = DataSelectionTypes.dstDistinctExternalDocuments) Then vSQLStatement.Distinct = True
      '      If mvParameters.Exists("NumberOfRows") Then vSQLStatement.MaxRows = mvParameters("NumberOfRows").LongValue + 1
      If mvParameters.Exists("NumberOfRows") Then
        vSqlStatement.RecordSetOptions = CDBConnection.RecordSetOptions.NoDataTable
        pDataTable.MaximumRows = mvParameters("NumberOfRows").IntegerValue + 1
        pDataTable.CheckAccess = True
      End If
      'If mvExamSelectionType = DataSelectionTypes.dstContactDocuments Then
      vAttrs = Replace(vAttrs, "cl.communications_log_number", "DISTINCT_DOCUMENT_NUMBER")
      'End If ' Or mvExamSelectionType = DataSelectionTypes.dstDistinctDocuments Or mvExamSelectionType = DataSelectionTypes.dstDistinctExternalDocuments Then vAttrs = Replace(vAttrs, "cl.communications_log_number", "DISTINCT_DOCUMENT_NUMBER")
      vAttrs = vAttrs.Replace("creator_header AS access_level", "ACCESS")
      vAttrs = vAttrs.Replace("standard_document AS ", "")
      pDataTable.FillFromSQL(mvEnv, vSqlStatement, vAttrs)
      'Set document access is now done in the DataTable pDataTable.SetDocumentAccess()
      GetDescriptions(pDataTable, "StandardDocument")
      For Each vRow As CDBDataRow In pDataTable.Rows
        If vRow.Item("Direction") = "I" Then
          vRow.Item("Direction") = DataSelectionText.String18677    'In
        Else
          vRow.Item("Direction") = DataSelectionText.String18678    'Out
        End If
        If vRow.Item("CallDuration").Length > 0 Then
          vRow.Item("CallDuration") = vRow.Item("CallDuration").Insert(2, ":").Insert(5, ":")
          vRow.Item("TotalDuration") = vRow.Item("TotalDuration").Insert(2, ":").Insert(5, ":")
        End If
      Next

    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamExamSchedule)>
    Private Sub GetWorkstreamExamSchedule(ByVal pDataTable As CDBDataTable)
      If mvParameters.ContainsKey("WorkstreamId") = False Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterNotFound, "WorkstreamId")
      ElseIf mvParameters.ParameterExists("WorkstreamId").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "WorkstreamId")
      Else
        Dim vFields As String = {"workstream_links.workstream_link_id", "workstream_links.workstream_id", "workstream_links.exam_schedule_id", "exam_sessions.exam_session_id",
                                 "exam_sessions.exam_session_code", "exam_sessions.exam_session_description", "exam_sessions.exam_session_year", "exam_sessions.exam_session_month", "exam_sessions.sequence_number",
                                 "exam_centres.exam_centre_id", "exam_centres.exam_centre_code", "exam_centres.exam_centre_description",
                                 "exam_units.exam_unit_id", "exam_units.exam_unit_code", "exam_units.exam_unit_description",
                                 "exam_schedule.start_date", "exam_schedule.start_time", "exam_schedule.end_time"
                                }.AsCommaSeperated
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("workstreams", "workstreams.workstream_id", "workstream_links.workstream_id")
        vAnsiJoins.Add("exam_schedule", "exam_schedule.exam_schedule_id", "workstream_links.exam_schedule_id")
        vAnsiJoins.Add("exam_units", "exam_units.exam_unit_id", "exam_schedule.exam_unit_id")
        vAnsiJoins.Add("exam_sessions", "exam_sessions.exam_session_id", "exam_units.exam_session_id")
        vAnsiJoins.Add("exam_centres", "exam_centres.exam_centre_id", "exam_schedule.exam_centre_id")

        Dim vWhereFields As New CDBFields()
        vWhereFields.Add("workstreams.workstream_group", WorkstreamGroup)
        vWhereFields.Add("workstream_links.workstream_id", mvParameters("WorkstreamId").Value)

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "workstream_links", vWhereFields, "exam_schedule.start_date, exam_sessions.exam_session_year desc, exam_sessions.exam_session_month desc, exam_sessions.sequence_number desc", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, vFields, "", True)
      End If

    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamActionsFromTemplates)>
    Private Sub GetWorkstreamActionsFromTemplate(ByVal pDataTable As CDBDataTable)
      If mvParameters.ParameterExists("WorkstreamId").IntegerValue = 0 Then
        RaiseError(Utilities.Common.DataAccessErrors.daeParameterValueInvalid, "WorkstreamId")
      End If

      Dim vAttrs As String = "a.action_number, master_action, action_level, sequence_number, action_desc, action_text, a.action_priority, action_priority_desc, "
      vAttrs &= "a.action_status, action_status_desc, a.document_class, duration_days, duration_hours, duration_minutes, deadline, scheduled_on, completed_on, "
      vAttrs &= "repeat_count, action_template_number, use_working_days, use_negative_offsets, a.created_by, a.created_on, delayed_activation, actioner_setting,"
      vAttrs &= "manager_setting, delay_days, delay_months, deadline_days, deadline_months, repeat_days, repeat_months, outlook_id, u.department,"
      vAttrs &= " 'Y' AS existing, 'N' AS ""delete"", '0' AS access_rights"

      Dim vAnsiJoins As New AnsiJoins()
      With vAnsiJoins
        .Add("action_priorities ap", "a.action_priority", "ap.action_priority")
        .Add("action_statuses acs", "a.action_status", "acs.action_status")
        .Add("users u", "a.created_by", "u.logname")
        .Add("document_classes dc", "a.document_class", "dc.document_class")
        .Add("action_links al", "a.action_number", "al.action_number")
      End With

      Dim vWhereFields As New CDBFields(New CDBField("al.workstream_id", mvParameters("WorkstreamId").IntegerValue))
      With vWhereFields
        .Add("action_template_number", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual)
        '.Add("a.completed_on", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoEqual)
        '.Add("a.completed_on#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracket Or CDBField.FieldWhereOperators.fwoOR)
        .Add("a.action_status", CDBField.FieldTypes.cftCharacter, "'" & Action.GetActionStatusCode(Action.ActionStatuses.astCancelled) & "','" & Action.GetActionStatusCode(Action.ActionStatuses.astCompleted) & "'", CDBField.FieldWhereOperators.fwoNotIn)
        'Add Document Class joins
        .Add("a.created_by", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracketTwice)
        .Add("dc.creator_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("a.created_by#2", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("u.department", mvEnv.User.Department)
        .Add("dc.department_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        .Add("a.created_by#3", mvEnv.User.UserID, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("u.department#2", mvEnv.User.Department, CDBField.FieldWhereOperators.fwoNotEqual)
        .Add("dc.public_header", "Y", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
      End With

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "actions a", vWhereFields, "deadline ASC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      Dim vAccessRights As New AccessRights()
      For Each vRow As CDBDataRow In pDataTable.Rows
        vAccessRights.Init(mvEnv)
        vRow.Item("AccessRights") = CInt(vAccessRights.GetClassRights(vRow.Item("DocumentClass"), vRow.Item("Department"), vRow.Item("CreatedBy"))).ToString
      Next

    End Sub

    <EnumEquivalent(DataSelectionTypes.dstWorkstreamCategories)>
    Private Sub GetWorkstreamCategories(ByVal pDataTable As CDBDataTable)

      Dim vFields As String = "cat.activity,cat.activity_value,cat.quantity,cat.activity_date,cat.source,cat.valid_from,cat.valid_to,cat.amended_by,cat.amended_on,cat.notes,activity_desc,activity_value_desc,source_desc,rgb_value,ctl.workstream_id,cat.category_id,'' NoteFlag, ''Status,''StatusOrder"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("categories cat", "ctl.category_id", "cat.category_id")
      vAnsiJoins.Add("activities a", "cat.activity", "a.activity")
      vAnsiJoins.Add("activity_values av", "cat.activity", "av.activity", "cat.activity_value", "av.activity_value")
      vAnsiJoins.Add("sources s", "cat.source", "s.source")

      Dim vWhereFields As New CDBFields()
      If mvParameters.Exists("CategoryId") Then
        vWhereFields.Add("ctl.category_id", mvParameters("CategoryId").Value)
      Else
        If mvParameters.Exists("WorkstreamId") Then vWhereFields.Add("Workstream_id", mvParameters("WorkstreamId").Value)
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "category_links ctl", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, "", True)

      Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
      Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")

      If vStatus Then pDataTable.Columns("Status").AttributeName = "status" 'Why
      If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag" 'Why

      For Each vRow As CDBDataRow In pDataTable.Rows
        If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
        If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
      Next

    End Sub

    Private Sub ApplyOwnershipAccessSecurity(pDataTable As CDBDataTable, pAccessLevelColumn As String, Optional vAllowedCols As List(Of String) = Nothing, Optional ByVal pBrowseAccessLevelIdentifier As String = "B")
      If vAllowedCols Is Nothing Then vAllowedCols = New List(Of String)
      If pDataTable IsNot Nothing AndAlso pDataTable.Columns.ContainsKey(pAccessLevelColumn) Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item(pAccessLevelColumn) = pBrowseAccessLevelIdentifier Then
            For vIdx As Integer = 1 To pDataTable.Columns.Count
              If Not vAllowedCols.Contains(pDataTable.Columns(vIdx).Name) Then
                vRow.Item(vIdx) = Nothing
              End If
            Next
          End If
        Next
      End If
    End Sub

  End Class

End Namespace
