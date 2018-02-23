Namespace Access

  Public Class ExamDataSelection
    Inherits DataSelection

    Public Enum ExamDataSelectionTypes As Integer
      dstNone = 5000
      dstExamBookingUnits
      dstExamCandidateActivites
      dstExamCentreAssessmentTypes
      dstExamCentreContacts
      dstExamCentreUnits
      dstExamCentreUnitSelection
      dstExamCentres
      dstExamExemptions
      dstExamExemptionUnits
      dstExamExemptionUnitSelection
      dstExamPersonnel
      dstExamPersonnelAssessTypes
      dstExamPersonnelExpenses
      dstExamSchedule
      dstExamSchedulePersonnel
      dstExamSessionCentres
      dstExamSessionCentreSelection
      dstExamSessions
      dstExamStudentBookingUnits
      dstExamStudentComponentResults
      dstExamStudentEligibility
      dstExamStudentExemptionHistory
      dstExamStudentExemptions
      dstExamStudentHeader
      dstExamStudentResults
      dstExamStudentUnitHeader
      dstExamUnitAssessmentTypes
      dstExamUnitEligibilityChecks
      dstExamUnitGrades
      dstExamUnitLinks
      dstExamUnitPersonnel
      dstExamUnitPrerequisites
      dstExamUnitProducts
      dstExamUnits
      dstExamScheduleAllCentres
      dstExamPersonnelMarkerInfo
      dstExamUnitCandidates
      dstExamUnitMarkerAllocation
      dstExamUnitMarkerAllocationList
      dstExamMarkerList
      dstExamMaintenanceButtons     'These are SelectionPages
      dstExamMaintenanceCourses     'These are SelectionPages
      dstExamMaintenancePersonnel   'These are SelectionPages
      dstExamMaintenanceCentres     'These are SelectionPages
      dstExamMaintenanceSessions    'These are SelectionPages
      dstExamMaintenanceExemptions  'These are SelectionPages
      dstExamCentreActions
      dstExamCentreHistory
      dstExamCentreCategories
      dstExamUnitLinkCategories
      dstExamCentreUnitLinkCategories
      dstExamAccreditationHistory
      dstExamCentreUnitDetails
      dstExamUnitGradeHistory
      dstExamUnitHeaderGradeHistory
      dstExamCentreDocuments
      dstExamUnitLinkDocuments
      dstExamCentreUnitLinkDocuments
      dstExamDocumentSubjects
      dstDocumentHistory
      dstExamUnitStudyModes
      dstExamCentreUnitStudyModes
      dstExamUnitCertRunTypes
      dstExamBookingStudyModes
      dstExamCertReprintTypes
      dstExamSessionLookup
      dstExamCentreLookup
      dstExamUnitLookup
    End Enum

    Private mvExamSelectionType As ExamDataSelectionTypes

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pDataSelectionType As ExamDataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsageType As DataSelectionUsages)
      Init(pEnv, pDataSelectionType, pParams, pListType, pUsageType, "")
    End Sub

    Private Overloads Sub Init(ByVal pEnv As CDBEnvironment, ByVal pType As ExamDataSelectionTypes, ByVal pParams As CDBParameters, ByVal pListType As DataSelectionListType, ByVal pUsage As DataSelectionUsages, ByVal pGroup As String)
      mvEnv = pEnv
      mvExamSelectionType = pType
      mvParameters = pParams
      mvUsage = pUsage
      mvGroupCode = pGroup
      mvAvailableUsages = DataSelectionUsages.dsuCare Or DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
      mvDisplayListItems = Nothing
      mvDataSelectionListType = pListType

      Dim vPrimaryList As Boolean

      Select Case mvExamSelectionType
        Case ExamDataSelectionTypes.dstExamBookingUnits
          mvResultColumns = "ExamBookingUnitId,ExamBookingId,ExamUnitId1,ExamUnitId,ExamUnitCode,ExamUnitDescription,ExamScheduleId,ExamPersonnelId,ExamCandidateNumber,DeskNumber,BatchNumber,TransactionNumber,LineNumber,AttemptNumber,ExamStudentUnitStatus,OriginalMark,ModeratedMark,TotalMark,OriginalGrade,ModeratedGrade,TotalGrade,OriginalResult,ModeratedResult,TotalResult,EntryDate,ExpiryDate,DoneDate,ExpirySession,ActivityGroup,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamStudentUnitStatus,OriginalMark,ModeratedMark,TotalMark,OriginalGrade,ModeratedGrade,TotalGrade,OriginalResult,ModeratedResult,TotalResult,EntryDate,ExpiryDate,DoneDate,ExpirySession"
          mvHeadings = "Status,Mark,Moderated,Total,Grade,Moderated,Total,Result,Moderated,Total,Entry Date,Expiry Date,Done Date,Expiry Session"
          mvRequiredItems = "ExamBookingUnitId,ExamBookingId,ExamUnitId,ExamScheduleId,ExamPersonnelId,ExamUnitId1"
          mvCode = "XSU"
        Case ExamDataSelectionTypes.dstExamCandidateActivites
          mvResultColumns = "ExamCandidateActivityId,ExamBookingUnitId,ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,Quantity,Status,SourceDesc,ValidFrom,ValidTo,NoteFlag"
          mvHeadings = "Category,Value,Quantity,Status,Source,Valid From,Valid To,Notes ?"
          mvRequiredItems = "ExamBookingUnitId,ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,Notes,ExamCandidateActivityId"
          mvMaintenanceDesc = "Category"
          mvCode = "XCA"
        Case ExamDataSelectionTypes.dstExamCentreAssessmentTypes
          mvResultColumns = "ExamCentreAssessmentTypeId,ExamCentreId,ExamAssessmentType,ExamAssessmentTypeDesc,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamAssessmentType,ExamAssessmentTypeDesc"
          mvHeadings = "Assessment Type,Description"
          mvRequiredItems = "ExamCentreAssessmentTypeId,ExamCentreId"
          mvCode = "XCT"
        Case ExamDataSelectionTypes.dstExamCentreActions
          mvResultColumns = "MasterAction,ActionLevel,SequenceNumber,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,CreatedBy,CreatedOn,Deadline,ScheduledOn,CompletedOn,ActionPriority,ActionStatus,LinkType,LinkTypeDesc,SortColumn,ContactNumber,ContactName,PhoneNumber,OrganisationNumber,Name,Position,Topic,SubTopic,TopicDesc,SubTopicDesc,DurationDays,DurationHours,DurationMinutes,DocumentClass,ActionText,OutlookId"
          mvSelectColumns = "LinkTypeDesc,ActionNumber,ActionDesc,ActionPriorityDesc,ActionStatusDesc,ActionText,DetailItems,Deadline,CreatedBy,NewColumn,ScheduledOn,CreatedOn,NewColumn2,CompletedOn"
          mvHeadings = DataSelectionText.String17785
          mvRequiredItems = "ActionStatus,MasterAction"
          mvCode = "CAC"
        Case ExamDataSelectionTypes.dstExamCentreHistory
          mvResultColumns = "ExamCentreHistoryId,ExamCentreId,ExamCentreDescTimestamp,ExamCentreDescription,AmendedBy,AmendedOn,CreatedBy,CreatedOn"
          mvSelectColumns = "ExamCentreHistoryId,ExamCentreId,ExamCentreDescTimestamp,ExamCentreDescription,AmendedBy,AmendedOn,CreatedBy,CreatedOn"
          mvHeadings = "Exam Centre History Id,Exam Centre Id,Timestamp,Description,Amended By,Amended On,Created By,Created On"
          mvRequiredItems = "ExamCentreHistoryId,ExamCentreId,ExamCentreDescTimestamp"
          mvDisplayTitle = "Centre History"
          mvCode = "XCH"
        Case ExamDataSelectionTypes.dstExamCentreContacts
          mvResultColumns = "ExamCentreContactId,ExamCentreId,ContactNumber,ExamContactType,ExamContactTypeDesc,CreatedBy,CreatedOn,AmendedBy,AmendedOn," & ContactNameResults()
          mvSelectColumns = "ContactName,ExamContactTypeDesc"
          mvHeadings = "Name,Contact Type"
          mvRequiredItems = "ExamCentreContactId,ExamCentreId"
          mvCode = "XCC"
        Case ExamDataSelectionTypes.dstExamCentreUnits
          mvResultColumns = "ExamCentreUnitId,ExamCentreId,ExamUnitId,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",ExamUnitLinkId"
          mvSelectColumns = ""
          mvHeadings = ""
          mvRequiredItems = "ExamCentreUnitId,ExamCentreId,ExamUnitId"
          mvCode = "XCU"
        Case ExamDataSelectionTypes.dstExamCentreUnitSelection
          mvResultColumns = "ExamUnitId1,ExamCentreUnitId,ExamCentreId,ExamUnitId,ExamUnitCode,ExamUnitDescription"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",ExamUnitLinkId,ParentUnitLinkId,LocalName,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo"
          mvSelectColumns = ""
          mvHeadings = ""
          mvRequiredItems = "ExamUnitId1,ExamCentreUnitId,ExamCentreId,ExamUnitId,ExamUnitlinkId,ParentUnitLinkId,LocalName"
        Case ExamDataSelectionTypes.dstExamCentres
          mvResultColumns = "ExamCentreId,OrganisationNumber,AddressNumber,ContactNumber,ExamCentreCode,ExamCentreDescription,ValidFrom,ValidTo,Capacity,LastVisitDate,NextVisitDate,ExamCentreParentId,AdditionalCapacity,AcceptSpecialRequirements,Overseas,WebPublish,ExamCentreRateType,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",AccreditationStatus,AccreditationValidFrom,AccreditationValidTo"
          mvSelectColumns = "ExamCentreCode,ExamCentreDescription,ValidFrom,ValidTo,Capacity,LastVisitDate,NextVisitDate,Overseas"
          mvHeadings = "Centre Code,Exam Centre Description,Valid From,Valid To,Capacity,Last Visit Date,Next Visit Date,Overseas"
          mvRequiredItems = "ExamCentreId,ExamCentreParentId"
          mvCode = "XPC"
        Case ExamDataSelectionTypes.dstExamExemptions
          mvResultColumns = "ExamExemptionId,ExamExemptionCode,ExamExemptionDescription,Product,Rate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamExemptionCode,ExamExemptionDescription,Product,Rate"
          mvHeadings = "Exam Exemption Code,Exam Exemption Description,Product,Rate"
          mvRequiredItems = "ExamExemptionId"
          mvCode = "XEP"
        Case ExamDataSelectionTypes.dstExamExemptionUnits
          mvResultColumns = "ExamExemptionUnitId,ExamUnitId,ExamExemptionProductId,ExamUnitCode,ExamUnitDescription,ExamExemptionCode,ExamExemptionDescription,ProductCode,Product,RateCode,Rate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription"
          mvHeadings = "Code,Description"
          mvRequiredItems = "ExamExemptionUnitId,ExamUnitId,ExamExemptionProductId"
          mvCode = "XEU"
        Case ExamDataSelectionTypes.dstExamExemptionUnitSelection
          mvResultColumns = "ExamUnitId1,ExamExemptionUnitId,ExamExemptionId,ExamUnitId,ExamUnitCode,ExamUnitDescription,ExamUnitLinkId,ParentUnitLinkId"
          mvSelectColumns = ""
          mvHeadings = ""
          mvRequiredItems = "ExamUnitId1,ExamExemptionUnitId,ExamExemptionId,ExamUnitId"
        Case ExamDataSelectionTypes.dstExamPersonnel
          mvResultColumns = "ExamPersonnelId,ContactNumber,ValidFrom,ValidTo,ExamPersonnelType,ExamPersonnelTypeDesc,ExamMarker,TrainedDate,MaximumStudents,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn," & ContactNameResults()
          mvSelectColumns = "ContactName,ValidFrom,ValidTo,ExamPersonnelType,TrainedDate,MaximumStudents,Notes"
          mvHeadings = "Name,Valid From,Valid To,Exam Personnel Type,Trained Date,Maximum Students,Notes"
          mvRequiredItems = "ExamPersonnelId"
          mvCode = "XP"
        Case ExamDataSelectionTypes.dstExamPersonnelAssessTypes
          mvResultColumns = "ExamPersonnelAssessTypeId,ExamPersonnelId,ExamAssessmentType,ExamAssessmentTypeDesc,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamAssessmentType,ExamAssessmentTypeDesc"
          mvHeadings = "Assessment Type,Description"
          mvRequiredItems = "ExamPersonnelAssessTypeId,ExamPersonnelId"
          mvCode = "XPT"
        Case ExamDataSelectionTypes.dstExamPersonnelExpenses
          mvResultColumns = "ExamPersonnelExpenseId,ExamPersonnelId,ExamExpenseType,Amount,AppliedDate,PaidDate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamExpenseType,Amount,AppliedDate,PaidDate"
          mvHeadings = "Expense Type,Amount,Applied Date,Paid Date"
          mvRequiredItems = "ExamPersonnelExpenseId,ExamPersonnelId"
          mvCode = "XPE"
        Case ExamDataSelectionTypes.dstExamSchedule
          mvResultColumns = "ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitId,ExamCentreCode,ExamCentreDescription,StartDate,StartTime,EndTime,Capacity,NumberOfCandidates,AdditionalCapacity,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamCentreCode,ExamCentreDescription,StartDate,StartTime,EndTime,Capacity,NumberOfCandidates"
          mvHeadings = "Centre,Description,Start Date,Start Time,End Time,Capacity,Candidates"
          mvRequiredItems = "ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitId"
          mvCode = "XSC"
        Case ExamDataSelectionTypes.dstExamSchedulePersonnel
          mvResultColumns = "ExamSchedulePersonnelId,ExamScheduleId,ExamPersonnelId,ExamPersonnelType,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamPersonnelType"
          mvHeadings = "Exam Personnel Type"
          mvRequiredItems = "ExamSchedulePersonnelId,ExamScheduleId,ExamPersonnelId"
          mvCode = "XSP"
        Case ExamDataSelectionTypes.dstExamSessionCentres
          mvResultColumns = "ExamSessionCentreId,ExamCentreId,ExamSessionId,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = ""
          mvHeadings = ""
          mvRequiredItems = "ExamSessionCentreId,ExamCentreId,ExamSessionId"
          mvCode = "XST"
        Case ExamDataSelectionTypes.dstExamSessionCentreSelection
          mvResultColumns = "ExamSessionCentreId,ExamCentreParentId,ExamCentreId,ExamSessionId,ExamCentreCode,ExamCentreDescription"
          mvSelectColumns = ""
          mvHeadings = ""
          mvRequiredItems = "ExamSessionCentreId,ExamCentreId,ExamSessionId"
          mvCode = "XST"
        Case ExamDataSelectionTypes.dstExamSessions
          mvResultColumns = "ExamSessionId,ExamSessionYear,ExamSessionMonth,ExamSessionCode,ExamSessionDescription,SequenceNumber,ValidFrom,ValidTo,HomeClosingDate,OverseasClosingDate,WebPublish,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn,ResultsReleaseDate"
          mvSelectColumns = "ExamSessionYear,ExamSessionMonth,ExamSessionCode,ExamSessionDescription,ValidFrom,ValidTo,HomeClosingDate,OverseasClosingDate,Notes"
          mvHeadings = "Exam Session Year,Exam Session Month,Exam Session Code,Exam Session Description,Valid From,Valid To,Home Closing Date,Overseas Closing Date,Notes"
          mvRequiredItems = "ExamSessionId"
          mvCode = "XSS"
        Case ExamDataSelectionTypes.dstExamStudentEligibility
          mvResultColumns = "ExamUnitEligibilityCheckId,ContactNumber,Proven,ProvedDate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "Proven,ProvedDate"
          mvHeadings = "Proven,Proved Date"
          mvRequiredItems = "ExamUnitEligibilityCheckId"
          mvCode = "XEC"
        Case ExamDataSelectionTypes.dstExamStudentExemptionHistory
          mvResultColumns = "ExamStudentExemptionHistId,ExamStudentExemptionId,ExamExemptionStatus,ExamExemptionStatusDesc,StatusDate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamExemptionStatus,ExamExemptionStatusDesc,StatusDate"
          mvHeadings = "Status,Description,Status Date"
          mvRequiredItems = "ExamStudentExemptionHistId,ExamStudentExemptionId"
          mvCode = "XEH"
          mvDisplayTitle = "History"
        Case ExamDataSelectionTypes.dstExamStudentExemptions
          mvResultColumns = "ExamStudentExemptionId,ExamExemptionProductId,ContactNumber,BatchNumber,TransactionNumber,LineNumber,ExamExemptionStatus,StatusDate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamExemptionStatus,StatusDate"
          mvHeadings = "Exam Exemption Status,Status Date"
          mvRequiredItems = "ExamStudentExemptionId,ExamExemptionProductId"
          mvCode = "XSE"
        Case ExamDataSelectionTypes.dstExamStudentHeader
          Dim vFields As String = "exam_student_id,exam_unit_id,contact_number,first_session_id,last_session_id,last_marked_date,last_graded_date,created_by,created_on,amended_by,amended_on"
          mvResultColumns = "ExamStudentId,ExamUnitId,ContactNumber,FirstSessionId,LastSessionId,LastMarkedDate,LastGradedDate,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "Contact Number,FirstSessionId,LastSessionId,LastMarkedDate,LastGradedDate"
          mvHeadings = "Contact Number,First Session,Last Session,Last Marked,Last Graded"
          mvRequiredItems = "ExamStudentId,ExamUnitId"
          mvCode = "XSH"
        Case ExamDataSelectionTypes.dstExamStudentUnitHeader
          mvResultColumns = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId1,ExamUnitId,ExamUnitCode,ExamUnitDescription,Attempts,CurrentMark,CurrentGrade,CurrentResult,ExamGradeSequenceNumber,GradeIsPass,Expires,ExamUnitLinkId,ParentUnitLinkId,CreatedBy,CreatedOn,AmendedBy,AmendedOn,CanEditResults,ResultsReleaseDate,PreviousMark,PreviousGrade,PreviousResult"
          mvSelectColumns = "Attempts,CurrentMark,CurrentGrade,CurrentResult"
          mvHeadings = "ExamUnitCode,ExamUnitDescription,Attempts,Mark,Grade,Result"
          mvRequiredItems = "ExamStudentUnitHeaderId,ExamStudentHeaderId,ExamUnitId1,ExamUnitId,CanEditResults"
          mvCode = "XUH"
        Case ExamDataSelectionTypes.dstExamUnitAssessmentTypes
          mvResultColumns = "ExamUnitAssessmentTypeId,ExamUnitId,ExamAssessmentType,ExamAssessmentTypeDesc,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamAssessmentType,ExamAssessmentTypeDesc"
          mvHeadings = "Assessment Type,Description"
          mvRequiredItems = "ExamUnitAssessmentTypeId,ExamUnitId"
          mvCode = "XAT"
        Case ExamDataSelectionTypes.dstExamUnitEligibilityChecks
          mvResultColumns = "ExamUnitEligibilityCheckId,ExamUnitId,EligibilityCheckText,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "EligibilityCheckText"
          mvHeadings = "Eligibility Check Text"
          mvRequiredItems = "ExamUnitEligibilityCheckId,ExamUnitId"
          mvCode = "XUE"
        Case ExamDataSelectionTypes.dstExamUnitGrades
          mvResultColumns = "ExamUnitGradeId,ExamUnitId,ExamUnitCode,ExamGrade,ExamGradeDesc,SequenceNumber,ConditionNumber,ClauseNumber,ConditionDesc,ExamGradeConditionType,ExamGradeConditionTypeDesc,GradeUnits,ExamGradeOperator,RequiredValue,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamGradeDesc,ConditionDesc,ExamGradeConditionTypeDesc,GradeUnits,ExamGradeOperator,RequiredValue"
          mvHeadings = "Grade,Condition,Type,Units,Operator,Required Value"
          mvRequiredItems = "ExamUnitGradeId,ExamUnitId,ExamGrade"
          mvCode = "XUG"
        Case ExamDataSelectionTypes.dstExamUnitLinks
          mvResultColumns = "ExamUnitId1,ExamUnitId2,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamUnitId1,ExamUnitId2"
          mvHeadings = "Exam Unit Id 1,Exam Unit Id 2"
          mvRequiredItems = ""
        Case ExamDataSelectionTypes.dstExamUnitPersonnel
          mvResultColumns = "ExamUnitPersonnelId,ExamPersonnelId,ExamUnitId,ValidFrom,ValidTo,ExamPersonnelType,ExamPersonnelTypeDesc,MaximumStudents,GeographicalRegion,ExamMarkerOption,ExamMarkerOptionDesc,ActualLoadSize,CreatedBy,CreatedOn,AmendedBy,AmendedOn,ContactNumber," & ContactNameResults()
          mvSelectColumns = "ContactName,ValidFrom,ValidTo,ExamPersonnelTypeDesc,MaximumStudents,GeographicalRegion,ExamMarkerOptionDesc"
          mvHeadings = "Name,Valid From,Valid To,Type,Max Students,Geographical Region,Marker Option"
          mvRequiredItems = "ExamUnitPersonnelId,ExamPersonnelId,ExamUnitId"
          mvCode = "XUP"
        Case ExamDataSelectionTypes.dstExamUnitPrerequisites
          mvResultColumns = "ExamUnitId,ExamPrerequisiteUnitId,ExamPrerequisiteUnitCode,ExamPrerequisiteUnitDescription,MinimumGrade,MinimumGradeDescription,PassRequired,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamPrerequisiteUnitCode,ExamPrerequisiteUnitDescription,PassRequired,MinimumGradeDescription"
          mvHeadings = "Unit Code,Exam Unit Description,Pass Required,Minimum Grade"
          mvRequiredItems = "ExamUnitId,ExamPrerequisiteUnitId"
          mvCode = "XUPR"
        Case ExamDataSelectionTypes.dstExamUnitProducts
          mvResultColumns = "ExamUnitProductId,ExamUnitId,ProductCode,Product,RateCode,Rate,Quantity,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ProductCode,Product,RateCode,Rate,Quantity"
          mvHeadings = "Product,Description,Rate,Description,Quantity"
          mvRequiredItems = "ExamUnitProductId,ExamUnitId"
          mvCode = "XUR"
        Case ExamDataSelectionTypes.dstExamUnits
          mvResultColumns = "ExamUnitId1,ExamUnitId,ExamBaseUnitId,ExamSessionId,ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ScheduleRequired,MarkerRequired,ExamQuestion,ExamUnitStatus,SequenceNumber,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,ExamUnitReplacedById,AllowExemptions,ExemptionMark,ExamMarkerStatus,PapersPerMarker,IncludeView,ExcludeView,AllowDowngrade,WebPublish,ActivityGroup,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",LongDescription,ExamUnitLinkId,ParentUnitLinkId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,IsGradingEndpoint,LocalName,CourseAccreditation,CourseAccreditationValidFrom,CourseAccreditationValidTo"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ExamUnitStatus,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,AllowExemptions,ExemptionMark,ExamMarkerStatus,PapersPerMarker,Notes,IsGradingEndpoint"
          mvHeadings = "Unit Code,Exam Unit Description,Subject,Skill Level,Exam Unit Type,Exam Unit Status,Session Based,Valid From,Valid To,Product,Rate,Date Approved,Registration Date,Qcf Level,Number Of Credits,Nvq Code,Svq Code,Unit Time Limit,Time Limit Type,Minimum Students,Maximum Students,Student Count,Minimum Age,Allow Bookings,Exam Mark Type,Mark Factor,Awarding Body,Allow Exemptions,Exemption Mark,Papers per Marker,Notes,Grading End-Point"
          mvRequiredItems = "ExamUnitId1,ExamUnitId,ExamSessionId,ExamUnitReplacedById,ScheduleRequired,MarkerRequired,ExamQuestion"
          mvCode = "XU"
        Case ExamDataSelectionTypes.dstExamStudentBookingUnits
          mvResultColumns = "ExamUnitId1,ParentUnitLinkId,ExamUnitId,ExamUnitLinkId,ExamBaseUnitId,ExamSessionId,ExamBookingUnitId,ExamBookingId,TotalMark,TotalGrade,TotalResult,ExamStudentUnitStatus,ContactNumber,DoneDate,SpecialRequirements,Booked,Passed,ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ScheduleRequired,ExamUnitStatus,SequenceNumber,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,ExamUnitReplacedById,AllowExemptions,ExemptionMark,ActivityGroup,Notes,ExamCentreUnitId,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvResultColumns &= ",AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,ExamQuestion"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ExamUnitStatus,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,AllowExemptions,ExemptionMark,Notes,ContactNumber,Booked,Passed,TotalMark,DoneDate,SpecialRequirements"
          mvHeadings = "Unit Code,Exam Unit Description,Subject,Skill Level,Exam Unit Type,Exam Unit Status,Session Based,Valid From,Valid To,Product,Rate,Date Approved,Registration Date,Qcf Level,Number Of Credits,Nvq Code,Svq Code,Unit Time Limit,Time Limit Type,Minimum Students,Maximum Students,Student Count,Minimum Age,Allow Bookings,Exam Mark Type,Mark Factor,Awarding Body,Allow Exemptions,Notes,Contact Number,Booked,Passed,Total Mark,Done Date,Special Requirements"
          mvRequiredItems = "ExamUnitId1,ExamUnitId,ExamSessionId,ExamBookingUnitId,ExamBookingId,ExamUnitReplacedById,ScheduleRequired,CanEditResults"
          mvCode = "XBU"
        Case ExamDataSelectionTypes.dstExamStudentComponentResults
          mvResultColumns = "ExamBookingUnitId,ContactNumber,ExamUnitCode,ExamUnitDescription,RawMark,RawMarkCheck,OriginalMark,OriginalMarkCheck,ExamMarkType,MarkFactor"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,RawMark"
          mvHeadings = "Code,Description,Raw Mark,Grade,Pass/Fail"
          mvRequiredItems = "ExamBookingUnitId,ExamUnitCode,ExamUnitDescription,ExamMarkType,RawMark,RawMarkCheck,OriginalMark,OriginalMarkCheck,MarkFactor"
        Case ExamDataSelectionTypes.dstExamStudentResults
          mvResultColumns = "ExamBookingUnitId,ContactNumber,RawMark,RawMarkCheck,OriginalMark,OriginalMarkCheck,OriginalGrade,OriginalGradeCheck,OriginalResult,OriginalResultCheck,ExamMarkType,ExamUnitChildLink,ExamUnitId,ExamUnitCode," & ContactNameResults()
          mvSelectColumns = "ContactNumber,ContactName,RawMark,OriginalGrade,OriginalResult"
          mvHeadings = "Contact Number,Name,Raw Mark,Grade,Pass/Fail"
          mvRequiredItems = "ExamBookingUnitId,ExamMarkType,RawMark,RawMarkCheck,OriginalMark,OriginalMarkCheck,OriginalGrade,OriginalGradeCheck,OriginalResult,OriginalResultCheck,ExamUnitChildLink,ExamUnitId,ExamUnitCode"
          'mvCode = "XSR"
        Case ExamDataSelectionTypes.dstExamScheduleAllCentres
          mvResultColumns = "ExamScheduleId,ExamSessionId,ExamCentreId,ExamCentreParentId,ExamUnitId,ExamCentreCode,ExamCentreDescription,StartDate,StartTime,EndTime,Capacity,AdditionalCapacity,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          mvSelectColumns = "ExamCentreCode,ExamCentreDescription,StartDate,StartTime,EndTime,Capacity"
          mvHeadings = "Centre,Description,Start Date,Start Time,End Time,Capacity"
          mvRequiredItems = "ExamScheduleId,ExamSessionId,ExamCentreId,ExamUnitId,ExamCentreParentId"
        Case ExamDataSelectionTypes.dstExamPersonnelMarkerInfo
          mvResultColumns = "ExamPersonnelId,ExamSessionId,ExamCentreId,ExamUnitId,ExamSessionCode,ExamSessionDescription,ExamUnitCode,ExamUnitDescripion,ExamCentreCode,ExamCentreDescription,NumberOfPapers,MarkerNumber"
          mvSelectColumns = "ExamSessionCode,ExamUnitCode,ExamCentreCode,NumberOfPapers,MarkerNumber"
          mvHeadings = "Session,Unit,Centre,No of Papers,Marker No"
          mvRequiredItems = "ExamPersonnelId,ExamSessionId,ExamCentreId,ExamUnitId"
          mvCode = "XPM"
        Case ExamDataSelectionTypes.dstExamUnitCandidates
          mvResultColumns = "ExamBookingId,ExamCentreId,ExamUnitId,ExamScheduleId,ExamCentreCode,ExamCentreDescription,ExamCandidateNumber"
          mvSelectColumns = "ExamCentreCode,ExamCentreDescription,ExamCandidateNumber"
          mvHeadings = "Centre,Description,Candidate Number"
          mvRequiredItems = "ExamBookingId,ExamCentreId,ExamUnitId,ExamScheduleId"
          mvCode = "XUC"
        Case ExamDataSelectionTypes.dstExamUnitMarkerAllocation
          mvResultColumns = "ExamUnitId,ExamPersonnelId,ContactNumber,MarkerNumber,NumberOfPapers,Unallocated," & ContactNameResults()
          mvSelectColumns = "ContactNumber,ContactName,MarkerNumber,NumberOfPapers"
          mvHeadings = "Contact Number,Name,Marker No,No of Papers"
          mvRequiredItems = "ExamUnitId,ExamPersonnelId,ContactNumber,MarkerNumber,Unallocated"
          mvCode = "XMA"
        Case ExamDataSelectionTypes.dstExamUnitMarkerAllocationList
          mvResultColumns = "Select,ExamMarkingBatchDetailId,ExamBookingUnitId,ExamUnitId,ExamPersonnelId,ExamCentreId,ExamCentreDscription,ExamCentreCode,ExamCandidateNumber,ContactNumber," & ContactNameResults()
          mvSelectColumns = "Select,ContactNumber,ContactName,ExamCandidateNumber,ExamCentreCode,ExamCentreDscription"
          mvHeadings = "Select,Contact Number,Name,Candidate Number,Centre,Description"
          mvRequiredItems = "Select,ExamMarkingBatchDetailId,ExamBookingUnitId,ExamUnitId,ExamPersonnelId,ContactNumber"
          mvCode = "XLM"
        Case ExamDataSelectionTypes.dstExamMarkerList
          mvResultColumns = "ContactNumber,ExamUnitId,ExamPersonnelId,ExamPersonnelType,ExamPersonnelTypeDesc,ExamMarkerStatus," & ContactNameResults()
          mvSelectColumns = "ContactNumber,ContactName,ExamPersonnelTypeDesc"
          mvHeadings = "Contact Number,Name,Marker Type"
          mvRequiredItems = "ContactNumber,ExamPersonnelId,ExamUnitId"
          mvCode = "XMK"
        Case ExamDataSelectionTypes.dstExamMaintenanceButtons
          mvResultColumns = "Courses,Personnel,Centres,Sessions,Exemptions"
          mvSelectColumns = "Courses,Personnel,Centres,Sessions,Exemptions"
          mvHeadings = "Courses,Personnel,Centres,Sessions,Exemptions"
          mvRequiredItems = ""
          mvCode = "XMB"
        Case ExamDataSelectionTypes.dstExamMaintenanceCourses
          mvResultColumns = "Personnel,Requirements,Resources,Grading,Assessment Types,Prerequisites,Marker Allocation,Schedule,Categories,Course Documents,Study Modes,Certificates"
          mvSelectColumns = "Personnel,Requirements,Resources,Grading,Assessment Types,Prerequisites,Marker Allocation,Schedule,Categories,Course Documents,Study Modes,Certificates"
          mvHeadings = "Personnel,Requirements,Resources,Grading,Assessment Types,Prerequisites,Marker Allocation,Schedule,Categories,Course Documents,Study Modes,Certificates"
          mvRequiredItems = ""
          mvCode = "XMC"
        Case ExamDataSelectionTypes.dstExamMaintenancePersonnel
          mvResultColumns = "Expenses,Marker Information,Assessment Types"
          mvSelectColumns = "Expenses,Marker Information,Assessment Types"
          mvHeadings = "Expenses,Marker Information,Assessment Types"
          mvRequiredItems = ""
          mvCode = "XMP"
        Case ExamDataSelectionTypes.dstExamMaintenanceCentres
          mvResultColumns = "Contacts,Courses,Assessment Types,Actions,Categories,Centre Documents"
          mvSelectColumns = "Contacts,Courses,Assessment Types,Actions,Categories,Centre Documents"
          mvHeadings = "Contacts,Courses,Assessment Types,Actions,Categories,Centre Documents"
          mvRequiredItems = ""
          mvCode = "XME"
          CheckCustomForms(mvCode)
        Case ExamDataSelectionTypes.dstExamMaintenanceSessions
          mvResultColumns = "Centres"
          mvSelectColumns = "Centres"
          mvHeadings = "Centres"
          mvRequiredItems = ""
          mvCode = "XMS"
        Case ExamDataSelectionTypes.dstExamMaintenanceExemptions
          mvResultColumns = "Courses"
          mvSelectColumns = "Courses"
          mvHeadings = "Courses"
          mvRequiredItems = ""
          mvCode = "XMX"
        Case ExamDataSelectionTypes.dstExamCentreCategories
          mvResultColumns = "ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,ExamCentreId,CategoryId,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,ActivityDate,Quantity,SourceCode,ValidFrom,ValidTo,NoteFlag,Status,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40029     'Activity Desc,Value,Quantity,Source,Valid from,Valid to,Notes?,Amended by,Amended on
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200"
          mvRequiredItems = "CategoryId,ExamCentreId,ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,RgbActivityValue"
          mvDescription = DataSelectionText.String40030    'Exam Centre Category
          mvCode = "XCCA"
        Case ExamDataSelectionTypes.dstExamUnitLinkCategories
          mvResultColumns = "ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,ExamUnitLinkId,CategoryId,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,ActivityDate,Quantity,SourceCode,ValidFrom,ValidTo,NoteFlag,Status,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40029     'Activity Desc,Value,Quantity,Valid from,Valid to,Notes?,Amended by,Amended on
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "CategoryId,ExamUnitLinkId,ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,RgbActivityValue"
          mvDescription = DataSelectionText.String40031    'Exam Unit Link Category
          mvCode = "XUCA"
        Case ExamDataSelectionTypes.dstExamCentreUnitLinkCategories
          mvResultColumns = "ActivityCode,ActivityValueCode,Quantity,ActivityDate,SourceCode,ValidFrom,ValidTo,AmendedBy,AmendedOn,Notes,ActivityDesc,ActivityValueDesc,SourceDesc,RgbActivityValue,ExamCentreUnitId,CategoryId,ExamUnitLinkId,NoteFlag,Status,StatusOrder"
          mvSelectColumns = "ActivityDesc,ActivityValueDesc,ActivityDate,Quantity,SourceCode,ValidFrom,ValidTo,NoteFlag,Status,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40029     'Activity Desc,Value,Quantity,Source,Valid from,Valid to,Notes?,Amended by,Amended on
          mvWidths = "1800,1800,1200,1200,1800,3600,1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "CategoryId,ExamUnitLinkId,ExamCentreUnitId,ActivityCode,ActivityValueCode,SourceCode,ValidFrom,ValidTo,RgbActivityValue"
          mvDescription = DataSelectionText.String40032    'Exam Centre Unit Categories
          mvCode = "XCUC"
        Case ExamDataSelectionTypes.dstExamAccreditationHistory
          mvResultColumns = "AccreditationStatus,AccreditationStatusDesc,ValidFrom,ValidTo,AmendedBy,AmendedOn,AccreditationId,ExamUnitLinkId"
          mvSelectColumns = "AccreditationStatus,AccreditationStatusDesc,ValidFrom,ValidTo,AmendedBy,AmendedOn"
          mvHeadings = DataSelectionText.String40034 'Accreditation Status,Accreditation Status Desc,Valid From,Valid To,Amended By,Amended On
          mvWidths = "1200,1800,1200,1200,1200,1200"
          mvRequiredItems = "AccreditationId,ExamUnitLinkId,AccreditationStatus,AccreditationStatusDesc,ValidFrom,ValidTo,AmendedBy,AmendedOn"
          mvMaintenanceDesc = "Accreditation History"
          mvDisplayTitle = "Accreditation History"
          mvCode = "XAH"
        Case ExamDataSelectionTypes.dstExamCentreUnitDetails
          mvResultColumns = "ExamUnitId1,ExamUnitId,ExamBaseUnitId,ExamSessionId,ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ScheduleRequired,MarkerRequired,ExamQuestion,ExamUnitStatus,SequenceNumber,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,ExamUnitReplacedById,AllowExemptions,ExemptionMark,ExamMarkerStatus,PapersPerMarker,IncludeView,ExcludeView,AllowDowngrade,WebPublish,ActivityGroup,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",LongDescription,ExamUnitLinkId,ParentUnitLinkId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,IsGradingEndpoint,LocalName,CourseAccreditation,CourseAccreditationValidFrom,CourseAccreditationValidTo"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ExamUnitStatus,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,AllowExemptions,ExemptionMark,ExamMarkerStatus,PapersPerMarker,Notes"
          mvHeadings = "Unit Code,Exam Unit Description,Subject,Skill Level,Exam Unit Type,Exam Unit Status,Session Based,Valid From,Valid To,Product,Rate,Date Approved,Registration Date,Qcf Level,Number Of Credits,Nvq Code,Svq Code,Unit Time Limit,Time Limit Type,Minimum Students,Maximum Students,Student Count,Minimum Age,Allow Bookings,Exam Mark Type,Mark Factor,Awarding Body,Allow Exemptions,Exemption Mark,Papers per Marker,Notes"
          mvRequiredItems = "ExamUnitId1,ExamUnitId,ExamSessionId,ExamUnitReplacedById,ScheduleRequired,MarkerRequired,ExamQuestion"
          mvCode = "XCD"
        Case ExamDataSelectionTypes.dstExamUnitGradeHistory
          mvResultColumns = "ExamGradeChangeHistoryId,ExamBookingUnitId,ExamBookingId,ExamUnitId,ExamStudentUnitHeaderId,GradeChangeReasonCode,PreviousMark,PreviousGrade,PreviousResult,ChangedBy,ChangedOn,GradeChangeReasonDesc"
          mvSelectColumns = "PreviousMark,PreviousGrade,PreviousResult,GradeChangeReasonDesc,ChangedBy,ChangedOn"
          mvHeadings = "Previous Mark,Previous Grade,Previous Result,Grade Change Reason Desc,Changed By,Changed On"
          mvWidths = "1,1,1,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = ""
          mvCode = "XGH"
        Case ExamDataSelectionTypes.dstExamCentreDocuments, ExamDataSelectionTypes.dstExamUnitLinkDocuments, ExamDataSelectionTypes.dstExamCentreUnitLinkDocuments
          mvResultColumns = "Dated,DocumentNumber,PackageCode,LabelName,ContactNumber,DocumentTypeDesc,CreatedBy,DepartmentDesc,OurReference,Direction,TheirReference,DocumentType,DocumentClass,DocumentClassDesc,StandardDocument,Source,Recipient,Forwarded,Archiver,Completed,TopicCode,TopicDesc,SubTopicCode,SubTopicDesc,CreatorHeader,DepartmentHeader,PublicHeader,DepartmentCode,Access,StandardDocumentDesc,Precis"
          mvResultColumns = mvResultColumns & ",Subject,CallDuration,TotalDuration,SelectionSet"      'Optional Attributes
          mvSelectColumns = "DocumentNumber,Dated,Direction,Subject,OurReference,DocumentTypeDesc,TopicDesc,SubTopicDesc,DetailItems,CreatedBy,Source,PackageCode,NewColumn,DepartmentDesc,StandardDocument,NewColumn2,DocumentClassDesc,TheirReference"
          mvHeadings = DataSelectionText.String17987     'Document Number,Dated,In/ Out,Subject,Reference+,Document Type,Topic+,Sub Topic,,Creator,Source,Package,,Department,Standard Document,,Document Class,Their Reference
          mvWidths = "1200,1200,1200,1500,1500,1500,1400,1400,1200,1400,1200,600,1200,1400,1200,1200,1200,1200"
          mvRequiredItems = "Access"
          mvDescription = DataSelectionText.String17988    'Contact Documents
          mvCode = "XDD"
          mvMaintenanceDesc = "Document"
          vPrimaryList = True
        Case ExamDataSelectionTypes.dstExamUnitHeaderGradeHistory
          mvResultColumns = "ExamGradeChangeHistoryId,ExamBookingUnitId,ExamBookingId,ExamUnitId,ExamStudentUnitHeaderId,GradeChangeReasonCode,PreviousMark,PreviousGrade,PreviousResult,ChangedBy,ChangedOn,GradeChangeReasonDesc"
          mvSelectColumns = "PreviousMark,PreviousGrade,PreviousResult,GradeChangeReasonDesc,ChangedBy,ChangedOn"
          mvHeadings = "Previous Mark,Previous Grade,Previous Result,Grade Change Reason Desc,Changed By,Changed On"
          mvWidths = "1,1,1,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = ""
          mvCode = "XHG"
        Case ExamDataSelectionTypes.dstExamUnitStudyModes
          mvResultColumns = "StudyMode,StudyModeDesc,Selected"
          mvSelectColumns = "StudyMode,StudyModeDesc,Selected"
          mvHeadings = "StudyMode,StudyModeDesc,Selected"
          mvWidths = "1200,1200,1200"
          mvRequiredItems = ""
          mvCode = "XSM"
        Case ExamDataSelectionTypes.dstExamCentreUnitStudyModes
          mvResultColumns = "StudyMode,StudyModeDesc,Selected"
          mvSelectColumns = "StudyMode,StudyModeDesc,Selected"
          mvHeadings = "StudyMode,StudyModeDesc,Selected"
          mvWidths = "1200,1200,1200"
          mvRequiredItems = ""
          mvCode = "XCM"
        Case ExamDataSelectionTypes.dstExamBookingStudyModes
          mvResultColumns = "StudyMode,StudyModeDesc"
          mvSelectColumns = "StudyMode,StudyModeDesc"
          mvHeadings = "StudyMode,StudyModeDesc"
          mvWidths = "100,100"
          mvRequiredItems = ""
        Case ExamDataSelectionTypes.dstExamUnitCertRunTypes
          mvResultColumns = "ExamUnitCertRunTypeId,ExamUnitLinkId,ExamCertRunType,ExamCertRunTypeDesc,IncludeView,IncludeViewDesc,ExcludeView,ExcludeViewDesc,StandardDocument,StandardDocumentDesc"
          mvSelectColumns = "ExamUnitCertRunTypeId,ExamUnitLinkId,ExamCertRunType,ExamCertRunTypeDesc,IncludeView,IncludeViewDesc,ExcludeView,ExcludeViewDesc,StandardDocument,StandardDocumentDesc"
          mvHeadings = "ExamUnitCertRunTypeId,ExamUnitLinkId,ExamCertRunType,ExamCertRunTypeDesc,IncludeView,IncludeViewDesc,ExcludeView,ExcludeViewDesc,,StandardDocument,StandardDocumentDesc"
          mvWidths = "1200,1200,1200"
          mvRequiredItems = "ExamUnitCertRunTypeId"
          mvCode = "XCR"
        Case ExamDataSelectionTypes.dstExamCertReprintTypes
          mvResultColumns = "ExamCertReprintType,ExamCertReprintTypeDesc"
          mvSelectColumns = "ExamCertReprintType,ExamCertReprintTypeDesc"
          mvHeadings = "ExamCertReprintType,ExamCertReprintTypeDesc"
          mvWidths = "1200,1200"
          mvRequiredItems = "ExamCertReprintType,ExamCertReprintTypeDesc"
          mvCode = "XDR"
        Case ExamDataSelectionTypes.dstExamSessionLookup
          mvResultColumns = "ExamSessionCode,ExamSessionDescription,ExamSessionId,SequenceNumber,ExamSessionYear,ExamSessionMonth"
          mvSelectColumns = "ExamSessionCode,ExamSessionDescription,ExamSessionId,SequenceNumber,ExamSessionYear,ExamSessionMonth"
          mvHeadings = "ExamSessionCode,ExamSessionDescription,ExamSessionId,SequenceNumber,ExamSessionYear,ExamSessionMonth"
          mvWidths = "1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamSessionCode,ExamSessionDescription,ExamSessionId,SequenceNumber,ExamSessionYear,ExamSessionMonth"
        Case ExamDataSelectionTypes.dstExamCentreLookup
          mvResultColumns = "ExamCentreCode,ExamCentreDescription,ExamCentreId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,Overseas,HomeClosingDate,OverseasClosingDate,ClosingDate"
          mvSelectColumns = "ExamCentreCode,ExamCentreDescription,ExamCentreId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,Overseas,HomeClosingDate,OverseasClosingDate,ClosingDate"
          mvHeadings = "ExamCentreCode,ExamCentreDescription,ExamCentreId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,Overseas,HomeClosingDate,OverseasClosingDate,ClosingDate"
          mvWidths = "1200,1200,1200,1200,1200,1200,1200,1200,1200,1200"
          mvRequiredItems = "ExamCentreCode,ExamCentreDescription,ExamCentreId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,Overseas,HomeClosingDate,OverseasClosingDate,ClosingDate"
        Case ExamDataSelectionTypes.dstExamUnitLookup
          mvResultColumns = "ExamUnitCode,ExamUnitDescription,ExamUnitId,ExamUnitLinkId"
          mvSelectColumns = "ExamUnitCode,ExamUnitDescription,ExamUnitId,ExamUnitLinkId"
          mvHeadings = "ExamUnitCode,ExamUnitDescription,ExamUnitId,ExamUnitLinkId"
          mvWidths = "1200,1200,1200,1200"
          mvRequiredItems = "ExamUnitCode,ExamUnitDescription,ExamUnitId,ExamUnitLinkId"
      End Select

      Dim vSelectItems() As String = mvSelectColumns.Split(","c)
      Dim vWidths As New StringBuilder
      For vIndex As Integer = 0 To vSelectItems.Length - 1
        If vIndex > 0 Then vWidths.Append(",")
        vWidths.Append("1000")
      Next
      mvWidths = vWidths.ToString

      If vPrimaryList Then
        If ((pUsage = DataSelectionUsages.dsuWEBServices) Or (pUsage = DataSelectionUsages.dsuSmartClient)) And pListType = DataSelectionListType.dsltEditing Then
          mvResultColumns = mvResultColumns & ",DetailItems,NewColumn,NewColumn2,NewColumn3,Spacer"
        End If
      End If

      Select Case pListType
        Case DataSelectionListType.dsltDefault
          'Do nothing
        Case DataSelectionListType.dsltUser
          ReadUserDisplayListItems(mvEnv.User.Department, mvEnv.User.Logname, pGroup, pUsage)
        Case DataSelectionListType.dsltEditing
          'GetDefaultDisplayListItems()
      End Select
    End Sub

    Public Overrides Function DataTable() As CDBDataTable
      If mvParameters Is Nothing Then mvParameters = New CDBParameters
      Dim vDataTable As New CDBDataTable
      If mvResultColumns.Length > 0 Then vDataTable.AddColumnsFromList(mvResultColumns)
      Select Case mvExamSelectionType
        Case ExamDataSelectionTypes.dstExamBookingUnits
          GetExamBookingUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamCandidateActivites
          GetExamCandidateActivities(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreAssessmentTypes
          GetExamCentreAssessmentTypes(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreActions
          GetExamCentreActions(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreHistory
          GetExamCentreHistory(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreContacts
          GetExamCentreContacts(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreUnits
          GetExamCentreUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreUnitSelection
          GetExamCentreUnitSelection(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentres
          GetExamCentres(vDataTable)
        Case ExamDataSelectionTypes.dstExamExemptions
          GetExamExemptions(vDataTable)
        Case ExamDataSelectionTypes.dstExamExemptionUnits
          GetExamExemptionUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamExemptionUnitSelection
          GetExamExemptionUnitSelection(vDataTable)
        Case ExamDataSelectionTypes.dstExamPersonnel
          GetExamPersonnel(vDataTable)
        Case ExamDataSelectionTypes.dstExamPersonnelAssessTypes
          GetExamPersonnelAssessTypes(vDataTable)
        Case ExamDataSelectionTypes.dstExamPersonnelExpenses
          GetExamPersonnelExpenses(vDataTable)
        Case ExamDataSelectionTypes.dstExamSchedule
          GetExamSchedule(vDataTable)
        Case ExamDataSelectionTypes.dstExamSchedulePersonnel
          GetExamSchedulePersonnel(vDataTable)
        Case ExamDataSelectionTypes.dstExamSessionCentres
          GetExamSessionCentres(vDataTable)
        Case ExamDataSelectionTypes.dstExamSessionCentreSelection
          GetExamSessionCentreSelection(vDataTable)
        Case ExamDataSelectionTypes.dstExamSessions
          GetExamSessions(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentEligibility
          GetExamStudentEligibility(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentExemptionHistory
          GetExamStudentExemptionHistory(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentExemptions
          GetExamStudentExemptions(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentHeader
          GetExamStudentHeader(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentUnitHeader
          GetExamStudentUnitHeader(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitAssessmentTypes
          GetExamUnitAssessmentTypes(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitEligibilityChecks
          GetExamUnitEligibilityChecks(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitGrades
          GetExamUnitGrades(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitLinks
          GetExamUnitLinks(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitPersonnel
          GetExamUnitPersonnel(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitPrerequisites
          GetExamUnitPrerequisites(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitProducts
          GetExamUnitProducts(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnits
          GetExamUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentBookingUnits
          GetExamStudentBookingUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentResults
          GetExamStudentResultEntry(vDataTable)
        Case ExamDataSelectionTypes.dstExamStudentComponentResults
          GetExamStudentResultEntryComponents(vDataTable)
        Case ExamDataSelectionTypes.dstExamScheduleAllCentres
          GetExamScheduleAllCentres(vDataTable)
        Case ExamDataSelectionTypes.dstExamPersonnelMarkerInfo
          GetExamPersonnelMarkerInfo(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitCandidates
          GetExamUnitCandidates(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitMarkerAllocation
          GetExamUnitMarkerAllocation(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitMarkerAllocationList
          GetExamUnitMarkerAllocationList(vDataTable)
        Case ExamDataSelectionTypes.dstExamMarkerList
          GetExamMarkerList(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreCategories
          GetExamCentreCategories(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitLinkCategories
          GetExamUnitLinkCategories(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreUnitLinkCategories
          GetExamCentreUnitLinkCategories(vDataTable)
        Case ExamDataSelectionTypes.dstExamAccreditationHistory
          GetAccreditationHistory(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreUnitDetails
          GetExamUnits(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitGradeHistory, ExamDataSelectionTypes.dstExamUnitHeaderGradeHistory
          GetExamUnitGradeHistory(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreDocuments, ExamDataSelectionTypes.dstExamCentreUnitLinkDocuments, ExamDataSelectionTypes.dstExamUnitLinkDocuments
          GetDocuments(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitStudyModes
          GetUnitStudyModes(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreUnitStudyModes
          GetCentreUnitStudyModes(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitCertRunTypes
          GetUnitCertRunTypes(vDataTable)
        Case ExamDataSelectionTypes.dstExamBookingStudyModes
          GetExamBookingStudyModes(vDataTable)
        Case ExamDataSelectionTypes.dstExamCertReprintTypes
          GetCertReprintTypes(vDataTable)
        Case ExamDataSelectionTypes.dstExamSessionLookup
          GetExamSessionLookup(vDataTable)
        Case ExamDataSelectionTypes.dstExamCentreLookup
          GetExamCentreLookup(vDataTable)
        Case ExamDataSelectionTypes.dstExamUnitLookup
          GetExamUnitLookup(vDataTable)
      End Select
      Return vDataTable
    End Function
    ''' <summary>
    '''  Get categories for Exam Centre
    ''' </summary>
    ''' <param name="pDataTable"> Data Table</param>
    ''' <remarks></remarks>
    Private Sub GetExamCentreCategories(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        Dim vFields As String = "cat.activity,cat.activity_value,cat.quantity,cat.activity_date,cat.source,cat.valid_from,cat.valid_to,cat.amended_by,cat.amended_on,cat.notes,activity_desc,activity_value_desc,source_desc,rgb_value,ctl.exam_centre_id,cat.category_id,'' NoteFlag, ''Status,''StatusOrder"
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("categories cat", "ctl.category_id", "cat.category_id")
        vAnsiJoins.Add("activities a", "cat.activity", "a.activity")
        vAnsiJoins.Add("activity_values av", "cat.activity", "av.activity", "cat.activity_value", "av.activity_value")
        vAnsiJoins.Add("sources s", "cat.source", "s.source")

        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("CategoryId") Then
          vWhereFields.Add("ctl.category_id", mvParameters("CategoryId").Value)
        Else
          If mvParameters.Exists("ExamCentreId") Then vWhereFields.Add("ctl.exam_centre_id", mvParameters("ExamCentreId").Value)
        End If

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "category_links ctl", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ",,")

        Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
        Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")

        If vStatus Then pDataTable.Columns("Status").AttributeName = "status" 'Why
        If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag" 'Why

        For Each vRow As CDBDataRow In pDataTable.Rows
          If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
          If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        Next
      End If
    End Sub
    ''' <summary>
    ''' Get categories for Exam Centre Unit Link 
    ''' </summary>
    ''' <param name="pDataTable">Data table</param>
    ''' <remarks></remarks>
    Private Sub GetExamCentreUnitLinkCategories(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        '"ActivityCode,ActivityValueCode, Quantity, ActivityDate, SourceCode, ValidFrom, ValidTo, AmendedBy, AmendedOn, Notes, ActivityDesc, ActivityValueDesc, SourceDesc, RgbActivityValue, ExamCentreUnitId, CategoryId"
        Dim vFields As String = "cat.activity,cat.activity_value,cat.quantity,cat.activity_date,cat.source,cat.valid_from,cat.valid_to,cat.amended_by,cat.amended_on,cat.notes,activity_desc,activity_value_desc,source_desc,rgb_value,ctl.exam_centre_unit_id,cat.category_id,ctl.exam_unit_link_id,'' NoteFlag, ''Status,''StatusOrder"
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("categories cat", "ctl.category_id", "cat.category_id")
        vAnsiJoins.Add("activities a", "cat.activity", "a.activity")
        vAnsiJoins.Add("activity_values av", "cat.activity", "av.activity", "cat.activity_value", "av.activity_value")
        vAnsiJoins.Add("sources s", "cat.source", "s.source")

        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("CategoryId") Then
          vWhereFields.Add("ctl.category_id", mvParameters("CategoryId").Value)
        Else
          If mvParameters.Exists("ExamCentreUnitId") Then vWhereFields.Add("ctl.exam_centre_unit_id", mvParameters("ExamCentreUnitId").Value)
          If mvParameters.Exists("ExamUnitLinkId") Then vWhereFields.Add("ctl.exam_unit_link_id", mvParameters("ExamUnitLinkId").Value, CDBField.FieldWhereOperators.fwoOR)
        End If

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "category_links ctl", vWhereFields, "ctl.exam_unit_link_id desc", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL)

        Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
        Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")

        If vStatus Then pDataTable.Columns("Status").AttributeName = "status" 'Why
        If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag" 'Why

        For Each vRow As CDBDataRow In pDataTable.Rows
          If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
          If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        Next
      End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pDatatable"></param>
    ''' <remarks></remarks>
    Private Sub GetAccreditationHistory(ByVal pDatatable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        '"AccreditationStatus,AccreditationStatusDesc,ValidFrom,ValidTo,AmendedBy,AmendedOn"
        Dim vFields As String = "eah.accreditation_status,accreditation_status_desc,accreditation_valid_from,accreditation_valid_to,eah.amended_by,eah.amended_on,eah.accreditation_id,ehl.exam_unit_link_id"
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("exam_accreditation_hist_links ehl", "eah.accreditation_id", "ehl.accreditation_id")
        vAnsiJoins.Add("exam_accreditation_statuses eas", "eah.accreditation_status", "eas.accreditation_status")

        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("ExamCentreUnitId") Then vWhereFields.Add("ehl.exam_centre_unit_id", mvParameters("ExamCentreUnitId").Value)
        If mvParameters.Exists("ExamUnitLinkId") Then vWhereFields.Add("ehl.exam_unit_link_id", mvParameters("ExamUnitLinkId").Value)
        If mvParameters.Exists("ExamCentreId") Then vWhereFields.Add("ehl.exam_centre_id", mvParameters("ExamCentreId").Value)

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_accreditation_history eah", vWhereFields, "eah.accreditation_id desc", vAnsiJoins)
        pDatatable.FillFromSQL(mvEnv, vSQL)
      End If
    End Sub


    ''' <summary>
    ''' Get Categories for Exam Unit Link
    ''' </summary>
    ''' <param name="pDataTable">Data Table</param>
    ''' <remarks></remarks>
    Private Sub GetExamUnitLinkCategories(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        Dim vFields As String = "cat.activity,cat.activity_value,cat.quantity,cat.activity_date,cat.source,cat.valid_from,cat.valid_to,cat.amended_by,cat.amended_on,cat.notes,activity_desc,activity_value_desc,source_desc,rgb_value,ctl.exam_unit_link_id,cat.category_id,'' NoteFlag, ''Status,''StatusOrder"
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("categories cat", "ctl.category_id", "cat.category_id")
        vAnsiJoins.Add("activities a", "cat.activity", "a.activity")
        vAnsiJoins.Add("activity_values av", "cat.activity", "av.activity", "cat.activity_value", "av.activity_value")

        vAnsiJoins.Add("sources s", "cat.source", "s.source")

        Dim vWhereFields As New CDBFields()
        If mvParameters.Exists("CategoryId") Then
          vWhereFields.Add("ctl.category_id", mvParameters("CategoryId").Value)
        Else
          If mvParameters.Exists("ExamUnitLinkId") Then vWhereFields.Add("ctl.exam_unit_link_id", mvParameters("ExamUnitLinkId").Value)
        End If

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "category_links ctl", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL)

        Dim vStatus As Boolean = pDataTable.Columns.ContainsKey("Status")
        Dim vNoteFlag As Boolean = pDataTable.Columns.ContainsKey("NoteFlag")

        If vStatus Then pDataTable.Columns("Status").AttributeName = "status" 'Why
        If vNoteFlag Then pDataTable.Columns("NoteFlag").AttributeName = "note_flag" 'Why

        For Each vRow As CDBDataRow In pDataTable.Rows
          If vNoteFlag AndAlso vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
          If vStatus Then vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        Next
      End If
    End Sub

    Private Sub GetExamBookingUnits(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamBookingUnitId") Then vWhereFields.Add("exam_booking_unit_id", mvParameters("ExamBookingUnitId").IntegerValue)
      If mvParameters.HasValue("ExamBookingId") Then vWhereFields.Add("exam_booking_id", mvParameters("ExamBookingId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ExamScheduleId") Then vWhereFields.Add("exam_schedule_id", mvParameters("ExamScheduleId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      Dim vFields As String = "exam_booking_unit_id,exam_booking_id," & mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & ",ebu.exam_unit_id,exam_unit_code,exam_unit_description,exam_schedule_id,exam_personnel_id,exam_candidate_number,desk_number,batch_number,transaction_number,line_number,attempt_number,exam_student_unit_status,original_mark,moderated_mark,total_mark,original_grade,moderated_grade,total_grade,original_result,moderated_result,total_result,entry_date,expiry_date,done_date,expiry_session,eu.activity_group,ebu.created_by,ebu.created_on,ebu.amended_by,ebu.amended_on"
      vAnsiJoins.Add("exam_units eu", "ebu.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "ebu.exam_unit_id", "eul.exam_unit_id_2")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_booking_units ebu", vWhereFields, mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & "," & mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub


    Private Sub GetExamCentreAssessmentTypes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamCentreAssessmentTypeId") Then vWhereFields.Add("exam_centre_assessment_type_id", mvParameters("ExamCentreAssessmentTypeId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      Dim vFields As String = "exam_centre_assessment_type_id,exam_centre_id,ecat.exam_assessment_type,exam_assessment_type_desc,ecat.created_by,ecat.created_on,ecat.amended_by,ecat.amended_on"
      vAnsiJoins.Add("exam_assessment_types eat", "ecat.exam_assessment_type", "eat.exam_assessment_type")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_centre_assessment_types ecat", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamCentreActions(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.HasValue("ActionNumber") Then
        vWhereFields.Add("eca.action_number", mvParameters("ActionNumber").IntegerValue)
      End If
      If mvParameters.HasValue("ExamCentreId") Then
        vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection,
                                   "act.master_action,act.action_level,act.sequence_number,act.action_number,act.action_desc,actp.action_priority_desc,acts.action_status_desc,act.created_by,act.created_on,act.deadline,act.scheduled_on,act.completed_on,actp.action_priority,acts.action_status,'R' AS type,'R' AS type,acts.action_status,act.duration_days,act.duration_hours,act.duration_minutes,act.document_class,act.action_text",
                                   "action_links eca",
                                   vWhereFields,
                                   "",
                                   New AnsiJoins({New AnsiJoin("actions act", "eca.action_number", "act.action_number"),
                                                  New AnsiJoin("action_priorities actp", "act.action_priority", "actp.action_priority"),
                                                  New AnsiJoin("action_statuses acts", "act.action_status", "acts.action_status")}))
      pDataTable.FillFromSQL(mvEnv, vSQL, "master_action,action_level,sequence_number,a.action_number,action_desc,action_priority_desc,action_status_desc,a.created_by,a.created_on,deadline,scheduled_on,completed_on,a.action_priority,a.action_status,alk.type,alk.type AS link_type_description,a.action_status AS sort_column,,,,,,,,,,,duration_days,duration_hours,duration_minutes,a.document_class,action_text,outlook_id")
      GetLookupData(pDataTable, "LinkType", "contact_actions", "type")
      GetActionersAndSubjects(pDataTable)
    End Sub

    Private Sub GetExamCentreHistory(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      If mvParameters.HasValue("ExamCentreId") Then
        vWhereFields.Add(ExamCentreHistory.AliasedColumnName(ExamCentreHistory.ColumnId.ExamCentreId), mvParameters("ExamCentreId").IntegerValue)
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection,
                                   ExamCentreHistory.GetAliasedColumnNameList(mvEnv).AsCommaSeperated,
                                   ExamCentreHistory.Table & " " & ExamCentreHistory.ShortName,
                                   vWhereFields,
                                   ExamCentreHistory.AliasedColumnName(ExamCentreHistory.ColumnId.ExamCentreDescriptionTimestamp) & " DESC")
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamCentreContacts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamCentreContactId") Then vWhereFields.Add("exam_centre_contact_id", mvParameters("ExamCentreContactId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")
      vAnsiJoins.Add("contacts c", "ecc.contact_number", "c.contact_number")
      vAnsiJoins.Add("exam_contact_types ect", "ecc.exam_contact_type", "ect.exam_contact_type")
      Dim vFields As String = "exam_centre_contact_id,exam_centre_id,ecc.contact_number,ecc.exam_contact_type,exam_contact_type_desc,ecc.created_by,ecc.created_on,ecc.amended_by,ecc.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_centre_contacts ecc", vWhereFields, "surname,forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)
    End Sub

    Private Sub GetExamCentreUnits(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamCentreUnitId") Then vWhereFields.Add("exam_centre_unit_id", mvParameters("ExamCentreUnitId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vFields As String = "exam_centre_unit_id,exam_centre_id,exam_unit_id,created_by,created_on,amended_by,amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then vFields = vFields + ",exam_unit_link_id"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_centre_units", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamCentreUnitSelection(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins

      If mvParameters.HasValue("ExamSessionId") Then
        vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      Else
        vWhereFields.Add("exam_session_id", CDBField.FieldTypes.cftInteger)
      End If
      vWhereFields.Add("exam_question", "N")
      Dim vFields As String = "exam_centre_unit_id,exam_centre_id,exam_unit_id,exam_unit_link_id,created_by,created_on,amended_by,amended_on,local_name"
      Dim vExamCentreFields As String = "exam_centre_unit_id,exam_centre_id,exam_unit_id,exam_unit_link_id,created_by,created_on,amended_by,amended_on,local_name,accreditation_status,accreditation_valid_from,accreditation_valid_to"
      Dim vUnitCentreSelect As New SQLStatement(mvEnv.Connection, vExamCentreFields, "exam_centre_units", New CDBField("exam_centre_id", mvParameters("ExamCentreId").IntegerValue))
      vUnitCentreSelect.UseAnsiSQL = True

      vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")

      'Assessment Type
      If mvEnv.GetConfigOption("ex_restrict_courses") Then
        'If the config option exam_restrict_centre_units is set then
        Dim vAssessWhere As New CDBFields
        vAssessWhere.Add("euat.exam_assessment_type")
        vAssessWhere.Add("euat.exam_assessment_type#2", CDBField.FieldTypes.cftInteger, "ecat.exam_assessment_type", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracketTwice Or CDBField.FieldWhereOperators.fwoCloseBracket)
        vAssessWhere.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue, CDBField.FieldWhereOperators.fwoCloseBracket)
        Dim vAssessJoins As New AnsiJoins
        vAssessJoins.AddLeftOuterJoin("exam_unit_assessment_types euat", "eu.exam_unit_id", "euat.exam_unit_id")
        vAssessJoins.AddLeftOuterJoin("exam_centre_assessment_types ecat", "euat.exam_assessment_type", "ecat.exam_assessment_type")
        Dim vAssessSQL As New SQLStatement(mvEnv.Connection, "DISTINCT eu.exam_unit_id", "exam_units eu", vAssessWhere, "", vAssessJoins)
        vAnsiJoins.Add(String.Format("( {0} ) eat", vAssessSQL.SQL), "eu.exam_unit_id", "eat.exam_unit_id")
      End If

      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")

      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) ecu", vUnitCentreSelect.SQL), "ecu.exam_unit_id", "eu.exam_unit_id", "ecu.exam_unit_link_id", "eul.exam_unit_link_id")
      vFields = mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & ",exam_centre_unit_id,exam_centre_id,eu.exam_unit_id,exam_unit_code,exam_unit_description"

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        vFields = vFields + ",eul.exam_unit_link_id,eul.parent_unit_link_id,ecu.local_name,ecu.accreditation_status,ecu.accreditation_valid_from,ecu.accreditation_valid_to,ecu.created_by,ecu.created_on,ecu.amended_by,ecu.amended_on"
      Else
        vFields = vFields + ",ecu.created_by,ecu.created_on,ecu.amended_by,ecu.amended_on"
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_units eu", vWhereFields, mvEnv.Connection.DBIsNull("eul.parent_unit_link_id", "0") & "," & mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description", vAnsiJoins)

      vSQL.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQL)

      If mvParameters.ContainsKey("Trader") OrElse mvParameters.ContainsKey("ResultEntry") Then
        If mvParameters.HasValue("ExamCentreId") Then
          Dim vFilterType As ExamAccreditationFilterTypes = ExamAccreditationFilterTypes.eafAllowRegistration
          If mvParameters.ParameterExists("ResultEntry").Bool Then vFilterType = ExamAccreditationFilterTypes.eafAllowResultEntry

          Dim vRemoveIds As New List(Of Integer)
          'Build a dictionary of all exam unit links and their parents.  We will gradually filter-out all un-accredited units and then remove all from the final list
          Dim vExamUnitLinks As New Dictionary(Of Integer, Integer)
          pDataTable.Rows.ForEach(Sub(vRow) vExamUnitLinks.Add(vRow.IntegerItem("ExamUnitLinkId"), vRow.IntegerItem("ParentUnitLinkId")))

          Dim vAccreditationFilter As New ExamAccreditationFilter(mvEnv, vExamUnitLinks, vFilterType) 'This class handles all the accreditation filtering
          'Filter-out unaccredited Centres and unaccredited centre-units
          Dim vSessionID As Integer = mvParameters.ParameterExists("ExamSessionId").IntegerValue
          'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
          Dim vCentreAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation))
          If vCentreAccreditationIsEnabled Then
            Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnitsAtUnaccreditedCentre(mvParameters("ExamCentreId").IntegerValue, vSessionID)
            vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
          End If

          'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
          Dim vCentreUnitAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation))
          If vCentreUnitAccreditationIsEnabled Then
            Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentreUnits(mvParameters("ExamCentreId").IntegerValue, vSessionID)
            vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
          End If

          If vRemoveIds IsNot Nothing AndAlso vRemoveIds.Count > 0 Then
            Dim vRemoveRows As List(Of CDBDataRow) = pDataTable.Rows.FindAll(Function(vRow) vRemoveIds.Contains(vRow.IntegerItem("ExamUnitLinkId")))
            vRemoveRows.ForEach(Sub(vRow) pDataTable.RemoveRow(vRow))
          End If
        End If
      End If
    End Sub

    Private Sub GetExamCentres(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamCentreParentId") Then vWhereFields.Add("exam_centre_parent_id", mvParameters("ExamCentreParentId").IntegerValue)
      Dim vFields As String = "exam_centre_id,organisation_number,address_number,contact_number,exam_centre_code,exam_centre_description,valid_from,valid_to,capacity,last_visit_date,next_visit_date,exam_centre_parent_id,additional_capacity,accept_special_requirements,overseas,web_publish,exam_centre_rate_type,created_by,created_on,amended_by,amended_on"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then vFields = vFields + ",accreditation_status,accreditation_valid_from,accreditation_valid_to"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_centres", vWhereFields, mvEnv.Connection.DBIsNull("exam_centre_parent_id", "0") & ",exam_centre_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamExemptions(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamExemptionId") Then vWhereFields.Add("exam_exemption_id", mvParameters("ExamExemptionId").IntegerValue)
      Dim vFields As String = "exam_exemption_id,exam_exemption_code,exam_exemption_description,product,rate,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_exemptions", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamExemptionUnits(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamExemptionUnitId") Then vWhereFields.Add("exam_exemption_unit_id", mvParameters("ExamExemptionUnitId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("eeu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ExamExemptionId") Then vWhereFields.Add("eeu.exam_exemption_id", mvParameters("ExamExemptionId").IntegerValue)
      Dim vFields As String = "exam_exemption_unit_id,eeu.exam_unit_id,eeu.exam_exemption_id,exam_unit_code,exam_unit_description,exam_exemption_code,exam_exemption_description,ee.product,product_desc,ee.rate,rate_desc,eeu.created_by,eeu.created_on,eeu.amended_by,eeu.amended_on"
      vAnsiJoins.Add("exam_units eu", "eeu.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_exemptions ee", "eeu.exam_exemption_id", "ee.exam_exemption_id")
      vAnsiJoins.Add("products p", "ee.product", "p.product")
      vAnsiJoins.Add("rates r", "ee.product", "r.product", "ee.rate", "r.rate")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_exemption_units eeu", vWhereFields, "exam_unit_code", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamExemptionUnitSelection(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamSessionId") Then
        vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      Else
        vWhereFields.Add("exam_session_id", CDBField.FieldTypes.cftInteger)
      End If
      vWhereFields.Add("allow_exemptions", "Y")
      vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      Dim vFields As String = "exam_exemption_unit_id,exam_exemption_id,exam_unit_id,created_by,created_on,amended_by,amended_on"
      Dim vUnitExemptionSelect As New SQLStatement(mvEnv.Connection, vFields, "exam_exemption_units", New CDBField("exam_exemption_id", mvParameters("ExamExemptionId").IntegerValue))
      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) ecu", vUnitExemptionSelect.SQL), "eu.exam_unit_id", "ecu.exam_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")
      vFields = mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & ",exam_exemption_unit_id,exam_exemption_id,eu.exam_unit_id,exam_unit_code,exam_unit_description,eul.exam_unit_link_id,eul.parent_unit_link_id,ecu.created_by,ecu.created_on,ecu.amended_by,ecu.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_units eu", vWhereFields, mvEnv.Connection.DBIsNull("parent_unit_link_id", "0") & "," & mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamPersonnel(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      If mvParameters.HasValue("ContactNumber") Then vWhereFields.Add("ep.contact_number", mvParameters("ContactNumber").IntegerValue)
      If mvParameters.HasValue("IsCurrent") AndAlso mvParameters("IsCurrent").Value = "Y" Then
        vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      End If
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")
      vAnsiJoins.Add("contacts c", "ep.contact_number", "c.contact_number")
      vAnsiJoins.Add("exam_personnel_types ept", "ep.exam_personnel_type", "ept.exam_personnel_type")
      Dim vFields As String = "exam_personnel_id,ep.contact_number,valid_from,valid_to,ep.exam_personnel_type,exam_personnel_type_desc,exam_marker,trained_date,maximum_students,ep.notes,ep.created_by,ep.created_on,ep.amended_by,ep.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_personnel ep", vWhereFields, "surname,forenames,exam_personnel_id desc", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)
    End Sub

    Private Sub GetExamPersonnelAssessTypes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamPersonnelAssessTypeId") Then vWhereFields.Add("exam_personnel_assess_type_id", mvParameters("ExamPersonnelAssessTypeId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      Dim vFields As String = "exam_personnel_assess_type_id,exam_personnel_id,epat.exam_assessment_type,exam_assessment_type_desc,epat.created_by,epat.created_on,epat.amended_by,epat.amended_on"
      vAnsiJoins.Add("exam_assessment_types eat", "epat.exam_assessment_type", "eat.exam_assessment_type")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_personnel_assess_types epat", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamPersonnelExpenses(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamPersonnelExpenseId") Then vWhereFields.Add("exam_personnel_expense_id", mvParameters("ExamPersonnelExpenseId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      Dim vFields As String = "exam_personnel_expense_id,exam_personnel_id,exam_expense_type,amount,applied_date,paid_date,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_personnel_expenses", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamSchedule(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamScheduleId") Then vWhereFields.Add("exam_schedule_id", mvParameters("ExamScheduleId").IntegerValue)
      If mvParameters.HasValue("ExamSessionId") Then vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vFields As String = "exam_schedule_id,exam_session_id,es.exam_centre_id,exam_unit_id,exam_centre_code,exam_centre_description,start_date,start_time,end_time,es.capacity,number_of_candidates,es.additional_capacity,es.created_by,es.created_on,es.amended_by,es.amended_on"
      vAnsiJoins.Add("exam_centres ec", "es.exam_centre_id", "ec.exam_centre_id")

      Dim vSubAnsiJoins As New AnsiJoins
      Dim vSubWhereFields As New CDBFields
      If mvParameters.HasValue("ExamUnitId") Then vSubWhereFields.Add("ebu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      vSubWhereFields.Add("eb.cancellation_reason", "")
      vSubAnsiJoins.Add("exam_bookings eb", "ebu.exam_booking_id", "eb.exam_booking_id")
      Dim vSubSelect As New SQLStatement(mvEnv.Connection, "exam_centre_id, count(*) as number_of_candidates", "exam_booking_units ebu", vSubWhereFields, "", vSubAnsiJoins)
      vSubSelect.GroupBy = "exam_centre_id"
      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) nc", vSubSelect.SQL), "ec.exam_centre_id", "nc.exam_centre_id")

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_schedule es", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamScheduleAllCentres(ByVal pDataTable As CDBDataTable)
      Dim vSQL As String =
          <ExamTreeSQL>
            WITH unit_centres
            AS
            (
              SELECT exam_schedule_id,esc.exam_session_id,ec.exam_centre_id,ec.exam_centre_parent_id,eu.exam_base_unit_id exam_unit_id,ec.exam_centre_code,exam_centre_description,start_date,start_time,end_time,es.capacity,es.additional_capacity,es.created_by,es.created_on,es.amended_by,es.amended_on
              FROM exam_units eu
	                JOIN exam_session_centres esc
		                  ON esc.exam_session_id = eu.exam_session_id
	                JOIN exam_centres ec
		                  ON ec.exam_centre_id = esc.exam_centre_id
  	              LEFT JOIN exam_schedule es
	    	              ON es.exam_session_id = eu.exam_session_id
                          AND es.exam_unit_id = eu.exam_unit_id
                          AND es.exam_centre_id = ec.exam_centre_id
              {0}{1} --WHERE CLAUSE ADDED HERE LATER
         		      AND EXISTS 
		          (
			         SELECT *
			         FROM exam_centre_units sub_ecu
					        JOIN exam_unit_links sub_eul
						        ON sub_eul.base_unit_link_id = sub_ecu.exam_unit_link_id
			         WHERE sub_ecu.exam_centre_id = esc.exam_centre_id
					        AND sub_eul.exam_unit_id_2 = eu.exam_unit_id
		          )
            )
            ,
            parent_centres AS
            (
                   SELECT DISTINCT exam_centre_parent_id
                   FROM unit_centres
                   WHERE NOT EXISTS
                   (
                          SELECT NULL
                          FROM unit_centres parent_also_offers_unit
                          WHERE parent_also_offers_unit.exam_centre_id = unit_centres.exam_centre_parent_id
                   )
            )
            ,
            ancestor_centres AS
            (
            SELECT exam_centres.exam_centre_id, exam_centres.exam_centre_parent_id, exam_centres.exam_centre_code, exam_centres.exam_centre_description
            FROM exam_centres
                   JOIN parent_centres
                          ON parent_centres.exam_centre_parent_id = exam_centres.exam_centre_id
            UNION ALL
                   SELECT parent_centre.exam_centre_id, parent_centre.exam_centre_parent_id, parent_centre.exam_centre_code, parent_centre.exam_centre_description
                   FROM exam_centres parent_centre
                          JOIN ancestor_centres
                                 ON ancestor_centres.exam_centre_parent_id = parent_centre.exam_centre_id
                    WHERE NOT EXISTS
                    (
                      SELECT null
                      FROM unit_centres parent_also_offers_unit
                      WHERE parent_also_offers_unit.exam_centre_id = parent_centre.exam_centre_id
                    )
            )
            SELECT exam_schedule_id,exam_session_id,exam_centre_id,exam_centre_parent_id,exam_unit_id,exam_centre_code,exam_centre_description,start_date
                                 ,start_time,end_time,capacity,additional_capacity,created_by,created_on,amended_by,amended_on,COALESCE(exam_centre_parent_id, 0)
            FROM unit_centres
            UNION ALL
            SELECT DISTINCT null,null,ac.exam_centre_id,ac.exam_centre_parent_id,null,exam_centre_code,exam_centre_description,null,null,null,null,null,null,null,null,null,COALESCE(exam_centre_parent_id, 0)
            FROM ancestor_centres ac
            ORDER BY COALESCE(exam_centre_parent_id, 0) ,exam_centre_description
          </ExamTreeSQL>.Value
      vSQL = String.Format(String.Format(vSQL, If(mvParameters.HasValue("ExamSessionId"), String.Format("WHERE esc.exam_session_id = {0}", mvParameters("ExamSessionId").IntegerValue.ToString), ""), If(mvParameters.HasValue("ExamUnitId"), String.Format("{0} eu.exam_unit_id = {1}", "{0}", mvParameters("ExamUnitId").IntegerValue.ToString), "")), If(mvParameters.HasValue("ExamSessionId"), " AND", "WHERE"))
      pDataTable.FillFromSQLDONOTUSE(mvEnv, vSQL)
    End Sub

    Private Sub GetExamPersonnelMarkerInfo(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamSessionId") Then vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      Dim vFields As String = "exam_personnel_id,em.exam_session_id,em.exam_centre_id,em.exam_unit_id,exam_session_code,exam_session_description,exam_unit_code,exam_unit_description,exam_centre_code,exam_centre_description,number_of_papers,marker_number"
      Dim vSubSelect As New SQLStatement(mvEnv.Connection, "exam_personnel_id,exam_session_id, exam_unit_id, exam_centre_id, count(*) as number_of_papers, marker_number", "exam_marking_batch_detail", vWhereFields)
      vSubSelect.GroupBy = "exam_personnel_id,exam_session_id, exam_unit_id, exam_centre_id, marker_number"
      vAnsiJoins.Add(String.Format("({0} ) em", vSubSelect.SQL), "em.exam_session_id", "es.exam_session_id")
      vAnsiJoins.Add("exam_centres ec", "em.exam_centre_id", "ec.exam_centre_id")
      vAnsiJoins.Add("exam_units eu", "em.exam_unit_id", "eu.exam_unit_id")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_sessions es", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamUnitCandidates(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("ebu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vFields As String = "eb.exam_booking_id,eb.exam_centre_id,ebu.exam_unit_id,exam_schedule_id,exam_centre_code,exam_centre_description,exam_candidate_number"
      vAnsiJoins.Add("exam_booking_units ebu", "eb.exam_booking_id", "ebu.exam_booking_id")
      vAnsiJoins.Add("exam_centres ec", "eb.exam_centre_id", "ec.exam_centre_id")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_bookings eb", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamSchedulePersonnel(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamSchedulePersonnelId") Then vWhereFields.Add("exam_schedule_personnel_id", mvParameters("ExamSchedulePersonnelId").IntegerValue)
      If mvParameters.HasValue("ExamScheduleId") Then vWhereFields.Add("exam_schedule_id", mvParameters("ExamScheduleId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      Dim vFields As String = "exam_schedule_personnel_id,exam_schedule_id,exam_personnel_id,exam_personnel_type,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_schedule_personnel", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamSessionCentres(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamSessionCentreId") Then vWhereFields.Add("exam_session_centre_id", mvParameters("ExamSessionCentreId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamSessionId") Then vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      Dim vFields As String = "exam_session_centre_id,exam_centre_id,exam_session_id,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_session_centres", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamSessionCentreSelection(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins

      vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      Dim vFields As String = "exam_session_centre_id,exam_session_id,exam_centre_id"
      Dim vSessionCentreSelect As New SQLStatement(mvEnv.Connection, vFields, "exam_session_centres", New CDBField("exam_session_id", mvParameters("ExamSessionId").IntegerValue))
      vAnsiJoins.AddLeftOuterJoin(String.Format("({0}) esc", vSessionCentreSelect.SQL), "ec.exam_centre_id", "esc.exam_centre_id")
      vFields = "exam_session_centre_id,exam_centre_parent_id,ec.exam_centre_id,exam_session_id,exam_centre_code,exam_centre_description"

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_centres ec", vWhereFields, mvEnv.Connection.DBIsNull("exam_centre_parent_id", "0") + ",exam_centre_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamSessions(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamSessionId") Then vWhereFields.Add("exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      Dim vFields As String = "exam_session_id,exam_session_year,exam_session_month,exam_session_code,exam_session_description,sequence_number,"
      vFields &= "valid_from,valid_to,home_closing_date,overseas_closing_date,web_publish,notes,created_by,created_on,amended_by,amended_on,"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then vFields &= "results_release_date"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "exam_sessions", vWhereFields, mvEnv.Connection.DBIsNull("sequence_number", "10000") & ",exam_session_year DESC,exam_session_month DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields)
    End Sub

    Private Sub GetExamStudentEligibility(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitEligibilityCheckId") Then vWhereFields.Add("exam_unit_eligibility_check_id", mvParameters("ExamUnitEligibilityCheckId").IntegerValue)
      Dim vFields As String = "exam_unit_eligibility_check_id,contact_number,proven,proved_date,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_student_eligibility", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamStudentExemptionHistory(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamStudentExemptionHistId") Then vWhereFields.Add("exam_student_exemption_hist_id", mvParameters("ExamStudentExemptionHistId").IntegerValue)
      If mvParameters.HasValue("ExamStudentExemptionId") Then vWhereFields.Add("exam_student_exemption_id", mvParameters("ExamStudentExemptionId").IntegerValue)
      Dim vFields As String = "exam_student_exemption_hist_id,exam_student_exemption_id,eseh.exam_exemption_status,exam_exemption_status_desc,status_date,eseh.created_by,eseh.created_on,eseh.amended_by,eseh.amended_on"
      vAnsiJoins.Add("exam_exemption_statuses ees", "eseh.exam_exemption_status", "ees.exam_exemption_status")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_student_exemption_history eseh", vWhereFields, "status_date DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamStudentExemptions(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamStudentExemptionId") Then vWhereFields.Add("exam_student_exemption_id", mvParameters("ExamStudentExemptionId").IntegerValue)
      If mvParameters.HasValue("ExamExemptionProductId") Then vWhereFields.Add("exam_exemption_product_id", mvParameters("ExamExemptionProductId").IntegerValue)
      Dim vFields As String = "exam_student_exemption_id,exam_exemption_product_id,contact_number,batch_number,transaction_number,line_number,exam_exemption_status,status_date,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_student_exemptions", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamStudentHeader(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamStudentId") Then vWhereFields.Add("exam_student_id", mvParameters("ExamStudentId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ContactNumber") Then vWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
      Dim vFields As String = "exam_student_id,exam_unit_id,contact_number,first_session_id,last_session_id,last_marked_date,last_graded_date,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_student_header", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamStudentUnitHeader(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamStudentUnitHeaderId") Then vWhereFields.Add("esuh.exam_student_unit_header_id", mvParameters("ExamStudentUnitHeaderId").IntegerValue)
      If mvParameters.HasValue("ExamStudentHeaderId") Then vWhereFields.Add("esuh.exam_student_header_id", mvParameters("ExamStudentHeaderId").IntegerValue)
      If mvParameters.HasValue("ContactNumber") Then
        vAnsiJoins.Add("exam_student_header esh", "esh.exam_student_header_id", "esuh.exam_student_header_id")
        vWhereFields.Add("esh.contact_number", mvParameters("ContactNumber").IntegerValue)
      End If

      Dim vAttrs As String = "esuh.exam_student_unit_header_id,esuh.exam_student_header_id,{0},esuh.exam_unit_id"
      vAttrs &= ",eu.exam_unit_code,eu.exam_unit_description,esuh.attempts,esuh.current_mark,esuh.current_grade,esuh.current_result,"
      vAttrs &= "{1},eg.grade_is_pass,esuh.expires,eul.exam_unit_link_id,eul.parent_unit_link_id,esuh.created_by,"
      vAttrs &= "esuh.created_on,esuh.amended_by,esuh.amended_on,{2}"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vAttrs &= ",esuh.results_release_date,esuh.previous_mark,esuh.previous_grade,esuh.previous_result"
      Else
        vAttrs &= ",,,,"
      End If
      Dim vFields As String = String.Format(vAttrs, "exam_unit_id_1", "exam_grade_sequence_number", "can_edit_results")
      vAttrs = RemoveBlankItems(String.Format(vAttrs, mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & " AS exam_unit_id_1", "eg.sequence_number AS exam_grade_sequence_number", "'Y' as can_edit_results"))


      vAnsiJoins.Add("exam_units eu", "esuh.exam_unit_id", "eu.exam_unit_id")
      vAnsiJoins.Add("exam_unit_links eul", "esuh.exam_unit_id", "eul.exam_unit_id_2", "esuh.exam_unit_link_id", "eul.exam_unit_link_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "esuh.current_grade", "eg.exam_grade")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "exam_student_unit_header esuh", vWhereFields, mvEnv.Connection.DBIsNull("eul.parent_unit_link_id", "0") & "," & mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields)

      RestrictExamResults(pDataTable, "Current")

    End Sub

    Private Sub GetExamUnitAssessmentTypes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitAssessmentTypeId") Then vWhereFields.Add("exam_unit_assessment_type_id", mvParameters("ExamUnitAssessmentTypeId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vFields As String = "exam_unit_assessment_type_id,exam_unit_id,euat.exam_assessment_type,exam_assessment_type_desc,euat.created_by,euat.created_on,euat.amended_by,euat.amended_on"
      vAnsiJoins.Add("exam_assessment_types eat", "euat.exam_assessment_type", "eat.exam_assessment_type")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_assessment_types euat", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamUnitEligibilityChecks(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitEligibilityCheckId") Then vWhereFields.Add("exam_unit_eligibility_check_id", mvParameters("ExamUnitEligibilityCheckId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vFields As String = "exam_unit_eligibility_check_id,exam_unit_id,eligibility_check_text,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_eligibility_checks", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
      pDataTable.Columns("EligibilityCheckText").FieldType = CDBField.FieldTypes.cftMemo
    End Sub

    Private Sub GetExamUnitGrades(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins

      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "eug.exam_unit_id")

      If mvParameters.HasValue("ExamUnitGradeId") Then
        vWhereFields.Add("exam_unit_grade_id", mvParameters("ExamUnitGradeId").IntegerValue)
      ElseIf mvParameters.HasValue("ExamUnitId") Then
        vWhereFields.Add("eu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      ElseIf mvParameters.HasValue("ExamSessionId") Then
        vWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      ElseIf mvParameters.HasValue("ExamSessionCode") AndAlso mvParameters("ExamSessionCode").Value <> "NONSESSION" Then
        vWhereFields.Add("es.exam_session_code", mvParameters("ExamSessionCode").Value)
        vAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "eu.exam_session_id")
      Else
        vWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
      End If

      vAnsiJoins.Add("exam_grades eg", "eg.exam_grade", "eug.exam_grade")
      vAnsiJoins.Add("maintenance_lookup ml", "ml.lookup_code", "eug.exam_grade_condition_type")

      vWhereFields.Add("ml.table_name", "exam_unit_grades")
      vWhereFields.Add("ml.attribute_name", "exam_grade_condition_type")

      Dim vFields As String = "exam_unit_grade_id,eug.exam_unit_id,exam_unit_code,eug.exam_grade,eg.exam_grade_desc,eug.sequence_number,condition_number,clause_number, eug.exam_grade AS condition_desc ,exam_grade_condition_type,lookup_desc AS exam_grade_condition_type_desc,grade_units,exam_grade_operator,required_value,eug.created_by,eug.created_on,eug.amended_by,eug.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_grades eug", vWhereFields, "eu.exam_unit_id,eug.sequence_number desc,condition_number,clause_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)

      Dim LastGradeAndUnit As String = ""
      Dim LastCondition As Integer = -1
      For Each vRow As CARE.Access.CDBDataRow In pDataTable.Rows
        If (LastGradeAndUnit <> vRow.Item("ExamGrade") & vRow.Item("ExamUnitId")) Then
          vRow.Item("ConditionDesc") = "IF"
          LastGradeAndUnit = vRow.Item("ExamGrade") & vRow.Item("ExamUnitId")
          LastCondition = vRow.IntegerItem("ConditionNumber")
        ElseIf (LastCondition <> vRow.IntegerItem("ConditionNumber")) Then
          vRow.Item("ConditionDesc") = "OR"
          LastCondition = vRow.IntegerItem("ConditionNumber")
        Else
          vRow.Item("ConditionDesc") = "   AND"
        End If
      Next
    End Sub

    Private Sub GetExamUnitLinks(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vFields As String = "exam_unit_id_1,exam_unit_id_2,created_by,created_on,amended_by,amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_links", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamUnitPersonnel(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitPersonnelId") Then vWhereFields.Add("exam_unit_personnel_id", mvParameters("ExamUnitPersonnelId").IntegerValue)
      If mvParameters.HasValue("ExamPersonnelId") Then vWhereFields.Add("exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      vAnsiJoins.Add("exam_personnel ep", "ep.exam_personnel_id", "eup.exam_personnel_id")
      vAnsiJoins.Add("contacts c", "ep.contact_number", "c.contact_number")
      vAnsiJoins.Add("exam_personnel_types ept", "ep.exam_personnel_type", "ept.exam_personnel_type")
      vAnsiJoins.AddLeftOuterJoin("exam_marker_options emo", "eup.exam_marker_option", "emo.exam_marker_option")
      Dim vFields As String = "exam_unit_personnel_id,eup.exam_personnel_id,exam_unit_id,eup.valid_from,eup.valid_to,ep.exam_personnel_type,exam_personnel_type_desc,eup.maximum_students,geographical_region,eup.exam_marker_option,exam_marker_option_desc,actual_load_size,eup.created_by,eup.created_on,eup.amended_by,eup.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_unit_personnel eup", vWhereFields, "surname,forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, "contact_number," & ContactNameItems())
    End Sub

    Private Sub GetExamUnitPrerequisites(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("eup.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ExamPrerequisiteUnitId") Then vWhereFields.Add("eup.exam_prerequisite_unit_id", mvParameters("ExamPrerequisiteUnitId").IntegerValue)
      Dim vFields As String = "eup.exam_unit_id,eup.exam_prerequisite_unit_id,eu.exam_unit_code exam_prereq_unit_code,eu.exam_unit_description exam_prereq_unit_description,eup.minimum_grade,eg.exam_grade_desc,eup.pass_required,eup.created_by,eup.created_on,eup.amended_by,eup.amended_on"
      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "eup.exam_prerequisite_unit_id")
      vAnsiJoins.AddLeftOuterJoin("exam_grades eg", "eg.exam_grade", "eup.minimum_grade")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_prerequisites eup", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)

      For Each vDataRow As CDBDataRow In pDataTable.Rows
        vDataRow.SetYNValue("PassRequired", True)
      Next vDataRow
    End Sub

    Private Sub GetExamUnitProducts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitProductId") Then vWhereFields.Add("exam_unit_product_id", mvParameters("ExamUnitProductId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      vAnsiJoins.Add("products p", "eup.product", "p.product")
      vAnsiJoins.Add("rates r", "eup.product", "r.product", "eup.rate", "r.rate")
      Dim vFields As String = "exam_unit_product_id,exam_unit_id,eup.product,product_desc,eup.rate,rate_desc,quantity,eup.created_by,eup.created_on,eup.amended_by,eup.amended_on"
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_unit_products eup", vWhereFields, "", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub

    Private Sub GetExamUnits(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitId") Then
        vWhereFields.Add("eu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Else
        If mvParameters.HasValue("ExamSessionId") Then
          vWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
        ElseIf mvParameters.HasValue("ExamSessionCode") Then
          vWhereFields.Add("es.exam_session_code", mvParameters("ExamSessionCode").Value)
          vAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "eu.exam_session_id")
        Else
          vWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
        End If
      End If

      If mvParameters.HasValue("SessionBased") Then vWhereFields.Add("eu.session_based", mvParameters("SessionBased").Value)
      If mvParameters.HasValue("ValidTo") Then vWhereFields.Add("eu.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ValidTo").Value, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      If mvParameters.HasValue("ValidFrom") Then vWhereFields.Add("eu.valid_from", CDBField.FieldTypes.cftDate, mvParameters("ValidFrom").Value, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      If mvParameters.HasValue("UValidTo_SValidFrom") Then vWhereFields.Add("eu.valid_to#1", CDBField.FieldTypes.cftDate, mvParameters("UValidTo_SValidFrom").Value, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      If mvParameters.HasValue("ExamUnitReplacedById") Then vWhereFields.Add("eu.exam_unit_replaced_by_id", mvParameters("ExamUnitReplacedById").IntegerValue)
      If mvParameters.HasValue("IsGradeApply") Then vWhereFields.Add("Exam_Question", CDBField.FieldTypes.cftCharacter, "Y", CDBField.FieldWhereOperators.fwoNotEqual)
      If mvParameters.HasValue("ExamUnitLinkId") Then vWhereFields.Add("eul.exam_unit_link_id", mvParameters("ExamUnitLinkId").IntegerValue)

      'mvResultColumns = "ExamUnitId1,ExamUnitId,ExamBaseUnitId,ExamSessionId,ExamUnitCode,ExamUnitDescription,Subject,SkillLevel,ExamUnitType,ScheduleRequired,MarkerRequired,ExamQuestion,ExamUnitStatus,SequenceNumber,SessionBased,ValidFrom,ValidTo,Product,Rate,DateApproved,RegistrationDate,QcfLevel,NumberOfCredits,NvqCode,SvqCode,UnitTimeLimit,TimeLimitType,MinimumStudents,MaximumStudents,StudentCount,MinimumAge,AllowBookings,ExamMarkType,MarkFactor,AwardingBody,ExamUnitReplacedById,AllowExemptions,ExemptionMark,ExamMarkerStatus,PapersPerMarker,IncludeView,ExcludeView,AllowDowngrade,WebPublish,ActivityGroup,Notes,CreatedBy,CreatedOn,AmendedBy,AmendedOn"
      'If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then mvResultColumns = mvResultColumns + ",LongDescription,ExamUnitLinkId,AccreditationStatus,AccreditationValidFrom,AccreditationValidTo,LocalName,CourseAccreditation,CourseAccreditationValidFrom,CourseAccreditationValidTo"

      Dim vFields As String = mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & ",eu.exam_unit_id,eu.exam_base_unit_id,eu.exam_session_id,exam_unit_code,exam_unit_description,subject,skill_level,eu.exam_unit_type,schedule_required,'' marker_required,exam_question,exam_unit_status,eu.sequence_number,session_based,eu.valid_from,eu.valid_to,product,rate,date_approved,registration_date,qcf_level,number_of_credits,nvq_code,svq_code,unit_time_limit,time_limit_type,minimum_students,maximum_students,student_count,minimum_age,allow_bookings,exam_mark_type,mark_factor,awarding_body,exam_unit_replaced_by_id,allow_exemptions,exemption_mark,exam_marker_status,papers_per_marker,include_view,exclude_view,allow_downgrade,eu.web_publish,eu.activity_group,eu.notes,eu.created_by,eu.created_on,eu.amended_by,eu.amended_on"
      Dim vExtraItems As String = "local_name,ecu.accreditation_status,ecu.accreditation_valid_from,ecu.accreditation_valid_to"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        vFields = vFields + ",eul.long_description,eul.exam_unit_link_id,eul.parent_unit_link_id,eul.accreditation_status,eul.accreditation_valid_from,eul.accreditation_valid_to,eu.is_grading_endpoint"
        If mvParameters.HasValue("ExamCentreUnitId") Then
          vFields = vFields + ",local_name,ecu.accreditation_status,ecu.accreditation_valid_from,ecu.accreditation_valid_to"
        Else
          vFields = vFields + ",,,,,"
        End If
      End If

      vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")
      vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")

      If mvParameters.HasValue("ExamCentreUnitId") Then
        vAnsiJoins.Add("exam_centre_units ecu", "ecu.exam_unit_link_id", "eul.exam_unit_link_id")
        vWhereFields.Add("ecu.exam_centre_unit_id", mvParameters("ExamCentreUnitId").IntegerValue)
      End If
      Dim vOrderByClause As String = String.Empty
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then vOrderByClause = mvEnv.Connection.DBIsNull("parent_unit_link_id", "0") & ", "
      vOrderByClause &= mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description"
      Dim vSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), "exam_units eu", vWhereFields, vOrderByClause, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, String.Empty, vExtraItems, True)

      For Each vRow As CARE.Access.CDBDataRow In pDataTable.Rows
        If (vRow.Item("ExamMarkerStatus").Length > 0 AndAlso (vRow.Item("ExamMarkerStatus") = "N" OrElse vRow.Item("ExamMarkerStatus") = "E")) Then
          vRow.Item("MarkerRequired") = "N"
        Else
          vRow.Item("MarkerRequired") = "Y"
        End If
      Next
    End Sub
    Private Function RemoveBlankItems(ByVal pItems As String) As String
      While pItems.Contains(",,")
        pItems = pItems.Replace(",,", ",")
      End While
      'BR 8771: Remove any last comma from end of line in case blank item(s) were at end:
      If pItems.EndsWith(",") Then pItems = pItems.Substring(0, pItems.Length - 1)
      Return pItems
    End Function
    Private Sub GetExamStudentBookingUnits(ByVal pDataTable As CDBDataTable)
      'When called from Validate Exam Booking units parameters are; ContactNumber, ExamSessionId, AllowBookings
      'When called from BookingCourses treeview on the client (Trader) parameters are; ContactNumber, ExamSessionId, AllowBookings

      'When called from BookingUnits treeview on the client parameters are; ContactNumber, ExamSessionId, ExamBookingId, StudentBookingOnly

      'the main uses of this unit are:  
      '2) with a ContactNumber and no value for StudentBookingOnly = the zero based template, with the contacts booked items (with marks if applicable)
      '3) with a ContactNumber and with StudentBookingOnly:Y = only the booked items for this student (with marks if applicable)
      'all of the above can and should be filtered by session too.

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitId") Then
        vWhereFields.Add("exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Else
        If mvParameters.HasValue("ExamSessionId") Then
          vWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
          'If mvParameters.HasValue("ExamBookingId") Then vWhereFields.Add("eb.exam_booking_id", mvParameters("ExamBookingId").IntegerValue)
        ElseIf mvParameters.HasValue("ExamSessionCode") Then
          vWhereFields.Add("es.exam_session_code", mvParameters("ExamSessionCode").Value)
          vAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "eu.exam_session_id")
        Else
          vWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
        End If
      End If
      If mvParameters.HasValue("SessionBased") Then vWhereFields.Add("session_based", mvParameters("SessionBased").Value)
      If mvParameters.HasValue("ValidTo") Then vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, mvParameters("ValidTo").Value, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      If mvParameters.HasValue("ExamUnitReplacedById") Then vWhereFields.Add("exam_unit_replaced_by_id", mvParameters("ExamUnitReplacedById").IntegerValue)
      If mvParameters.HasValue("AllowBookings") Then vWhereFields.Add("allow_bookings", mvParameters("AllowBookings").Value)

      Dim vFields As String = mvEnv.Connection.DBIsNull("exam_unit_id_1", "0") & ",eul.parent_unit_link_id,eu.exam_unit_id,eul.exam_unit_link_id,eu.exam_base_unit_id,eu.exam_session_id,0,0,0,0,'','','','','',"
      vFields &= "'N' AS booked,'N' AS passed,exam_unit_code,[EXAM_UNIT_DESCRIPTION_TO_BE_REPLACED],subject,skill_level,eu.exam_unit_type,schedule_required,exam_unit_status,eu.sequence_number,session_based,eu.valid_from,eu.valid_to,"
      vFields &= "product,rate,date_approved,registration_date,qcf_level,number_of_credits,nvq_code,svq_code,unit_time_limit,time_limit_type,minimum_students,maximum_students,student_count,minimum_age,allow_bookings,"
      vFields &= "exam_mark_type,mark_factor,awarding_body,exam_unit_replaced_by_id,allow_exemptions,exemption_mark,eu.activity_group,eu.notes,[EXAM_CENTRE_UNIT_ID],eu.created_by,eu.created_on,eu.amended_by,eu.amended_on"
      vFields &= ",eul.accreditation_status, eul.accreditation_valid_from, eul.accreditation_valid_to, eut.exam_question"

      vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")
      vAnsiJoins.Add("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2")
      If mvParameters.HasValue("ExamCentreId") Then 'Display the Local Name instead of the Unit Description
        If mvParameters.HasValue("ExamSessionId") Then
          vAnsiJoins.AddLeftOuterJoin("exam_centre_units ecbu", "ecbu.exam_unit_link_id", "eul.base_unit_link_id", "ecbu.exam_centre_id", mvParameters("ExamCentreId").Value) 'Session-based, so link to the Centre via the Base Link Id
          vFields = vFields.Replace("[EXAM_UNIT_DESCRIPTION_TO_BE_REPLACED]", "COALESCE(ecbu.local_name, exam_unit_description)") '
          vFields = vFields.Replace("[EXAM_CENTRE_UNIT_ID]", "ecbu.exam_centre_unit_id") '
        Else
          vAnsiJoins.Add("exam_centre_units ecu", "ecu.exam_unit_link_id", "eul.exam_unit_link_id", "ecu.exam_centre_id", mvParameters("ExamCentreId").Value) 'Non-Session-based, so link to the Centre directly via the Link Id
          vFields = vFields.Replace("[EXAM_UNIT_DESCRIPTION_TO_BE_REPLACED]", "COALESCE(ecu.local_name, exam_unit_description)")
          vFields = vFields.Replace("[EXAM_CENTRE_UNIT_ID]", "ecu.exam_centre_unit_id") '
        End If

      End If

      Dim vExtraSort As String = ""
      Dim vAddExtras As Boolean = False
      If mvParameters.HasValue("ContactNumber") Then vAddExtras = True
      If mvParameters.HasValue("StudentBookingOnly") AndAlso mvParameters("StudentBookingOnly").Value = "Y" Then vAddExtras = True

      If vAddExtras Then
        Dim vBookingWhereFields As New CDBFields
        Dim vBookingAnsiJoins As New AnsiJoins

        vBookingAnsiJoins.AddLeftOuterJoin("exam_booking_units ebu", "eb.exam_booking_id", "ebu.exam_booking_id")
        vBookingAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "ebu.exam_unit_link_id", "eul.exam_unit_link_id")
        vBookingAnsiJoins.AddLeftOuterJoin("exam_units eu", "ebu.exam_unit_id", "eu.exam_unit_id")
        vBookingAnsiJoins.AddLeftOuterJoin("exam_centre_units ecu", "ecu.exam_unit_link_id", "ebu.exam_unit_link_id", "ecu.exam_centre_id", "eb.exam_centre_id") 'For non-session based
        vBookingAnsiJoins.AddLeftOuterJoin("exam_centre_units ecbu", "ecbu.exam_unit_link_id", "eul.base_unit_link_id", "ecbu.exam_centre_id", "eb.exam_centre_id") 'for session based

        If mvParameters.HasValue("ExamSessionId") Then
          vBookingWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
        ElseIf mvParameters.HasValue("ExamSessionCode") Then
          vBookingWhereFields.Add("es.exam_session_code", mvParameters("ExamSessionCode").Value)
          vBookingAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "eu.exam_session_id")
        Else
          vBookingWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
        End If

        ' if student marks only, if no mark data exists, dont return template data
        If mvParameters.HasValue("StudentBookingOnly") AndAlso mvParameters("StudentBookingOnly").Value = "Y" Then vWhereFields.Add("eb.contact_number", "", CDBField.FieldWhereOperators.fwoNotEqual)

        ' if contact is supplied, limit to this contact
        If mvParameters.HasValue("ContactNumber") Then
          vBookingWhereFields.Add("eb.contact_number", mvParameters("ContactNumber").IntegerValue)
          If Not mvParameters.HasValue("StudentBookingOnly") Then
            vBookingWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
            vBookingWhereFields.Add("ebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings
            'Ignore bookings that have been graded and student failed
            vBookingWhereFields.Add("ebu.total_result", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
            vBookingWhereFields.Add("ebu.total_result#2", CDBField.FieldTypes.cftCharacter, "F", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
          End If
        Else
          vExtraSort = "contact_number,"
        End If

        Dim vInnerFields As String = "eb.contact_number,ebu.exam_unit_id,eb.exam_booking_id,ebu.exam_booking_unit_id, ebu.exam_unit_link_id, ebu.total_mark, ebu.total_grade, ebu.total_result, ebu.exam_student_unit_status, ebu.done_date, eb.exam_session_id, eb.special_requirements, ecu.local_name, ecbu.local_name base_local_name"
        Dim vGroupByFields As String = "eb.contact_number,ebu.exam_unit_id,eb.exam_booking_id,ebu.exam_booking_unit_id, ebu.exam_unit_link_id, ebu.total_mark, ebu.total_grade, ebu.total_result, ebu.exam_student_unit_status, ebu.done_date, eb.exam_session_id, eb.special_requirements, ecu.local_name, ecbu.local_name"
        Dim vBookingSQL As New SQLStatement(mvEnv.Connection, vInnerFields, "exam_bookings eb", vBookingWhereFields, "", vBookingAnsiJoins)
        vBookingSQL.GroupBy = vGroupByFields ' added as multi summary header lines caused cartesian join.

        If mvParameters.ContainsKey("ContactExamDetails") AndAlso String.Compare(mvParameters("ContactExamDetails").Value, "Y", True) = 0 Then
          vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) eb", vBookingSQL.SQL), "eb.exam_unit_id", "eu.exam_unit_id", "eb.exam_unit_link_id", "eul.exam_unit_link_id", mvEnv.Connection.DBIsNull("eb.exam_session_id", "0"), mvEnv.Connection.DBIsNull("eu.exam_session_id", "0"))
        Else
          vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) eb", vBookingSQL.SQL), "eb.exam_unit_id", "eu.exam_unit_id", mvEnv.Connection.DBIsNull("eb.exam_session_id", "0"), mvEnv.Connection.DBIsNull("eu.exam_session_id", "0"))
        End If


        vFields = vFields.Replace("0,0,0,0,'','','','',''", "exam_booking_unit_id,eb.exam_booking_id,eb.total_mark,total_grade,total_result,eb.exam_student_unit_status," + mvEnv.Connection.DBIsNull("eb.contact_number", "0") + " contact_number,eb.done_date,eb.special_requirements")
        vFields = vFields.Replace("[EXAM_UNIT_DESCRIPTION_TO_BE_REPLACED]", "COALESCE(eb.local_name, eb.base_local_name, exam_unit_description)")

        If mvParameters.HasValue("ContactNumber") Then
          Dim vPassesWhereFields As New CDBFields
          vPassesWhereFields.Add("eg.grade_is_pass", "Y")
          vPassesWhereFields.Add("expires", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoOpenBracket)
          vPassesWhereFields.Add("expires#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
          vPassesWhereFields.Add("contact_number", mvParameters("ContactNumber").IntegerValue)
          Dim vPassesAnsiJoins As New AnsiJoins
          vPassesAnsiJoins.Add("exam_student_unit_header esuh", "esh.exam_student_header_id", "esuh.exam_student_header_id")
          vPassesAnsiJoins.Add("exam_grades eg", "esuh.current_grade", "eg.exam_grade")
          Dim vPassesSQL As New SQLStatement(mvEnv.Connection, "esuh.exam_unit_id, grade_is_pass", "exam_student_header esh ", vPassesWhereFields, "", vPassesAnsiJoins)
          If mvParameters.HasValue("ExamSessionId") Then
            vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) esg", vPassesSQL.SQL), "esg.exam_unit_id", "eu.exam_base_unit_id")
          Else
            vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) esg", vPassesSQL.SQL), "esg.exam_unit_id", "eu.exam_unit_id") 'unit id is base unit id for non-session based bookings
          End If
          vFields = vFields.Replace("'N' AS passed", "grade_is_pass AS passed")
        End If
      End If

      'In case the code's got this far and the exam unit description hasn't been replaced yet, replace with a default value
      vFields = vFields.Replace("[EXAM_UNIT_DESCRIPTION_TO_BE_REPLACED]", "exam_unit_description")
      vFields = vFields.Replace("[EXAM_CENTRE_UNIT_ID]", "''")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_units eu", vWhereFields, vExtraSort & mvEnv.Connection.DBIsNull("eul.parent_unit_link_id", "0") & "," & mvEnv.Connection.DBIsNull("eu.sequence_number", "10000") & ",exam_unit_description", vAnsiJoins)
      vSQL.UseAnsiSQL = True
      pDataTable.FillFromSQL(mvEnv, vSQL)

      For Each vRow As CARE.Access.CDBDataRow In pDataTable.Rows
        If (vRow.Item("ExamBookingUnitId") <> "0" AndAlso vRow.Item("ExamBookingUnitId") <> "") Then vRow.Item("Booked") = "Y"
        If (vRow.Item("ExamBaseUnitId") = "0" AndAlso vRow.Item("ExamSessionId") = "") Then vRow.Item("ExamBaseUnitId") = vRow.Item("ExamUnitId") ' for session zero ensure base unit id is set (this is the base unit)
      Next

      'We may have to change this because this fix is temporary and will needs to be changed before GA
      If mvParameters.ContainsKey("AllowBookings") = True AndAlso pDataTable.Rows.Count > 0 Then
        'Do not check Accreditation or Study modes when validating the exam unit while booking, as these units are already validated for both
        If Not mvParameters.ContainsKey("ValidateParameters") Then
          Dim vRemoveIds As New List(Of Integer)

          'build a dictionary of all exam unit links and their parents.  We will gradually filter-out all un-accredited units and then remove all from the final list
          Dim vExamUnitLinks As New Dictionary(Of Integer, Integer)
          pDataTable.Rows.ForEach(Sub(vRow) vExamUnitLinks.Add(vRow.IntegerItem("ExamUnitLinkId"), vRow.IntegerItem("ParentUnitLinkId")))

          Dim vSessionId As Integer = 0
          If mvParameters.ContainsKey("ExamSessionId") Then
            vSessionId = mvParameters("ExamSessionId").IntegerValue
          End If

          Dim vAccreditationFilter As New ExamAccreditationFilter(mvEnv, vExamUnitLinks, ExamAccreditationFilterTypes.eafAllowRegistration) 'This class handles all the accreditation filtering

          'Filter-out un-accredited Exam Units
          Dim vUnitAccreditationIsEnabled As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamUnitAccreditation)
          If vUnitAccreditationIsEnabled.Length > 0 AndAlso vUnitAccreditationIsEnabled = "Y" Then
            Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnaccreditedUnits()
            vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
          End If

          'Filter-out unaccredited Centres and unaccredited centre-units
          If mvParameters.HasValue("ExamCentreId") Then
            'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
            Dim vCentreAccreditationIsEnabled As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation)
            If vCentreAccreditationIsEnabled.Length > 0 AndAlso vCentreAccreditationIsEnabled = "Y" Then
              Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnitsAtUnaccreditedCentre(mvParameters("ExamCentreId").IntegerValue, vSessionId)
              vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
            End If

            'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
            Dim vCentreUnitAccreditationIsEnabled As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation)
            If vCentreUnitAccreditationIsEnabled.Length > 0 AndAlso vCentreUnitAccreditationIsEnabled = "Y" Then
              Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentreUnits(mvParameters("ExamCentreId").IntegerValue, vSessionId)
              vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
            End If
          End If

          If vRemoveIds IsNot Nothing AndAlso vRemoveIds.Count > 0 Then
            Dim vRemoveRows As List(Of CDBDataRow) = pDataTable.Rows.FindAll(Function(vRow) vRemoveIds.Contains(vRow.IntegerItem("ExamUnitLinkId")))
            vRemoveRows.ForEach(Sub(vRow) pDataTable.RemoveRow(vRow))
          End If

          'Now check study modes (only if any study modes have been created)
          If mvEnv.Connection.GetCount("study_modes", Nothing) <> 0 Then
            Dim vStudyModeFilter As New ExamStudyModeFilter(mvEnv, vExamUnitLinks)
            Dim vStudyMode As String = ""
            If mvParameters.ContainsKey("StudyMode") Then
              vStudyMode = mvParameters("StudyMode").Value
            End If

            Dim vIncludeIds As New List(Of Integer)
            If mvParameters.HasValue("ExamCentreId") Then
              vIncludeIds = vStudyModeFilter.GetCentreUnitsForStudyMode(vStudyMode, mvParameters("ExamCentreId").IntegerValue, vSessionId)
            Else
              vIncludeIds = vStudyModeFilter.GetUnitsForStudyMode(vStudyMode, vSessionId)
            End If
            If vIncludeIds IsNot Nothing Then
              Dim vRemoveRows As List(Of CDBDataRow) = pDataTable.Rows.FindAll(Function(vRow) Not vIncludeIds.Contains(vRow.IntegerItem("ExamUnitLinkId")))
              vRemoveRows.ForEach(Sub(vRow) pDataTable.RemoveRow(vRow))
            End If
          End If
        End If
      End If

    End Sub

    Public Function FilterByUnitStudyMode(ByVal pDT As CDBDataTable) As CDBDataTable
      Dim vParentNodes As New List(Of Integer)
      Dim vDataTable As New CDBDataTable
      Dim vDTable As CDBDataTable = pDT
      Dim vStudyModes As IList(Of String) = Nothing
      Dim vToRemove As New List(Of Integer)
      For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1

        Dim vUnitLinkId As Integer = pDT.Rows.Item(vRowNumber).IntegerItem("ExamUnitLinkId")

        If Not vParentNodes.Contains(vUnitLinkId) Then

          If Not HasChildren(vUnitLinkId, pDT.ConvertToDataTable) Then
            If mvParameters.ContainsKey("ExamCentreId") Then
              vStudyModes = GetCentreUnitLinkStudyModes(mvParameters("ExamCentreId").IntegerValue, vUnitLinkId)
            Else
              vStudyModes = GetUnitLinkStudyModes(vUnitLinkId)
            End If

            If mvParameters.ContainsKey("StudyMode") Then
              If vStudyModes.Contains(mvParameters("StudyMode").Value) Then
                If Not vParentNodes.Contains(vUnitLinkId) Then
                  vParentNodes.Add(vUnitLinkId)
                End If
                For Each vParent As Integer In GetAncestorUnitLinks(vUnitLinkId)
                  If Not vParentNodes.Contains(vParent) Then
                    vParentNodes.Add(vParent)
                  End If
                Next vParent
              End If
            Else
              If vStudyModes.Count = 0 Then
                If Not vParentNodes.Contains(vUnitLinkId) Then
                  vParentNodes.Add(vUnitLinkId)
                End If
                For Each vParent As Integer In GetAncestorUnitLinks(vUnitLinkId)
                  If Not vParentNodes.Contains(vParent) Then
                    vParentNodes.Add(vParent)
                  End If
                Next vParent
              End If
            End If
          End If
        Else
          If Not vParentNodes.Contains(CInt(pDT.Rows.Item(vRowNumber).Item("ParentUnitLinkId"))) Then
            vParentNodes.Add(CInt(pDT.Rows.Item(vRowNumber).Item("ParentUnitLinkId")))
          End If
        End If

      Next vRowNumber

      For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1
        If Not vParentNodes.Contains(CInt(pDT.Rows.Item(vRowNumber).Item("ExamUnitLinkId"))) Then
          pDT.Rows.RemoveAt(vRowNumber)
        End If
      Next vRowNumber
      Return pDT
    End Function

    Private Function GetCentreUnitLinkStudyModes(ByVal pCentreId As Integer, ByVal pUnitLinkId As Integer) As IList(Of String)
      Dim vUnits As New List(Of Integer)({GetBaseUnitLinkId(pUnitLinkId)})
      vUnits.AddRange(GetDescendantUnitLinks(vUnits(0)))
      vUnits.AddRange(GetAncestorUnitLinks(vUnits(0)))
      Dim vResult As IList(Of String) = New List(Of String)(From vRow As DataRow In (New SQLStatement(mvEnv.Connection,
                                                                                                      "DISTINCT study_mode",
                                                                                                      "exam_centre_units xcu",
                                                                                                      New CDBFields({New CDBField("exam_unit_link_id", vUnits),
                                                                                                                     New CDBField("exam_centre_id", pCentreId)}),
                                                                                                      "",
                                                                                                      New AnsiJoins({New AnsiJoin("exam_centre_unit_study_modes xcusm",
                                                                                                                                  "xcusm.exam_centre_unit_link_id",
                                                                                                                                  "xcu.exam_centre_unit_id")}))).GetDataTable.AsEnumerable
                                                            Select CStr(vRow("study_mode"))).AsReadOnly
      Return If(vResult.Count > 0, vResult, GetUnitLinkStudyModes(pUnitLinkId))
    End Function

    Private Function HasChildren(ByVal pUnitLinkId As Integer, pData As DataTable) As Boolean
      Dim vRelevantChildren As DataView = pData.AsDataView
      vRelevantChildren.RowFilter = "ExamUnitLinkId in (" & pUnitLinkId & "," & GetDescendantUnitLinks(pUnitLinkId).AsCommaSeperated & ")"
      Return vRelevantChildren.ToTable.Rows.Count > 1
    End Function

    Private Function GetUnitLinkStudyModes(ByVal pUnitLinkId As Integer) As IList(Of String)
      Dim vUnits As New List(Of Integer)({GetBaseUnitLinkId(pUnitLinkId)})
      vUnits.AddRange(GetDescendantUnitLinks(vUnits(0)))
      vUnits.AddRange(GetAncestorUnitLinks(vUnits(0)))
      Return New List(Of String)(From vRow As DataRow In (New SQLStatement(mvEnv.Connection,
                                                          "DISTINCT study_mode",
                                                          "exam_unit_study_modes",
                                                          New CDBFields({New CDBField("exam_unit_link_id", vUnits)}))).GetDataTable.AsEnumerable
                                 Select CStr(vRow("study_mode"))).AsReadOnly
    End Function

    Private Function GetBaseUnitLinkId(ByVal pUnitLinkId As Integer) As Integer
      Dim vTemp As DataTable = New SQLStatement(mvEnv.Connection,
                                                             "base_unit_link_id",
                                                             "exam_unit_links",
                                                             New CDBFields({New CDBField("exam_unit_link_id", pUnitLinkId)})).GetDataTable
      Return If(vTemp.Rows.Count > 0 AndAlso Not vTemp.Rows(0).IsNull("base_unit_link_id") AndAlso CInt(vTemp.Rows(0)("base_unit_link_id")) <> 0,
                CInt(vTemp.Rows(0)("base_unit_link_id")),
                pUnitLinkId)
    End Function

    Private Function GetDescendantUnitLinks(ByVal pUnitLinkId As Integer) As IEnumerable(Of Integer)
      Dim vSql As New StringBuilder
      vSql.AppendLine("WITH cte(exam_unit_link_id, ")
      vSql.AppendLine("         parent_unit_link_id) ")
      vSql.AppendLine("     AS (SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("         WHERE  eul.exam_unit_link_id = " & pUnitLinkId.ToString & " ")
      vSql.AppendLine("         UNION ALL ")
      vSql.AppendLine("         SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("                inner join cte ")
      vSql.AppendLine("                        ON cte.exam_unit_link_id = eul.parent_unit_link_id) ")
      vSql.AppendLine("SELECT cte.exam_unit_link_id ")
      vSql.AppendLine("FROM   cte ")
      vSql.AppendLine("WHERE  cte.exam_unit_link_id <> " & pUnitLinkId.ToString)
      Return From vRow As DataRow In New SQLStatement(mvEnv.Connection, vSql.ToString).GetDataTable().AsEnumerable
                                                      Select CInt(vRow("exam_unit_link_id"))
    End Function

    Private Function GetAncestorUnitLinks(ByVal pUnitLinkId As Integer) As IEnumerable(Of Integer)
      Dim vSql As New System.Text.StringBuilder
      vSql.AppendLine("WITH cte(exam_unit_link_id, parent_unit_link_id) ")
      vSql.AppendLine("     AS (SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("         WHERE  eul.exam_unit_link_id = " & pUnitLinkId.ToString & " ")
      vSql.AppendLine("         UNION ALL ")
      vSql.AppendLine("         SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("                INNER JOIN cte ")
      vSql.AppendLine("                        ON cte.parent_unit_link_id = eul.exam_unit_link_id) ")
      vSql.AppendLine("SELECT cte.exam_unit_link_id ")
      vSql.AppendLine("FROM   cte ")
      vSql.AppendLine("WHERE  cte.exam_unit_link_id <> " & pUnitLinkId.ToString)
      Return From vRow As DataRow In New SQLStatement(mvEnv.Connection, vSql.ToString).GetDataTable().AsEnumerable
                                                      Select CInt(vRow("exam_unit_link_id"))
    End Function

    Private Sub GetExamStudentResultEntry(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vExamUnits As String = ""
      Dim vMultipleExamUnits As Boolean = False
      Dim vBySessions As Boolean

      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")

      If mvEnv.GetConfig("ex_allow_multiple_results", "N") = "Y" Then
        vMultipleExamUnits = True
      End If

      If mvParameters.HasValue("ExamUnitId") AndAlso vMultipleExamUnits Then
        vExamUnits = GetLinkedExamUnits(mvParameters("ExamUnitId").Value)
      Else
        vExamUnits = mvParameters("ExamUnitId").Value
      End If

      vBySessions = mvParameters.HasValue("ExamSessionId") AndAlso mvParameters("ExamSessionId").IntegerValue > 0

      Dim vFields As String = "exam_booking_unit_id,c.contact_number,raw_mark,raw_mark AS raw_mark_check,original_mark,original_mark AS original_mark_check,original_grade,original_grade AS original_grade_check,original_result,original_result AS original_result_check,eu.exam_mark_type,eul.exam_unit_id_1,eu.exam_unit_id,eu.exam_unit_code"
      If vBySessions Then
        vWhereFields.Add("eb.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      Else
        vWhereFields.Add("eb.exam_session_id", 0)
      End If

      If vBySessions Then
        If mvParameters.HasValue("ExamCentreId") Then
          'If exam unit type for the exam unit requires scheduling then the exam schedule should be for exam centre id. 
          vWhereFields.Add("esc.exam_centre_id", mvParameters("ExamCentreId").IntegerValue, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
          'If scheduling is not required for the exam unit type then no exam schedule record should exists and exam bookings record should be for the exam centre id.
          vWhereFields.Add("esc.exam_schedule_id", 0, CDBField.FieldWhereOperators.fwoNullOrEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("eb.exam_centre_id", mvParameters("ExamCentreId").IntegerValue, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
        End If
      Else
        If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      End If

      If vMultipleExamUnits Then
        If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("ebu.exam_unit_id", vExamUnits, CARE.Data.CDBField.FieldWhereOperators.fwoIn)
      Else
        vWhereFields.Add("ebu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      End If

      vAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")
      vAnsiJoins.Add("exam_unit_links eu_links", "eu_links.exam_unit_link_id", "ebu.exam_unit_link_id")
      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "ebu.exam_unit_id")

      If vBySessions Then
        vAnsiJoins.Add("exam_unit_types eut", "eut.exam_unit_type", "eu.exam_unit_type")
        vAnsiJoins.AddLeftOuterJoin("exam_schedule esc", "ebu.exam_schedule_id", "esc.exam_schedule_id")
      End If

      vAnsiJoins.Add("contacts c", "eb.contact_number", "c.contact_number")
      Dim vSubAnsiJoins As New AnsiJoins
      Dim vSubWhereFields As New CDBFields
      vSubAnsiJoins.Add("exam_units eu_sub", "eu_sub.exam_unit_id", "eul_sub.exam_unit_id_2")
      vSubAnsiJoins.Add("exam_unit_types eut", "eut.exam_unit_type", "eu_sub.exam_unit_type")
      vSubWhereFields.Add("eut.exam_question", "Y")
      vSubWhereFields.Add("eu_sub.exam_marker_status", CDBField.FieldTypes.cftCharacter, "N", CARE.Data.CDBField.FieldWhereOperators.fwoNullOrNotEqual) ' Exclude marker status of 'N' - No Markers Required (No Result Entry)
      Dim vSubSQL As New SQLStatement(mvEnv.Connection, "distinct exam_unit_id_1,parent_unit_link_id", "exam_unit_links eul_sub", vSubWhereFields, "", vSubAnsiJoins)
      vAnsiJoins.AddLeftOuterJoin(String.Format("( {0} ) eul", vSubSQL.SQL), "ebu.exam_unit_id", "eul.exam_unit_id_1", "ebu.exam_unit_link_id", "eul.parent_unit_link_id")
      vWhereFields.Add("eb.cancelled_on") 'Ignore cancellations
      vWhereFields.Add("ebu.total_result")
      If vBySessions Then
        'If exam unit type for the exam unit does not require scheduling then return row.   
        vWhereFields.Add("eut.schedule_required", "N", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        'If exam unit type for the exam unit requires scheduling then only return if row has an exam schedule record.   
        vWhereFields.Add("esc.exam_session_id", 0, CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_booking_units ebu", vWhereFields, "ebu.exam_unit_id", vAnsiJoins) '"c.contact_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)

      If pDataTable.Rows.Count > 0 Then
        pDataTable = ExamUnit.IsUnitAccredited(mvEnv, pDataTable, True, True)
      End If
    End Sub
    ''' <summary>
    ''' Gets all the list of all the Exam Units that are linked (directly - indirectly) with the main exam unit
    ''' </summary>
    ''' <param name="pExamUnit">Main Exam Unit</param>
    ''' <returns>Comma seprated list of all the linked exam unit(s)</returns>
    ''' <remarks></remarks>
    Private Function GetLinkedExamUnits(ByVal pExamUnit As String) As String
      Dim vResult As New StringBuilder(pExamUnit)
      Dim vSqlString As String = "Select exam_unit_id_2 from exam_unit_links where exam_unit_id_1=" & "'" & pExamUnit & "'"
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQLDONOTUSE(mvEnv, vSqlString)

      For Each vDataRow As CDBDataRow In vDataTable.Rows
        vResult.Append("," + (vDataRow.Item("exam_unit_id_2")))
      Next
      Return vResult.ToString
    End Function

    Private Sub GetExamStudentResultEntryComponents(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins

      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")
      ' Mark fields deliberately doubled up to allow "check" fields in service response
      Dim vFields As String = "ebu_comps.exam_booking_unit_id,c.contact_number,eu.exam_unit_code,eu.exam_unit_description,ebu_comps.raw_mark,ebu_comps.raw_mark,ebu_comps.original_mark,ebu_comps.original_mark,eu.exam_mark_type,eu.mark_factor"
      If mvParameters.HasValue("ExamSessionId") Then vWhereFields.Add("eb.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("esc.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("eul.exam_unit_id_1", mvParameters("ExamUnitId").IntegerValue)
      If mvParameters.HasValue("ContactNumber") Then vWhereFields.Add("c.contact_number", mvParameters("ContactNumber").IntegerValue)
      vAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")
      vAnsiJoins.Add("exam_unit_links eul", "eul.parent_unit_link_id", "ebu.exam_unit_link_id")
      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "eul.exam_unit_id_2")
      vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")
      vAnsiJoins.Add("exam_schedule esc", "ebu.exam_schedule_id", "esc.exam_schedule_id")
      vAnsiJoins.Add("contacts c", "eb.contact_number", "c.contact_number")
      vAnsiJoins.Add("exam_bookings eb_comps", "eb_comps.contact_number", "c.contact_number")
      vAnsiJoins.Add("exam_booking_units ebu_comps", "ebu_comps.exam_booking_id", "eb_comps.exam_booking_id", "ebu_comps.exam_unit_id", "eu.exam_unit_id")

      vWhereFields.Add("eut.exam_question", "Y") 'Only return Question type components
      vWhereFields.Add("eu.exam_marker_status", CDBField.FieldTypes.cftCharacter, "N", CARE.Data.CDBField.FieldWhereOperators.fwoNullOrNotEqual) ' Exclude marker status of 'N' - No Markers Required (No Result Entry)
      vWhereFields.Add("eb.cancelled_on") 'Ignore cancellations
      vWhereFields.Add("eb_comps.cancelled_on") 'Ignore cancellations      
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_booking_units ebu", vWhereFields, "c.contact_number, eu.exam_unit_code", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)
    End Sub

    Private Sub GetExamCandidateActivities(ByVal pDataTable As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation) Then
        Dim vWhereFields As New CDBFields
        Dim vAnsiJoins As New AnsiJoins
        Dim vFields As String = "eca.exam_candidate_activity_id,eca.exam_booking_unit_id,eca.activity,eca.activity_value,eca.quantity,eca.activity_date,eca.source,eca.valid_from,eca.valid_to,eca.amended_by,eca.amended_on,eca.notes,a.activity_desc,av.activity_value_desc,s.source_desc,'' NoteFlag, ''Status, '' StatusOrder"
        vAnsiJoins.Add("activities a", "a.activity", "eca.activity")
        vAnsiJoins.Add("activity_values av", "av.activity_value", "eca.activity_value", "av.activity", "eca.activity")
        vAnsiJoins.Add("sources s", "s.source", "eca.source")
        If mvParameters.HasValue("ExamBookingUnitId") Then vWhereFields.Add("eca.exam_booking_unit_id", mvParameters("ExamBookingUnitId").IntegerValue)
        If mvParameters.HasValue("Activity") Then vWhereFields.Add("eca.activity", mvParameters("Activity").Value)
        If mvParameters.HasValue("ActivityValue") Then vWhereFields.Add("eca.activity_value", mvParameters("ActivityValue").Value)
        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "exam_candidate_activities eca", vWhereFields, "", vAnsiJoins)
        pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)

        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item("Notes").Length > 0 Then vRow.Item("NoteFlag") = "Y"
          vRow.SetCurrentFutureHistoric("Status", "StatusOrder")
        Next

        pDataTable.ReOrderRowsByColumn("StatusOrder")
      End If
    End Sub

    Private Sub GetExamUnitMarkerAllocation(ByVal pDataTable As CDBDataTable)
      Dim vSubWhereFields As New CDBFields
      Dim vSubAnsiJoins As New AnsiJoins
      If mvParameters.HasValue("ExamUnitId") Then vSubWhereFields.Add("embd.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      Dim vSubFields As String = "embd.exam_unit_id,embd.exam_personnel_id,c.contact_number,embd.marker_number,count(*) AS number_of_papers"
      vSubAnsiJoins.Add("exam_booking_units ebu", "embd.exam_booking_unit_id", "ebu.exam_booking_unit_id")
      vSubAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")
      vSubAnsiJoins.Add("exam_personnel ep", "ep.exam_personnel_id", "embd.exam_personnel_id")
      vSubAnsiJoins.Add("contacts c", "ep.contact_number", "c.contact_number")
      vSubWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
      vSubWhereFields.Add("ebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings
      Dim vSubSelect As New SQLStatement(mvEnv.Connection, vSubFields, "exam_marking_batch_detail embd", vSubWhereFields, "", vSubAnsiJoins)
      vSubSelect.GroupBy = "embd.exam_unit_id,embd.exam_personnel_id,c.contact_number,embd.marker_number"

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName.Replace("c.contact_number,", "")
      Dim vFields As String = "papers.exam_unit_id,papers.exam_personnel_id,con.contact_number,papers.marker_number,papers.number_of_papers,'N' unallocated"
      vAnsiJoins.Add(String.Format("({0} ) papers", vSubSelect.SQL), "papers.contact_number", "con.contact_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "contacts con", vWhereFields, "con.surname,con.forenames,marker_number", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, ContactNameItems)

      If mvParameters.HasValue("ExamUnitId") AndAlso mvParameters.HasValue("GetUnallocatedCount") AndAlso mvParameters("GetUnallocatedCount").Value = "Y" Then
        Dim vDT As CDBDataTable = New CDBDataTable()
        vDT.AddColumnsFromList("MarkerNumber,MarkerNumberDesc")
        Dim vMaxNumber As Integer = IntegerValue(mvEnv.GetConfig("ex_markings_per_paper", "1"))
        If vMaxNumber < 1 Then vMaxNumber = 1
        For i As Integer = 1 To vMaxNumber
          Dim vCountWhereFields As New CDBFields
          Dim vCountAnsiJoins As New AnsiJoins
          Dim vCountFields As String = "ebu.exam_booking_unit_id"
          vCountAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")

          Dim vCountSubWhereFields As New CDBFields
          vCountSubWhereFields.Add("subembd.exam_booking_unit_id", CDBField.FieldTypes.cftInteger, "ebu.exam_booking_unit_id")
          vCountSubWhereFields.Add("subembd.marker_number", i)
          Dim vCountSubSelect As New SQLStatement(mvEnv.Connection, "subembd.exam_booking_unit_id", "exam_marking_batch_detail subembd", vCountSubWhereFields)
          vCountWhereFields.Add("Exclude", vCountSubSelect.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
          vCountWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
          vCountWhereFields.Add("ebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings
          vCountWhereFields.Add("ebu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
          Dim vCountSelect As New SQLStatement(mvEnv.Connection, vCountFields, "exam_booking_units ebu", vCountWhereFields, "", vCountAnsiJoins)
          Dim vCount As Integer = mvEnv.Connection.GetCountFromStatement(vCountSelect)

          If vCount > 0 Then
            Dim vRow As CDBDataRow = pDataTable.AddRow()
            vRow.Item("ExamUnitId") = mvParameters("ExamUnitId").Value
            vRow.Item("ExamPersonnelId") = ""
            vRow.Item("ContactNumber") = ""
            vRow.Item("ContactName") = ""
            vRow.Item("MarkerNumber") = i.ToString()
            vRow.Item("Unallocated") = "Y"
            vRow.Item("NumberOfPapers") = vCount.ToString()
          End If
        Next
      End If
    End Sub

    Private Sub GetExamUnitMarkerAllocationList(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vFields As String = "'' as select_rec,embd.exam_marking_batch_detail_id,ebu.exam_booking_unit_id,ebu.exam_unit_id,embd.exam_personnel_id,ec.exam_centre_id,ec.exam_centre_description,ec.exam_centre_code,ebu.exam_candidate_number"

      If mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("embd.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      vWhereFields.Add("eb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
      vWhereFields.Add("ebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings

      vAnsiJoins.Add("exam_bookings eb", "eb.exam_booking_id", "ebu.exam_booking_id")
      vAnsiJoins.Add("contacts c", "eb.contact_number", "c.contact_number")
      vAnsiJoins.AddLeftOuterJoin("exam_centres ec", "ec.exam_centre_id", "eb.exam_centre_id")

      If Not mvParameters.HasValue("ExamPersonnelId") Then ' Get Unallocated papers
        If Not mvParameters.HasValue("MarkerNumber") Then RaiseError(DataAccessErrors.daeParameterNotFound, "MarkerNumber")
        If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("ebu.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
        vFields = vFields.Replace("embd.", "'' ")
        Dim vSubWhereFields As New CDBFields
        vSubWhereFields.Add("subembd.exam_booking_unit_id", CDBField.FieldTypes.cftInteger, "ebu.exam_booking_unit_id")
        vSubWhereFields.Add("subembd.marker_number", mvParameters("MarkerNumber").IntegerValue)
        Dim vSubSQL As New SQLStatement(mvEnv.Connection, "subembd.exam_booking_unit_id", "exam_marking_batch_detail subembd", vSubWhereFields)
        vWhereFields.Add("Exclude", vSubSQL.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
      Else
        If mvParameters.HasValue("ExamUnitId") Then vWhereFields.Add("embd.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
        vAnsiJoins.Add("exam_marking_batch_detail embd", "embd.exam_booking_unit_id", "ebu.exam_booking_unit_id")
        vWhereFields.Add("embd.exam_personnel_id", mvParameters("ExamPersonnelId").IntegerValue)
        If mvParameters.HasValue("MarkerNumber") Then vWhereFields.Add("embd.marker_number", mvParameters("MarkerNumber").IntegerValue)
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_booking_units ebu", vWhereFields, "surname,forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, vFields, "contact_number," & ContactNameItems())
    End Sub

    Private Sub GetExamMarkerList(ByVal pDataTable As CDBDataTable)
      ' Method returns list of markers available to mark exam units
      ' Parameters required:  ExamBookingUnitId (list of ExamBookingUnitId) and ExamUnitId

      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      Dim vContact As New Contact(mvEnv)
      Dim vConAttrs As String = vContact.GetRecordSetFieldsName
      Dim vFields As String = "eup.exam_unit_id,eup.exam_personnel_id,ept.exam_personnel_type,ept.exam_personnel_type_desc,eu.exam_marker_status"
      Dim vSubSqlStr As String = ""

      vWhereFields.Add("ept.exam_marker", "Y")
      vWhereFields.Add("eup.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      vWhereFields.Add("eup.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("eup.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
      vWhereFields.Add("eu.exam_marker_status", CDBField.FieldTypes.cftCharacter, "'A','R'", CDBField.FieldWhereOperators.fwoIn) ' N=None (No Result Entry), E=None(Allow Result Entry), R=Required, A=Allocated, L=Locked

      If mvParameters.HasValue("ExamBookingUnitId") AndAlso mvParameters("ExamBookingUnitId").Value.Length > 0 Then ' build sub sql to exclude markers already marking passed in papers
        Dim vExamBookingUnitIdArray As String() = mvParameters("ExamBookingUnitId").Value.Split(","c)
        Dim vExamBookingUnitIdCsv As String = ""

        For i As Integer = vExamBookingUnitIdArray.GetLowerBound(0) To vExamBookingUnitIdArray.GetUpperBound(0)
          Dim vExamBookingUnitId As Integer
          If Integer.TryParse(vExamBookingUnitIdArray(i), vExamBookingUnitId) Then
            If vExamBookingUnitIdCsv.Length > 0 Then vExamBookingUnitIdCsv = vExamBookingUnitIdCsv + ","
            vExamBookingUnitIdCsv += vExamBookingUnitId.ToString()
          End If
        Next

        If vExamBookingUnitIdCsv.Length > 0 Then
          Dim vInnerFields As String = "embd.exam_personnel_id"
          Dim vInnerAnsiJoins As New AnsiJoins
          Dim vInnerWhereFields As New CDBFields
          vInnerAnsiJoins.Add("exam_booking_units innerebu", "embd.exam_booking_unit_id", "innerebu.exam_booking_unit_id")
          vInnerAnsiJoins.Add("exam_bookings innereb", "innereb.exam_booking_id", "innerebu.exam_booking_id")
          vInnerWhereFields.Add("innerebu.exam_booking_unit_id", vExamBookingUnitIdCsv, CDBField.FieldWhereOperators.fwoIn)
          vInnerWhereFields.Add("innereb.cancellation_reason", CDBField.FieldTypes.cftCharacter)  'Ignore cancelled bookings
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitCancellation) Then vInnerWhereFields.Add("innerebu.cancellation_reason", CDBField.FieldTypes.cftCharacter) 'Ignore cancelled bookings
          Dim vInnerSQL As New SQLStatement(mvEnv.Connection, vInnerFields, "exam_marking_batch_detail embd", vInnerWhereFields, "", vInnerAnsiJoins)
          vWhereFields.Add("eup.exam_personnel_id", vInnerSQL.SQL, CDBField.FieldWhereOperators.fwoNotIn)
        End If
      End If

      vAnsiJoins.Add("exam_units eu", "eu.exam_unit_id", "eup.exam_unit_id")
      vAnsiJoins.Add("exam_personnel_types ept", "eup.exam_personnel_type", "ept.exam_personnel_type")
      vAnsiJoins.Add("exam_personnel ep", "ep.exam_personnel_id", "eup.exam_personnel_id")
      vAnsiJoins.Add("contacts c", "ep.contact_number", "c.contact_number")
      Dim vSQL As New SQLStatement(mvEnv.Connection, vFields & "," & vConAttrs, "exam_unit_personnel eup", vWhereFields, "surname,forenames", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL, "c.contact_number," & vFields, ContactNameItems)
    End Sub

    ''' <summary>
    ''' Check If the Centers are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Envronment class</param>
    ''' <param name="pDT">DatTable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsCentreAccredited(ByVal pEnv As CDBEnvironment, ByVal pDT As CDBDataTable, ByVal pTrader As Boolean) As CDBDataTable

      If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation).Length > 0 AndAlso _
        pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation) = "Y" Then
        If pDT IsNot Nothing AndAlso pDT.Columns.ContainsKey("accreditation_status") Then

          For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1
            If Not CheckCentreAccreditationStatus(pEnv, pDT.Rows(vRowNumber).IntegerItem("exam_centre_id"), pTrader) Then
              pDT.RemoveRow(pDT.Rows(vRowNumber))
            End If
          Next
        End If
      Else
        Return pDT
      End If
      Return pDT
    End Function

    ''' <summary>
    ''' Check If the Centers are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Environment class</param>
    ''' <param name="pCentreId">Centre Id </param>
    ''' <param name="pTrader">Trader flag</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckCentreAccreditationStatus(ByVal pEnv As CDBEnvironment, ByVal pCentreId As Integer, ByVal pTrader As Boolean) As Boolean
      Dim vAnsiJoin As New AnsiJoins
      Dim vFields As String = "ec.accreditation_status,allow_registration,ignore_accreditation_validity,allow_result_entry,ec.accreditation_valid_from,ec.accreditation_valid_to"
      Dim vWhereClause As New CDBFields
      Dim vResult As Boolean = False


      vWhereClause.Add("ec.exam_centre_id", pCentreId)
      vAnsiJoin.Add("exam_accreditation_statuses acs", "acs.accreditation_status", "ec.accreditation_status")

      Dim vSql As New SQLStatement(pEnv.Connection, vFields, "exam_centres ec", vWhereClause, "", vAnsiJoin)
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQL(pEnv, vSql)

      If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then

        'CheckCentreAccreditationStatus If booking is allowed for centers, this should only be checked for trade application
        If pTrader AndAlso vDataTable.Rows(0).Item("allow_registration").Length > 0 AndAlso vDataTable.Rows(0).Item("allow_registration") = "Y" Then
          vResult = True
        Else
          vResult = False
        End If

        'Check if the dates are valid
        If vDataTable.Rows(0).Item("ignore_accreditation_validity") = "N" AndAlso vResult Then
          Dim vValidFrom As String = vDataTable.Rows(0).Item("accreditation_valid_from")
          Dim vValidTo As String = vDataTable.Rows(0).Item("accreditation_valid_to")

          If vValidFrom.Length > 0 AndAlso CDate(vValidFrom) > Date.Today Then
            vResult = False 'future
          ElseIf vValidTo.Length > 0 AndAlso CDate(vValidTo) < Date.Today Then
            vResult = False 'past 
          Else
            vResult = True
          End If

        ElseIf vDataTable.Rows(0).Item("ignore_accreditation_validity") = "Y" AndAlso vResult Then
          vResult = True
        Else
          vResult = False
        End If

        'Check if the result entry is allowed for the centre
        If Not pTrader AndAlso vResult Then
          If vDataTable.Rows(0).Item("allow_result_entry").Length > 0 AndAlso vDataTable.Rows(0).Item("allow_result_entry") = "Y" Then
            vResult = True
          Else
            vResult = False
          End If
        End If
      End If
      Return vResult
    End Function

    Private Sub GetExamUnitGradeHistory(ByVal pDataTable As CDBDataTable)
      Dim vRestrictResults As Boolean = False
      Dim vWhereFields As New CDBFields()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbWithholdExamResults) Then
        vWhereFields.Add("department", mvEnv.User.Department)
        If mvEnv.Connection.GetCount("exam_result_unrestricted_depts", vWhereFields) = 0 Then vRestrictResults = True
        vWhereFields.Clear()
      End If

      Dim vHistory As New ExamGradeChangeHistory(mvEnv)
      Dim vAttrs = {"egch.exam_grade_change_history_id,egch.exam_booking_unit_id,egch.exam_booking_id,egch.exam_unit_id,egch.exam_student_unit_header_id,egch.exam_grade_change_reason,egch.previous_mark,egch.previous_grade,egch.previous_result,egch.amended_by,egch.amended_on,exam_grade_change_reason_desc"}.AsCommaSeperated

      If mvParameters.ContainsKey("ExamBookingUnitId") Then vWhereFields.Add(If(vRestrictResults, "ebu.", "") + "exam_booking_unit_id", mvParameters("ExamBookingUnitId").IntegerValue)
      If mvParameters.ContainsKey("ExamStudentUnitHeaderId") Then vWhereFields.Add("egch.exam_student_unit_header_id", mvParameters("ExamStudentUnitHeaderId").IntegerValue)
      If mvParameters.ContainsKey("ExamUnitId") Then vWhereFields.Add("egch.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("exam_grade_change_reasons egcr", "egcr.exam_grade_change_reason", "egch.exam_grade_change_reason")

      If vRestrictResults Then
        With vWhereFields
          If mvParameters.ContainsKey("ContactNumber") Then .Add("ebu.contact_number", CDBField.FieldTypes.cftInteger, mvParameters("ContactNumber").Value, CDBField.FieldWhereOperators.fwoNullOrEqual)
          .Add("eses.results_release_date", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
          .Add("eses.results_release_date#2", CDBField.FieldTypes.cftInteger, "egch.amended_on", CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
          vAnsiJoins.Add("exam_booking_units ebu", "egch.exam_booking_unit_id", "ebu.exam_booking_unit_id")
          vAnsiJoins.Add("exam_bookings eb", "ebu.exam_booking_id", "eb.exam_booking_id")
          vAnsiJoins.AddLeftOuterJoin("exam_sessions eses", "eb.exam_session_id", "eses.exam_session_id")
          If mvParameters.ContainsKey("ExamStudentUnitHeaderId") Then
            .Add("esuh.exam_student_unit_header_id", mvParameters("ExamStudentUnitHeaderId").IntegerValue)
            vAnsiJoins.AddLeftOuterJoin("exam_student_header esh", "egch.exam_unit_id", "esh.exam_unit_id")
            vAnsiJoins.AddLeftOuterJoin("exam_student_unit_header esuh", "egch.exam_unit_id", "esuh.exam_unit_id", "esh.exam_student_header_id", "esuh.exam_student_header_id")
          End If
        End With
      End If

      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "exam_grade_change_history egch", vWhereFields, "egch.exam_grade_change_history_id DESC", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQL)
    End Sub
    ''' <summary>
    ''' Gets the communication_logs entry for Exam Units, Exam Centres or Exam Centre Units
    ''' </summary>
    ''' <param name="pDataTable">Data Table</param>
    ''' <remarks></remarks>
    Private Sub GetDocuments(ByVal pDataTable As CDBDataTable)

      'Exams Changes
      Dim vExamUnitLinks As Boolean = mvParameters.Exists("ExamUnitLinkId")
      Dim vExamCentre As Boolean = mvParameters.Exists("ExamCentreId")
      Dim vExamCentreUnit As Boolean = mvParameters.Exists("ExamCentreUnitId")

      'Dim vOutstandingDocs As Boolean
      'If mvParameters.Exists("Notified") Or mvParameters.Exists("Processed") Then vOutstandingDocs = True
      'Dim vHistoryDocs As Boolean = mvParameters.Exists("HistoryItems")
      'Dim vContactDocs As Boolean
      'If mvParameters.Exists("LinkType") Or mvParameters.Exists("ContactNumber") Then vContactDocs = True

      Dim vAttrs As String = "dated,cl.communications_log_number,cl.package,label_name,c.contact_number,document_type_desc,created_by,department_desc,our_reference,direction,their_reference,cl.document_type,cl.document_class,document_class_desc,standard_document,cl.source,recipient,forwarded,archiver,completed,cls.topic,topic_desc,cls.sub_topic,sub_topic_desc,creator_header,department_header,public_header,d.department,creator_header AS access_level,standard_document AS standard_document_desc"
      vAttrs &= ",precis,subject,call_duration,total_duration,selection_set"

      If Not mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing) Then vAttrs = vAttrs.Replace(",selection_set", ",")

      Dim vTables As New StringBuilder
      vTables.Append("document_log_links dl, communications_log cl, communications_log_subjects cls, contacts c, document_types dt, document_classes dc, departments d, topics t, sub_topics st")

      Dim vWhereFields As New CDBFields()
      If mvExamSelectionType = ExamDataSelectionTypes.dstExamUnitLinkDocuments Or mvExamSelectionType = ExamDataSelectionTypes.dstExamCentreDocuments Or mvExamSelectionType = ExamDataSelectionTypes.dstExamCentreUnitLinkDocuments Then   'DataSelectionTypes.dstDocuments Or mvExamSelectionType = DataSelectionTypes.dstDistinctDocuments Or mvExamSelectionType = DataSelectionTypes.dstDistinctExternalDocuments Or mvDataSelectionListType = DataSelectionTypes.dstEventDocuments Then

        If vExamUnitLinks Then
          vWhereFields.Add("dl.exam_unit_link_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamUnitLinkId").Value)
        ElseIf vExamCentreUnit Then
          vWhereFields.Add("dl.exam_centre_unit_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamCentreUnitId").Value)
        ElseIf vExamCentre Then
          vWhereFields.Add("dl.exam_centre_id", CDBField.FieldTypes.cftInteger, mvParameters("ExamCentreId").Value)
        End If
      End If
      vWhereFields.AddJoin("cl.communications_log_number#2", "cls.communications_log_number")
      vWhereFields.Add("primary", "Y").SpecialColumn = True
      vWhereFields.AddJoin("c.contact_number", "cl.contact_number")
      vWhereFields.AddJoin("dt.document_type", "cl.document_type")
      vWhereFields.AddJoin("dc.document_class", "cl.document_class")
      vWhereFields.AddJoin("d.department", "cl.department")
      vWhereFields.AddJoin("t.topic", "cls.topic")
      vWhereFields.AddJoin("st.topic", "t.topic")
      vWhereFields.AddJoin("st.sub_topic", "cls.sub_topic")
      If vExamCentre OrElse vExamCentreUnit OrElse vExamUnitLinks Then vWhereFields.AddJoin("cl.communications_log_number", "dl.communications_log_number")

      Dim vSqlStatement As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vAttrs), vTables.ToString, vWhereFields, "dated DESC, cl.communications_log_number DESC")
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

    Private Sub GetUnitStudyModes(ByVal pDataTable As CDBDataTable)
      Dim vOtherUnits As New List(Of String)

      Dim vSqlText As New StringBuilder()
      vSqlText.AppendLine("WITH CTE(id,")
      vSqlText.AppendLine("         parent_id)")
      vSqlText.AppendLine("AS   (SELECT exam_unit_link_id, ")
      vSqlText.AppendLine("             parent_unit_link_id ")
      vSqlText.AppendLine("      FROM   exam_unit_links ")
      vSqlText.AppendLine("      WHERE  parent_unit_link_id <> 0 ")
      vSqlText.AppendLine("      UNION ALL ")
      vSqlText.AppendLine("      SELECT exam_unit_link_id, ")
      vSqlText.AppendLine("             parent_id ")
      vSqlText.AppendLine("      FROM   exam_unit_links, CTE ")
      vSqlText.AppendLine("      WHERE  parent_unit_link_id = id) ")
      vSqlText.AppendLine("SELECT parent_id ")
      vSqlText.AppendLine("FROM   CTE ")
      vSqlText.AppendLine("WHERE  id = " & mvParameters("ExamUnitLinkId").Value & " ")
      vSqlText.AppendLine("ORDER BY parent_id")
      Dim vData As DataTable = (New SQLStatement(mvEnv.Connection, vSqlText.ToString)).GetDataTable
      vOtherUnits.AddRange(From vRow As DataRow In vData.AsEnumerable
                           Select CStr(vRow("parent_id")))

      vSqlText.Length = 0
      vSqlText.AppendLine("WITH CTE(id,")
      vSqlText.AppendLine("         parent_id)")
      vSqlText.AppendLine("AS   (SELECT exam_unit_link_id, ")
      vSqlText.AppendLine("             parent_unit_link_id ")
      vSqlText.AppendLine("      FROM   exam_unit_links ")
      vSqlText.AppendLine("      WHERE  parent_unit_link_id = " & mvParameters("ExamUnitLinkId").Value & " ")
      vSqlText.AppendLine("      UNION ALL ")
      vSqlText.AppendLine("      SELECT exam_unit_link_id, ")
      vSqlText.AppendLine("             parent_id ")
      vSqlText.AppendLine("      FROM   exam_unit_links, CTE ")
      vSqlText.AppendLine("      WHERE  parent_unit_link_id = id) ")
      vSqlText.AppendLine("SELECT id ")
      vSqlText.AppendLine("FROM   CTE ")
      vSqlText.AppendLine("WHERE  parent_id <> 0 AND ")
      vSqlText.AppendLine("       id <> " & mvParameters("ExamUnitLinkId").Value & " ")
      vSqlText.AppendLine("ORDER BY id")
      vData = (New SQLStatement(mvEnv.Connection, vSqlText.ToString)).GetDataTable
      vOtherUnits.AddRange(From vRow As DataRow In vData.AsEnumerable
                           Select CStr(vRow("id")))

      If vOtherUnits.Count = 0 OrElse New SQLStatement(mvEnv.Connection,
                        "study_mode",
                        "exam_unit_study_modes",
                        New CDBFields({New CDBField("exam_unit_link_id", vOtherUnits)})).GetDataTable.Rows.Count = 0 Then
        Dim vSql As New SQLStatement(mvEnv.Connection,
                                     "xsm.study_mode,xsm.study_mode_desc,CASE WHEN xusm.study_mode IS NULL THEN 'N' ELSE 'Y' END AS selected",
                                     "study_modes xsm",
                                     Nothing,
                                     "",
                                     New AnsiJoins({New AnsiJoin("exam_unit_study_modes xusm",
                                                                 "xsm.study_mode",
                                                                 "xusm.study_mode",
                                                                 "xusm.exam_unit_link_id",
                                                                 mvParameters("ExamUnitLinkId").Value,
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin)}))
        pDataTable.FillFromSQL(mvEnv, vSql)
      End If
    End Sub

    Private Sub GetCentreUnitStudyModes(ByVal pDataTable As CDBDataTable)
      Dim vSql As New SQLStatement(mvEnv.Connection,
                                     "xsm.study_mode,xsm.study_mode_desc, CASE WHEN xcusm.study_mode IS NULL THEN 'N' ELSE 'Y' END AS selected",
                                     "exam_unit_study_modes xusm",
                                     New CDBFields({New CDBField("xusm.exam_unit_link_id", mvParameters("ExamUnitLinkId").Value)}),
                                     "",
                                     New AnsiJoins({New AnsiJoin("study_modes xsm",
                                                                 "xsm.study_mode",
                                                                 "xusm.study_mode",
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin),
                                                    New AnsiJoin("exam_centre_unit_study_modes xcusm",
                                                                 "xcusm.study_mode",
                                                                 "xusm.study_mode",
                                                                 "xcusm.exam_centre_unit_link_id",
                                                                 mvParameters("ExamCentreUnitLinkId").Value,
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin)}))
      pDataTable.FillFromSQL(mvEnv, vSql)
    End Sub

    Private Sub GetExamBookingStudyModes(ByVal pDT As CDBDataTable)
      Dim vOtherUnits As New List(Of String)

      If mvParameters.ParameterExists("ExamBookingId").IntegerValue > 0 Then
        Dim vData As DataTable = (New SQLStatement(mvEnv.Connection, "ebu.exam_unit_link_id", "exam_bookings eb", _
                                  New CDBFields(New CDBField("ebu.exam_booking_id", mvParameters.ParameterExists("ExamBookingId").IntegerValue)), "", _
                                  New AnsiJoins({New AnsiJoin("exam_booking_units ebu", "eb.exam_booking_id", "ebu.exam_booking_id")}))).GetDataTable

        vOtherUnits.AddRange(From vRow As DataRow In vData.AsEnumerable
                             Select CStr(vRow("exam_unit_link_id")))
      Else
        vOtherUnits.Add(mvParameters("ExamUnitLinkId").Value)   'Inckude current Unit

        Dim vSqlText As New StringBuilder()
        vSqlText.AppendLine("WITH CTE(id,")
        vSqlText.AppendLine("         parent_id)")
        vSqlText.AppendLine("AS   (SELECT exam_unit_link_id, ")
        vSqlText.AppendLine("             parent_unit_link_id ")
        vSqlText.AppendLine("      FROM   exam_unit_links ")
        vSqlText.AppendLine("      WHERE  parent_unit_link_id <> 0 ")
        vSqlText.AppendLine("      UNION ALL ")
        vSqlText.AppendLine("      SELECT exam_unit_link_id, ")
        vSqlText.AppendLine("             parent_id ")
        vSqlText.AppendLine("      FROM   exam_unit_links, CTE ")
        vSqlText.AppendLine("      WHERE  parent_unit_link_id = id) ")
        vSqlText.AppendLine("SELECT parent_id ")
        vSqlText.AppendLine("FROM   CTE ")
        vSqlText.AppendLine("WHERE  id = " & mvParameters("ExamUnitLinkId").Value & " ")
        vSqlText.AppendLine("ORDER BY parent_id")
        Dim vData As DataTable = (New SQLStatement(mvEnv.Connection, vSqlText.ToString)).GetDataTable
        vOtherUnits.AddRange(From vRow As DataRow In vData.AsEnumerable
                             Select CStr(vRow("parent_id")))

        vSqlText.Length = 0
        vSqlText.AppendLine("WITH CTE(id,")
        vSqlText.AppendLine("         parent_id)")
        vSqlText.AppendLine("AS   (SELECT exam_unit_link_id, ")
        vSqlText.AppendLine("             parent_unit_link_id ")
        vSqlText.AppendLine("      FROM   exam_unit_links ")
        vSqlText.AppendLine("      WHERE  parent_unit_link_id = " & mvParameters("ExamUnitLinkId").Value & " ")
        vSqlText.AppendLine("      UNION ALL ")
        vSqlText.AppendLine("      SELECT exam_unit_link_id, ")
        vSqlText.AppendLine("             parent_id ")
        vSqlText.AppendLine("      FROM   exam_unit_links, CTE ")
        vSqlText.AppendLine("      WHERE  parent_unit_link_id = id) ")
        vSqlText.AppendLine("SELECT id ")
        vSqlText.AppendLine("FROM   CTE ")
        vSqlText.AppendLine("WHERE  parent_id <> 0 AND ")
        vSqlText.AppendLine("       id <> " & mvParameters("ExamUnitLinkId").Value & " ")
        vSqlText.AppendLine("ORDER BY id")
        vData = (New SQLStatement(mvEnv.Connection, vSqlText.ToString)).GetDataTable
        vOtherUnits.AddRange(From vRow As DataRow In vData.AsEnumerable
                             Select CStr(vRow("id")))
      End If

      Dim vWhereFields As New CDBFields(New CDBField("exam_unit_link_id", vOtherUnits))
      Dim vAnsiJoins As New AnsiJoins()
      Dim vTableName As String = "exam_unit_study_modes esm"
      If mvParameters.ParameterExists("ExamCentreId").IntegerValue > 0 Then
        'If StudyModes exist at Centre level then need to use differnt tables, so do a count to find out
        vWhereFields.Add("ecu.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
        vAnsiJoins.Add("exam_centre_unit_study_modes esm", "ecu.exam_centre_unit_id", "esm.exam_centre_unit_link_id")
        Dim vCountSQL As New SQLStatement(mvEnv.Connection, "", "exam_centre_units ecu", vWhereFields, "", vAnsiJoins)
        If mvEnv.Connection.GetCountFromStatement(vCountSQL) > 0 Then
          vTableName = "exam_centre_units ecu"
        Else
          vAnsiJoins.Clear()
          vWhereFields.Remove("ecu.exam_centre_id")
        End If
      End If
      vAnsiJoins.Add("study_modes sm", "esm.study_mode", "sm.study_mode")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "DISTINCT sm.study_mode,sm.study_mode_desc", vTableName, vWhereFields, "sm.study_mode_desc", vAnsiJoins)
      pDT.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub GetUnitCertRunTypes(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      If mvParameters.ContainsKey("ExamUnitLinkId") Then
        vWhereFields.Add(New CDBField("xucr.exam_unit_link_id", mvParameters("ExamUnitLinkId").Value))
      End If
      If mvParameters.ContainsKey("ExamCertRunType") Then
        vWhereFields.Add(New CDBField("xucr.exam_cert_run_type", mvParameters("ExamCertRunType").Value))
      End If
      Dim vSql As New SQLStatement(mvEnv.Connection,
                                     "xucr.exam_unit_cert_run_type_id,xucr.exam_unit_link_id,xucr.exam_cert_run_type,xcr.exam_cert_run_type_desc,xucr.include_view,iv.view_name_desc,xucr.exclude_view,xv.view_name_desc,xucr.standard_document,sd.standard_document_desc",
                                     "exam_unit_cert_run_types xucr",
                                     vWhereFields,
                                     "",
                                     New AnsiJoins({New AnsiJoin("exam_cert_run_types xcr",
                                                                 "xcr.exam_cert_run_type",
                                                                 "xucr.exam_cert_run_type",
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin),
                                                    New AnsiJoin("view_names iv",
                                                                 "iv.view_name",
                                                                 "xucr.include_view",
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin),
                                                    New AnsiJoin("view_names xv",
                                                                 "xv.view_name",
                                                                 "xucr.exclude_view",
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin),
                                                    New AnsiJoin("standard_documents sd",
                                                                 "sd.standard_document",
                                                                 "xucr.standard_document",
                                                                 AnsiJoin.AnsiJoinTypes.LeftOuterJoin)}))
      pDataTable.FillFromSQL(mvEnv, vSql)
    End Sub

    Private Sub GetCertReprintTypes(ByVal pDataTable As CDBDataTable)
      Dim vSql As New SQLStatement(mvEnv.Connection,
                                     "xcrt.exam_cert_reprint_type,xcrt.exam_cert_reprint_type_desc",
                                     "exam_cert_reprint_types xcrt")
      pDataTable.FillFromSQL(mvEnv, vSql)
    End Sub

    Private Sub AddWhereFieldFromParameter(ByVal pWhereFields As CDBFields, ByVal pParameterName As String, ByVal pFieldName As String)
      If mvParameters.Exists(pParameterName) Then pWhereFields.Add(pFieldName, mvParameters(pParameterName).Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
    End Sub

    Private Sub GetExamSessionLookup(ByVal pDataTable As CDBDataTable)
      Dim vAttrs As String = "exam_session_code,exam_session_description,es.exam_session_id," & mvEnv.Connection.DBIsNull("es.sequence_number", "10000") & ",exam_session_year,exam_session_month"
      If mvParameters.ParameterExists("ExamCentreId").IntegerValue > 0 Then vAttrs = "DISTINCT " & vAttrs

      Dim vOrderBy As String = mvEnv.Connection.DBIsNull("es.sequence_number", "10000") & ",exam_session_year DESC,exam_session_month DESC"

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("es.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("es.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      If mvParameters.ParameterExists("Trader").Bool = False AndAlso mvParameters.ParameterExists("ExamSessionId").IntegerValue > 0 Then vWhereFields.Add("es.exam_session_id", mvParameters("ExamSessionId").Value)

      Dim vAnsiJoins As New AnsiJoins
      If mvParameters.ParameterExists("ExamUnitId").IntegerValue > 0 Then
        Dim vSubWhere As New CDBFields
        vSubWhere.Add("eu1.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
        vSubWhere.Add("eu1.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vSubWhere.Add("eu1.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
        Dim vSubAnsi As New AnsiJoins
        vSubAnsi.Add("exam_units eu2", "eu1.exam_unit_code", "eu2.exam_unit_code")
        Dim vSubSQL As New SQLStatement(mvEnv.Connection, "eu2.exam_unit_id", "exam_units eu1", vSubWhere, "", vSubAnsi)
        vAnsiJoins.Add("exam_units eu", "es.exam_session_id", "eu.exam_session_id")
        vWhereFields.Add("eu.exam_unit_id", CDBField.FieldTypes.cftInteger, vSubSQL.SQL, CDBField.FieldWhereOperators.fwoIn)
        If (mvParameters.ParameterExists("ExamSessionCode").Value.Equals("NONSESSION", System.StringComparison.CurrentCultureIgnoreCase)) Then
          vWhereFields.Add("eu.session_based", "N")
        End If
      End If

      If mvParameters.ParameterExists("ExamCentreId").IntegerValue > 0 Then
        vAnsiJoins.Add("exam_schedule esh", "es.exam_session_id", "esh.exam_session_id")
        vAnsiJoins.Add("exam_centres ec", "esh.exam_centre_id", "ec.exam_centre_id")
        vWhereFields.Add("ec.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAttrs, "exam_sessions es", vWhereFields, vOrderBy, vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      Dim vAddNonSession As Boolean = ((mvParameters.ParameterExists("ExamSessionId").IntegerValue = 0 OrElse mvParameters.ParameterExists("Trader").Bool = True) AndAlso mvParameters.ParameterExists("NonSessionBased").Bool = True)
      If vAddNonSession = True AndAlso mvParameters.ParameterExists("ExamUnitId").IntegerValue > 0 Then
        'If we have an ExamUnitId, only add the NONSESSION row if the unit is available as non-session
        Dim vNSWhere As New CDBFields(New CDBField("eu1.exam_unit_id", mvParameters("ExamUnitId").IntegerValue))
        With vNSWhere
          .Add("eu1.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
          .Add("eu1.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
          .Add("eu2.session_based", "N")
        End With
        Dim vNSAnsi As New AnsiJoins({New AnsiJoin("exam_units eu2", "eu1.exam_unit_code", "eu2.exam_unit_code")})
        Dim vNSSQL As New SQLStatement(mvEnv.Connection, "", "exam_units eu1", vNSWhere, "", vNSAnsi)
        If mvEnv.Connection.GetCountFromStatement(vNSSQL) = 0 Then vAddNonSession = False
      End If

      If vAddNonSession Then
        Dim vNonSessionRow As CDBDataRow = pDataTable.InsertRow(0)
        vNonSessionRow.Item("ExamSessionCode") = "NONSESSION"
        vNonSessionRow.Item("ExamSessionDescription") = "Non-Session Based"
        vNonSessionRow.Item("ExamSessionId") = "0"
      End If
    End Sub

    Private Sub GetExamCentreLookup(ByVal pDataTable As CDBDataTable)
      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields

      Dim vBySession As Boolean = (mvParameters.ParameterExists("ExamSessionId").IntegerValue > 0)
      Dim vNonSessionBased As Boolean = (mvParameters.ParameterExists("ExamSessionCode").Value.Equals("NONSESSION", System.StringComparison.CurrentCultureIgnoreCase))

      Dim vFields As String = "DISTINCT exam_centre_code,exam_centre_description,ec.exam_centre_id,ec.accreditation_status,ec.accreditation_valid_from,ec.accreditation_valid_to,ec.overseas"
      If vBySession Then
        vFields += ",es.home_closing_date,es.overseas_closing_date,'' AS closing_date"
      Else
        vFields += ",'' AS home_closing_date,'' AS overseas_closing_date,'' AS closing_date"
      End If

      If mvParameters.ContainsKey("ExamCentreCode") Then
        vWhereFields.Add("ec.exam_centre_code", mvParameters("ExamCentreCode").Value)
      Else
        vWhereFields.Add("ec.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("ec.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      End If
      If mvParameters.ParameterExists("Trader").Bool = False AndAlso mvParameters.HasValue("ExamCentreId") Then vWhereFields.Add("ec.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)

      Dim vByUnit As Boolean = mvParameters.HasValue("ExamUnitId")
      If vByUnit Then
        Dim vSubWhere As New CDBFields
        vSubWhere.Add("eu1.exam_unit_id", mvParameters("ExamUnitId").IntegerValue)
        Dim vSubAnsi As New AnsiJoins
        vSubAnsi.Add("exam_units eu2", "eu1.exam_unit_code", "eu2.exam_unit_code")
        Dim vSubSQL As New SQLStatement(mvEnv.Connection, "eu2.exam_unit_id", "exam_units eu1", vSubWhere, "", vSubAnsi)
        vAnsiJoins.Add("exam_centre_units ecu", "ec.exam_centre_id", "ecu.exam_centre_id")
        vWhereFields.Add("ecu.exam_unit_id", CDBField.FieldTypes.cftInteger, vSubSQL.SQL, CDBField.FieldWhereOperators.fwoIn)
      End If

      'Join to Sessions
      If vBySession Then
        vAnsiJoins.Add("exam_session_centres esc", "esc.exam_centre_id", "ec.exam_centre_id")
        vAnsiJoins.Add("exam_sessions es", "es.exam_session_id", "esc.exam_session_id")
        vWhereFields.Add("esc.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      End If

      'Join to Exam Units
      If mvParameters.ParameterExists("Trader").Bool = True AndAlso vNonSessionBased = True Then
        'Look for non-session based
        'exam-centre-units to exam-units for session-based = N
        If vByUnit = False Then vAnsiJoins.Add("exam_centre_units ecu", "ec.exam_centre_id", "ecu.exam_centre_id") 'May have already added the join
        vAnsiJoins.Add("exam_units eu", "ecu.exam_unit_id", "eu.exam_unit_id")
        vWhereFields.Add("eu.session_based", "N")
      End If

      ' Centre Location Filtering
      If mvParameters.HasValue("ExamCentreLocation") Then
        If mvParameters("ExamCentreLocation").Value.Equals("H", System.StringComparison.CurrentCultureIgnoreCase) Then
          vWhereFields.Add("ec.overseas", "N", CDBField.FieldWhereOperators.fwoNullOrEqual)
        ElseIf mvParameters("ExamCentreLocation").Value.Equals("O", System.StringComparison.CurrentCultureIgnoreCase) Then
          vWhereFields.Add("ec.overseas", "Y")
        End If
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "exam_centres ec", vWhereFields, "exam_centre_description", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      If vBySession Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          If (vRow.Item("Overseas") = "Y" AndAlso vRow.Item("OverseasClosingDate") <> "") Then
            vRow.Item("ClosingDate") = vRow.Item("OverseasClosingDate")
          ElseIf vRow.Item("HomeClosingDate") <> "" Then
            vRow.Item("ClosingDate") = vRow.Item("HomeClosingDate")
          End If
        Next
      End If

      'Handle accreditation
      If (mvParameters.ContainsKey("Trader") OrElse mvParameters.ContainsKey("ResultEntry")) AndAlso pDataTable.Rows.Count > 0 Then
        'Build a List of all Exam Centres.  We will gradually filter-out all un-accredited Centres and then remove all from the final list
        Dim vExamCentres As New List(Of Integer)
        pDataTable.Rows.ForEach(Sub(vRow) vExamCentres.Add(vRow.IntegerItem("ExamCentreId")))

        Dim vRemoveIds As New List(Of Integer)

        Dim vAccreditationFilterType As ExamAccreditationFilterTypes = ExamAccreditationFilterTypes.eafAllowRegistration
        If mvParameters.ContainsKey("ResultEntry") Then vAccreditationFilterType = ExamAccreditationFilterTypes.eafAllowResultEntry
        Dim vAccreditationFilter As New ExamCentreAccreditationFilter(mvEnv, vExamCentres, vAccreditationFilterType)

        'Filter-out un-accredited Exam Centres
        Dim vCentreAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation))
        If vCentreAccreditationIsEnabled Then
          Dim vRemoveUnaccreditedCentres As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentres()
          If vRemoveUnaccreditedCentres IsNot Nothing Then vRemoveIds.AddRange(vRemoveUnaccreditedCentres)
        End If

        If vByUnit Then
          Dim vUnitAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamUnitAccreditation))
          If vUnitAccreditationIsEnabled Then
            Dim vRemoveUnaccreditedCentres As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentresForUnit(mvParameters("ExamUnitId").IntegerValue)
            If vRemoveUnaccreditedCentres IsNot Nothing Then vRemoveIds.AddRange(vRemoveUnaccreditedCentres)
          End If

          Dim vCentreUnitAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation))
          If vCentreUnitAccreditationIsEnabled Then
            Dim vSessionId As Integer = 0
            If vBySession Then vSessionId = mvParameters("ExamSessionId").IntegerValue
            Dim vRemoveUnaccreditedCentres As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentresForCentreUnit(mvParameters("ExamUnitId").IntegerValue, vSessionId)
            If vRemoveUnaccreditedCentres IsNot Nothing Then vRemoveIds.AddRange(vRemoveUnaccreditedCentres)
          End If
        End If

        If vRemoveIds IsNot Nothing AndAlso vRemoveIds.Count > 0 Then
          Dim vRemoveRows As List(Of CDBDataRow) = pDataTable.Rows.FindAll(Function(vRow) vRemoveIds.Contains(vRow.IntegerItem("ExamCentreId")))
          vRemoveRows.ForEach(Sub(vRow) pDataTable.RemoveRow(vRow))
        End If
      End If
    End Sub

    Private Sub GetExamUnitLookup(ByVal pDataTable As CDBDataTable)
      Dim vAnsiJoins As New AnsiJoins
      Dim vWhereFields As New CDBFields

      Dim vFields As String = "DISTINCT exam_unit_code,"
      If mvParameters.HasValue("ExamCentreId") AndAlso mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
        vFields &= mvEnv.Connection.DBIsNull("ecu.local_name", "exam_unit_description") & " AS exam_unit_description"
      Else
        vFields &= "exam_unit_description"
      End If
      If mvParameters.ContainsKey("FilterLookUp") Then  'BR19602 - Added to prevent multiples in dropdown list
        vFields &= ",NULL,NULL"
      Else
        vFields &= ",eu.exam_unit_id, eul.exam_unit_link_id"
      End If
      'These additional attributes are required for the accreditation filtering below and will be removed afterwards
      Dim vAdditionalFields As String = String.Empty
      If mvParameters.ContainsKey("Trader") OrElse mvParameters.ContainsKey("ResultEntry") Then
        vAdditionalFields = "DISTINCT eu.exam_unit_id"
        If mvParameters.HasValue("ExamCentreId") Then
          vAdditionalFields &= ",ecu.accreditation_status,ecu.accreditation_valid_from,ecu.accreditation_valid_to,ecu.exam_centre_unit_id"
        Else
          vAdditionalFields &= ",eul.accreditation_status,eul.accreditation_valid_from,eul.accreditation_valid_to,0 AS exam_centre_unit_id"
        End If
        vAdditionalFields &= ",eul.exam_unit_link_id,eul.parent_unit_link_id"
      End If

      vWhereFields.Add("eu.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("eu.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)

      Dim vUseSession As Boolean = (mvParameters.ParameterExists("ExamSessionId").IntegerValue > 0)
      Dim vNonSessionBased As Boolean = (mvParameters.ParameterExists("ExamSessionCode").Value.Equals("NONSESSION", System.StringComparison.CurrentCultureIgnoreCase))

      If vUseSession Then
        vWhereFields.Add("eu.exam_session_id", mvParameters("ExamSessionId").IntegerValue)
      ElseIf vNonSessionBased Then
        vWhereFields.Add("eu.exam_session_id", CDBField.FieldTypes.cftInteger)
        vWhereFields.Add("eu.session_based", "N")
      End If

      If mvParameters.HasValue("ExamUnitCode") Then
        vWhereFields.Add("exam_unit_code", mvParameters("ExamUnitCode").Value)
      ElseIf Not mvParameters.HasValue("AllUnits") Then
        vWhereFields.Add("eul.exam_unit_id_1", CDBField.FieldTypes.cftInteger, "0", CDBField.FieldWhereOperators.fwoNullOrEqual)
      End If

      If mvParameters.ParameterExists("ExcludeQuestionUnits").Bool Then
        ' Exclude units with unit types marked as questions
        vAnsiJoins.Add("exam_unit_types eut", "eu.exam_unit_type", "eut.exam_unit_type")
        vWhereFields.Add("eut.exam_question", "Y", CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If

      If mvParameters.ParameterExists("AllowMarkEntry").Bool Then
        ' Exclude marker status of 'N' - No Markers Required (No Result Entry)
        vWhereFields.Add("eu.exam_marker_status", CDBField.FieldTypes.cftCharacter, "N", CARE.Data.CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      End If

      If mvParameters.HasValue("AllowBookings") Then vWhereFields.Add("allow_bookings", mvParameters("AllowBookings").Value)
      If mvParameters.HasValue("WebPublish") Then vWhereFields.Add("web_publish", mvParameters("WebPublish").Value)
      If mvParameters.HasValue("SessionBased") Then vWhereFields.Add("session_based", mvParameters("SessionBased").Value)
      Dim vLinkToCentre As Boolean = mvParameters.HasValue("ExamCentreId")
      If vLinkToCentre Then
        If vUseSession Then
          vAnsiJoins.Add("exam_centre_units ecu", "eu.exam_base_unit_id", "ecu.exam_unit_id")
        Else
          vAnsiJoins.Add("exam_centre_units ecu", "eu.exam_unit_id", "ecu.exam_unit_id")
        End If
        vAnsiJoins.Add("exam_centres ec", "ecu.exam_centre_id", "ec.exam_centre_id")
        vWhereFields.Add("ec.exam_centre_id", mvParameters("ExamCentreId").IntegerValue)
        vWhereFields.Add("ec.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("ec.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      End If

      If Not mvParameters.ContainsKey("FilterLookUp") Then vAnsiJoins.AddLeftOuterJoin("exam_unit_links eul", "eu.exam_unit_id", "eul.exam_unit_id_2") 'BR19602

      If (mvParameters.ContainsKey("Trader") OrElse mvParameters.ContainsKey("ResultEntry")) Then
        Dim vDataTable As New CDBDataTable()
        Dim vSubSQLStatement As New SQLStatement(mvEnv.Connection, vAdditionalFields, "exam_units eu", vWhereFields, "exam_unit_link_id", vAnsiJoins)
        vDataTable.FillFromSQL(mvEnv, vSubSQLStatement)
        vDataTable.RemoveDuplicateRows("exam_unit_link_id")

        If vDataTable.Rows.Count > 0 Then
          Dim vFilter As ExamAccreditationFilterTypes = ExamAccreditationFilterTypes.eafAllowRegistration
          If mvParameters.ParameterExists("ResultEntry").Bool Then vFilter = ExamAccreditationFilterTypes.eafAllowResultEntry

          Dim vRemoveIds As New List(Of Integer)
          'Build a dictionary of all exam unit links and their parents.  We will gradually filter-out all un-accredited units and then remove all from the final list
          Dim vExamUnitLinks As New Dictionary(Of Integer, Integer)
          vDataTable.Rows.ForEach(Sub(vRow) vExamUnitLinks.Add(vRow.IntegerItem("exam_unit_link_id"), vRow.IntegerItem("parent_unit_link_id")))

          Dim vAccreditationFilter As New ExamAccreditationFilter(mvEnv, vExamUnitLinks, vFilter) 'This class handles all the accreditation filtering

          'Filter-out un-accredited Exam Units
          Dim vUnitAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamUnitAccreditation))
          If vUnitAccreditationIsEnabled Then
            Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnaccreditedUnits()
            vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
          End If

          'Filter-out unaccredited Centres and unaccredited centre-units
          If mvParameters.HasValue("ExamCentreId") Then
            Dim vSessionID As Integer = mvParameters.ParameterExists("ExamSessionId").IntegerValue
            'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
            Dim vCentreAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreAccreditation))
            If vCentreAccreditationIsEnabled Then
              Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnitsAtUnaccreditedCentre(mvParameters("ExamCentreId").IntegerValue, vSessionID)
              vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
            End If

            'Remove all Centre rows that are not accredited (or specifically whose accreditation status doesn't allow bookings)
            Dim vCentreUnitAccreditationIsEnabled As Boolean = BooleanValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation))
            If vCentreUnitAccreditationIsEnabled Then
              Dim vRemoveUnaccreditedUnits As List(Of Integer) = vAccreditationFilter.GetUnaccreditedCentreUnits(mvParameters("ExamCentreId").IntegerValue, vSessionID)
              vRemoveIds.AddRange(vRemoveUnaccreditedUnits)
            End If
          End If

          If vRemoveIds IsNot Nothing AndAlso vRemoveIds.Count > 0 Then
            Dim vRemoveRows As List(Of CDBDataRow) = vDataTable.Rows.FindAll(Function(vRow) vRemoveIds.Contains(vRow.IntegerItem("exam_unit_link_id")))
            vRemoveRows.ForEach(Sub(vRow) vDataTable.RemoveRow(vRow))
          End If

          'Having filtered the unaccredited data, reselect
          Dim vExamUnitIds As New List(Of Integer)
          vDataTable.Rows.ForEach(Sub(vRow) vExamUnitIds.Add(vRow.IntegerItem("exam_unit_id")))
          If vExamUnitIds.Count = 0 Then
            'We want to select no data as nothing is accredited
            vWhereFields.Add("eu.exam_unit_code", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoEqual)
          Else
            'Only select the accredited units
            vWhereFields.Add("eu.exam_unit_id", CDBField.FieldTypes.cftInteger, vExamUnitIds.AsCommaSeperated, CDBField.FieldWhereOperators.fwoIn)
          End If
        End If
      End If

      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, "exam_units eu", vWhereFields, "exam_unit_code", vAnsiJoins)
      pDataTable.FillFromSQL(mvEnv, vSQLStatement)
    End Sub

    Private Sub CheckCustomForms(ByVal pFormUsageCode As String)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("client", mvEnv.ClientCode)
      vWhereFields.Add("custom_form", mvEnv.FirstCustomFormNumber, CDBField.FieldWhereOperators.fwoBetweenFrom Or CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("custom_form#2", mvEnv.LastCustomFormNumber, CDBField.FieldWhereOperators.fwoBetweenTo)
      vWhereFields.Add("custom_form#3", 1000, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoGreaterThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
      vWhereFields.Add("form_usage_code", pFormUsageCode) 'Only get Exam Centre Custom Forms
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "custom_form,form_caption", "custom_forms", vWhereFields, "custom_form").GetRecordSet
      While vRecordSet.Fetch()
        mvResultColumns = mvResultColumns & ",CustomForm" & vRecordSet.Fields(1).Value
        mvHeadings = mvHeadings & "," & vRecordSet.Fields(2).Value
        mvWidths = mvWidths & ",300"
      End While
      vRecordSet.CloseRecordSet()
    End Sub

  End Class

End Namespace
