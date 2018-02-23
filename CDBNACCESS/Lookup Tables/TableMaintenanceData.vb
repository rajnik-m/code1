Imports System.Xml.Linq
Namespace Access

  Public Class TableMaintenanceData
    Inherits CARERecord

#Region "Private Members"

    Private mvConfirmDelete As Boolean
    Private mvStockFlagAfter As Boolean
    Private mvSetMembershipGroups As Boolean
    Private mvOrgGroup As String
    Private mvMemberTypeCode As String
    Private mvHistoric As Boolean
    Private mvMode As MaintenanceTypes
    Private mvTableMaintenance As Boolean

#End Region

#Region "Constructor"

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    Public Sub New(ByVal pEnv As CDBEnvironment, pMaintenanceTable As String)
      MyClass.New(pEnv)
      InitFromMaintenanceData(pMaintenanceTable)
    End Sub

#End Region

    Protected Overrides Sub AddFields()
      'Should not be called
    End Sub

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return mvClassFields.DatabaseTableName
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return mvClassFields.ContainsKey("amended_on") AndAlso mvClassFields.ContainsKey("amended_by")
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return mvClassFields.TableAlias
      End Get
    End Property

    Public Overrides ReadOnly Property NeedsMaintenanceInfo() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides ReadOnly Property IsMaintenanceTable() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides ReadOnly Property NoUniqueKey() As Boolean
      Get
        Dim vNoUniqueKey As Boolean = MyBase.NoUniqueKey
        If vNoUniqueKey = False Then
          Dim vUpdatableFields As Integer
          Dim vPrimaryKeys As Boolean = False
          For Each vClassField As ClassField In mvClassFields
            If vClassField.PrimaryKey Then
              vPrimaryKeys = True
            ElseIf vClassField.Name <> "amended_on" AndAlso vClassField.Name <> "amended_by" Then
              vUpdatableFields += 1
            End If
          Next
          If vPrimaryKeys AndAlso vUpdatableFields = 0 Then
            vNoUniqueKey = True
            mvClassFields.SetUniqueFieldsFromPrimaryKeys()                      'Retain the primary keys for unique check if any change
            For Each vClassField As ClassField In mvClassFields
              If vClassField.PrimaryKey Then
                vClassField.PrimaryKey = False
              End If
            Next
          End If
        End If
        Return vNoUniqueKey
      End Get
    End Property

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return GetUniqueKeyFieldNames()
    End Function

    Public Overrides Sub PreValidateParameterList(ByVal pType As CARERecord.MaintenanceTypes, ByVal pParameterList As CDBParameters)
      MyBase.PreValidateParameterList(pType, pParameterList)
      mvMode = pType
      If pParameterList.ContainsKey("TableMaintenance") Then mvTableMaintenance = pParameterList("TableMaintenance").Bool

      'Read the values and remove them from the param list as they are not part of the maint
      'attributes and will raise an error
      If pParameterList.ContainsKey("ConfirmDelete") Then
        mvConfirmDelete = pParameterList("ConfirmDelete").Bool
        pParameterList.Remove("ConfirmDelete")
      End If
      If pParameterList.ContainsKey("StockFlagAfter") Then
        mvStockFlagAfter = pParameterList("StockFlagAfter").Bool
        pParameterList.Remove("StockFlagAfter")
      End If
      If pParameterList.ContainsKey("SetMembershipGroups") Then
        mvSetMembershipGroups = pParameterList("SetMembershipGroups").Bool
        pParameterList.Remove("SetMembershipGroups")
      End If
      If pParameterList.ContainsKey("OrgGroup") Then
        mvOrgGroup = pParameterList("OrgGroup").Value
        pParameterList.Remove("OrgGroup")
      End If
      If pParameterList.ContainsKey("MemberTypeCode") Then
        mvMemberTypeCode = pParameterList("MemberTypeCode").Value
        pParameterList.Remove("MemberTypeCode")
      End If
      If pParameterList.ContainsKey("Historic") Then
        mvHistoric = pParameterList("Historic").Bool
        pParameterList.Remove("Historic")
      End If

      Select Case pParameterList("MaintenanceTableName").Value.ToLower
        Case "cpd_cycle_types"
          If Not ((pParameterList.HasValue("StartMonth") = True AndAlso pParameterList.HasValue("EndMonth") = True) _
          OrElse (pParameterList.HasValue("StartMonth") = False AndAlso pParameterList.HasValue("EndMonth") = False)) Then
            RaiseError(DataAccessErrors.daeCPDCycleTypesStartEndMonthsInvalid)
          Else
            If pType = MaintenanceTypes.Update AndAlso CanUpdateCPDCycles(pParameterList) = False Then
              RaiseError(DataAccessErrors.daeCPDCycleTypesStartEndMonthsCannotChange)
            End If
          End If

        Case "explorer_link_access_levels"
          If Not pParameterList.ContainsKey("Client") Then pParameterList.Add("Client", "")
          If Not pParameterList.ContainsKey("Department") Then pParameterList.Add("Department", "")
          If Not pParameterList.ContainsKey("Logname") Then pParameterList.Add("Logname", "")

        Case "organisation_groups"
          If pParameterList.ContainsKey("ViewInContactCard") Then
            Dim vOrgGroup As String = pParameterList.ParameterExists("OrganisationGroup").Value
            If vOrgGroup.Equals("ORG", StringComparison.InvariantCultureIgnoreCase) AndAlso pParameterList("ViewInContactCard").Bool = True Then
              RaiseError(DataAccessErrors.daeCannotSetViewInContactCardForGroup, vOrgGroup)
            End If
          End If

        Case "packages"
          If pParameterList.ContainsKey("StorageType") AndAlso pParameterList.ContainsKey("StoragePath") Then
            If Not String.IsNullOrWhiteSpace(pParameterList("StorageType").Value) AndAlso pParameterList("StorageType").Value.Equals("E", StringComparison.InvariantCultureIgnoreCase) Then
              'External storage
              Dim vExternalPath As String = pParameterList("StoragePath").Value.Trim
              If vExternalPath.Length = 0 Then
                RaiseError(DataAccessErrors.daeStoragePathMustBeSetForExternal)
              ElseIf vExternalPath.StartsWith("\\") = False Then
                RaiseError(DataAccessErrors.daeUNCPathOnly)
              End If
            End If
          End If

        Case "payment_frequencies"
          If (pType = MaintenanceTypes.Insert OrElse pType = MaintenanceTypes.Update) AndAlso pParameterList.HasValue("OffsetMonths") Then
            Dim vFrequency As Integer = pParameterList("Frequency").IntegerValue
            Dim vInterval As Integer = pParameterList("Interval").IntegerValue
            Dim vOffsetMonths As Integer = pParameterList("OffsetMonths").IntegerValue
            Dim vMaxOffset As Integer = 0
            If PaymentFrequency.IsOffsetMonthsValid(pParameterList("Period").Value, vFrequency, vInterval, vOffsetMonths, vMaxOffset) = False Then
              RaiseError(DataAccessErrors.daePayFrequencyOffsetMonthsInvalid, vMaxOffset.ToString)
            End If
          End If

        Case "mailings"
          If (pType = MaintenanceTypes.Insert OrElse pType = MaintenanceTypes.Update) Then
            'Validate Mailing parameters
            Dim vMailing As New Mailing(mvEnv)
            vMailing.Init(pParameterList("Mailing").Value)
            vMailing.PreValidateParameterList(pType, pParameterList)
          End If

        Case "config"
          If Not pParameterList.ContainsKey("Department") Then pParameterList.Add("Department", "") 'class treats department and logname as PK values so must always be supplied even if empty
          If Not pParameterList.ContainsKey("Logname") Then pParameterList.Add("Logname", "")
          If pParameterList.ContainsKey("ConfigName") Then
            Dim vConfigScope As Config.ConfigNameScope = mvEnv.GetConfigScopeLevel(pParameterList("ConfigName").Value)
            If pParameterList.HasValue("Department") AndAlso (vConfigScope And Config.ConfigNameScope.Department) < Config.ConfigNameScope.Department Then RaiseError(DataAccessErrors.daeScopeLevelError, ProjectText.LangDepartment)
            If pParameterList.HasValue("Logname") AndAlso (vConfigScope And Config.ConfigNameScope.User) < Config.ConfigNameScope.User Then RaiseError(DataAccessErrors.daeScopeLevelError, ProjectText.LangUser)
          End If
        Case "rate_nominal_accounts"
          'nominal_account_suffix Attribute allows nulls and is part of the primary key
          'add param if it doesn't exist 
          If Not pParameterList.ContainsKey("NominalAccountSuffix") Then pParameterList.Add("NominalAccountSuffix", "")

        Case "prize_draws"
          If pType = MaintenanceTypes.Insert OrElse pType = MaintenanceTypes.Update Then
            Dim vCloseDate As Nullable(Of Date) = Nothing
            If IsDate(pParameterList.ParameterExists("CloseDate").Value) Then
              vCloseDate = DateValue(pParameterList("CloseDate").Value)
            End If
            Dim vDrawDate As Date = DateValue(pParameterList("DrawDate").Value)
            If vCloseDate.HasValue = False OrElse vDrawDate.CompareTo(vCloseDate.Value) > 0 Then
              RaiseError(DataAccessErrors.daeParameterValueInvalid, "Close Date", pParameterList.ParameterExists("CloseDate").Value)
            End If
          End If

      End Select

      If pType = MaintenanceTypes.Insert Then
        'Get control values for the fields which are auto generated. Add the values to the param list as the 
        'Process maintenance data method checks for the params and its also required for initialisation of the care record class.
        Select Case pParameterList("MaintenanceTableName").Value
          Case "event_child_discount_levels"
            If pParameterList.ContainsKey("EventChildDiscountNumber") Then pParameterList.Remove("EventChildDiscountNumber")
            pParameterList.Add("EventChildDiscountNumber", mvEnv.GetControlNumber("EC").ToString)
          Case "event_extra_fee_multipliers"
            If pParameterList.ContainsKey("EventExtraFeeNumber") Then pParameterList.Remove("EventExtraFeeNumber")
            pParameterList.Add("EventExtraFeeNumber", mvEnv.GetControlNumber("EE").ToString)
          Case "event_fee_band_discounts"
            If pParameterList.ContainsKey("EventFeeBandDiscountNumber") Then pParameterList.Remove("EventFeeBandDiscountNumber")
            pParameterList.Add("EventFeeBandDiscountNumber", mvEnv.GetControlNumber("EB").ToString)
          Case "event_fees"
            If pParameterList.ContainsKey("EventFeeNumber") Then pParameterList.Remove("EventFeeNumber")
            pParameterList.Add("EventFeeNumber", mvEnv.GetControlNumber("EF").ToString)
          Case "internal_resources"
            If pParameterList.ContainsKey("ResourceNumber") Then pParameterList.Remove("ResourceNumber")
            pParameterList.Add("ResourceNumber", mvEnv.GetControlNumber("RN").ToString)
          Case "service_control_start_days", "service_start_days"
            If pParameterList.ContainsKey("StartDayNumber") Then pParameterList.Remove("StartDayNumber")
            pParameterList.Add("StartDayNumber", mvEnv.GetControlNumber("SD").ToString)
          Case "service_control_restrictions"
            If pParameterList.ContainsKey("ServiceRestrictionNumber") Then pParameterList.Remove("ServiceRestrictionNumber")
            pParameterList.Add("ServiceRestrictionNumber", mvEnv.GetControlNumber("SE").ToString)
          Case "sources"
            If pParameterList.ContainsKey("SourceNumber") Then pParameterList.Remove("SourceNumber")
            pParameterList.Add("SourceNumber", mvEnv.GetControlNumber("SR").ToString)
          Case "surveys"
            If pParameterList.ContainsKey("SurveyNumber") Then pParameterList.Remove("SurveyNumber")
            pParameterList.Add("SurveyNumber", mvEnv.GetControlNumber("SU").ToString)
          Case "survey_versions"
            If pParameterList.ContainsKey("SurveyVersionNumber") Then pParameterList.Remove("SurveyVersionNumber")
            pParameterList.Add("SurveyVersionNumber", mvEnv.GetControlNumber("SV").ToString)
          Case "survey_questions"
            If pParameterList.ContainsKey("SurveyQuestionNumber") Then pParameterList.Remove("SurveyQuestionNumber")
            pParameterList.Add("SurveyQuestionNumber", mvEnv.GetControlNumber("SQ").ToString)
          Case "survey_answers"
            If pParameterList.ContainsKey("SurveyAnswerNumber") Then pParameterList.Remove("SurveyAnswerNumber")
            pParameterList.Add("SurveyAnswerNumber", mvEnv.GetControlNumber("SA").ToString)
          Case "rate_modifiers"
            If pParameterList.ContainsKey("RateModifierNumber") Then pParameterList.Remove("RateModifierNumber")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifiersSequence) AndAlso pParameterList.HasValue("SequenceNumber") Then
              Dim vWhereFields As New CDBFields
              vWhereFields.Add("product", pParameterList("Product").Value)
              vWhereFields.Add("rate", pParameterList("Rate").Value)
              vWhereFields.Add("sequence_number", pParameterList("SequenceNumber").IntegerValue)
              If mvEnv.Connection.GetCount("rate_modifiers", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "Product, Rate and Sequence Number")
            End If

            pParameterList.Add("RateModifierNumber", mvEnv.GetControlNumber("RM").ToString)
          Case "merchant_details"
            'J397: Merchant Retail Number and Merchant ID both are primary keys but should not be treated as combined primary key.
            If pParameterList.HasValue("MerchantRetailNumber") AndAlso pParameterList.HasValue("MerchantId") Then
              Dim vWhereFields As New CDBFields
              vWhereFields.Add("merchant_retail_number", pParameterList("MerchantRetailNumber").Value)
              vWhereFields.Add("merchant_id", pParameterList("MerchantId").Value, CDBField.FieldWhereOperators.fwoOR)
              If mvEnv.Connection.GetCount("merchant_details", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "MerchantRetailNumber and/or MerchantId")
            End If
          Case "web_documents"
            If pParameterList.ContainsKey("WebDocumentNumber") Then pParameterList.Remove("WebDocumentNumber")
            pParameterList.Add("WebDocumentNumber", mvEnv.GetControlNumber("WD").ToString)
          Case "exam_accreditation_statuses"
            If pParameterList.HasValue("AccreditationStatus") Then
              Dim vWhereFields As New CDBFields
              vWhereFields.Add("accreditation_status", pParameterList("AccreditationStatus").Value)
              If mvEnv.Connection.GetCount("exam_accreditation_statuses", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "exam_accreditation_statuses")
            End If
            If pParameterList.ContainsKey("AccreditationStatusId") Then pParameterList.Remove("AccreditationStatusId")
            pParameterList.Add("AccreditationStatusId", mvEnv.GetControlNumber("ACS").ToString)
          Case "workstream_groups"
            'Check if Workstream Group value is not used for event_groups
            Dim vEventWhereFields As New CDBFields
            vEventWhereFields.Add("event_group", pParameterList("WorkstreamGroup").Value)
            If mvEnv.Connection.GetCount("event_groups", vEventWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Event Group")

            'Check if Workstream Group value is not used for organisation_groups
            Dim vOrganisationWhereFields As New CDBFields
            vOrganisationWhereFields.Add("organisation_group", pParameterList("WorkstreamGroup").Value)
            If mvEnv.Connection.GetCount("organisation_groups", vOrganisationWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Organisation Group")

            'Check if Workstream Group value is not used for contact_groups
            Dim vContactWhereFields As New CDBFields
            vContactWhereFields.Add("contact_group", pParameterList("WorkstreamGroup").Value)
            If mvEnv.Connection.GetCount("contact_groups", vContactWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Contact Group")
          Case "event_groups"
            'Check if Event Group value is not used for workstream_groups
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("workstream_group", pParameterList("EventGroup").Value)
            If mvEnv.Connection.GetCount("workstream_groups", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Workstream Group")

            'Check if Event Group value is not used for organisation_groups
            Dim vOrganisationWhereFields As New CDBFields
            vOrganisationWhereFields.Add("organisation_group", pParameterList("EventGroup").Value)
            If mvEnv.Connection.GetCount("organisation_groups", vOrganisationWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Organisation Group")

            'Check if Event Group value is not used for contact_groups
            Dim vContactWhereFields As New CDBFields
            vContactWhereFields.Add("contact_group", pParameterList("EventGroup").Value)
            If mvEnv.Connection.GetCount("contact_groups", vContactWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Contact Group")
          Case "organisation_groups"
            'Check if Organisation Group value is not used for workstream_groups
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("workstream_group", pParameterList("OrganisationGroup").Value)
            If mvEnv.Connection.GetCount("workstream_groups", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Workstream Group")

            'Check if Organisation Group value is not used for event_groups
            Dim vEventWhereFields As New CDBFields
            vEventWhereFields.Add("event_group", pParameterList("OrganisationGroup").Value)
            If mvEnv.Connection.GetCount("event_groups", vEventWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Event Group")

            'Check if Organisaiton Group value is not used for contact_groups
            Dim vContactWhereFields As New CDBFields
            vContactWhereFields.Add("contact_group", pParameterList("OrganisationGroup").Value)
            If mvEnv.Connection.GetCount("contact_groups", vContactWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Contact Group")
          Case "contact_groups"
            'Check if Contact Group value is not used for workstream_groups
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("workstream_group", pParameterList("ContactGroup").Value)
            If mvEnv.Connection.GetCount("workstream_groups", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Workstream Group")

            'Check if Contact Group value is not used for event_groups
            Dim vEventWhereFields As New CDBFields
            vEventWhereFields.Add("event_group", pParameterList("ContactGroup").Value)
            If mvEnv.Connection.GetCount("event_groups", vEventWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Event Group")

            'Check if Contact Group value is not used for organisation_groups
            Dim vOrganisationWhereFields As New CDBFields
            vOrganisationWhereFields.Add("organisation_group", pParameterList("ContactGroup").Value)
            If mvEnv.Connection.GetCount("organisation_groups", vOrganisationWhereFields) > 0 Then RaiseError(DataAccessErrors.daeValueAlreadyUsed, "Organisation Group")
          Case "fp_controls"
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("fp_application", pParameterList("FpApplication").Value)
            vWhereFields.Add("fp_page_type", pParameterList("FpPageType").Value)
            vWhereFields.Add("sequence_number", pParameterList("SequenceNumber").Value)
            If mvEnv.Connection.GetCount("fp_controls", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "Fp Application, Fp Page Type and Sequence Number")
        End Select
      ElseIf pType = MaintenanceTypes.Update Then
        Select Case pParameterList("MaintenanceTableName").Value
          Case "rate_modifiers"
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbRateModifiersSequence) AndAlso pParameterList.HasValue("SequenceNumber") Then
              Dim vWhereFields As New CDBFields
              vWhereFields.Add("product", pParameterList("Product").Value)
              vWhereFields.Add("rate", pParameterList("Rate").Value)
              vWhereFields.Add("sequence_number", pParameterList("SequenceNumber").IntegerValue)
              vWhereFields.Add("rate_modifier_number", pParameterList("RateModifierNumber").IntegerValue, CDBField.FieldWhereOperators.fwoNotEqual)
              If mvEnv.Connection.GetCount("rate_modifiers", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeRecordExists, "Product, Rate and Sequence Number")
            End If
        End Select
      End If
    End Sub

    Public Overrides Function GetUniqueKeyFields(ByVal pParams As CDBParameters) As Data.CDBFields
      Dim vFields As CDBFields = Nothing
      If NoUniqueKey Then
        vFields = New CDBFields
        CheckClassFields()
        'The table does not have a primary key defined. Use all the parameters to initialise the record
        Dim vParamName As String = String.Empty
        For Each vClassField As ClassField In mvClassFields
          If vClassField.Name <> "amended_by" AndAlso vClassField.Name <> "amended_on" _
           AndAlso vClassField.Name <> "created_by" AndAlso vClassField.Name <> "created_on" Then
            If mvMode = MaintenanceTypes.Update Then
              vParamName = "Old" & vClassField.ProperName
            Else
              vParamName = vClassField.ProperName
            End If
            Dim vNewField As CDBField
            If pParams.ContainsKey(vParamName) Then
              vNewField = New CDBField(vClassField.Name, vClassField.FieldType, pParams(vParamName).Value)
            Else
              vNewField = New CDBField(vClassField.Name, vClassField.FieldType, "")
            End If
            If vNewField.FieldType = CDBField.FieldTypes.cftMemo AndAlso pParams(vParamName).Value.Length > 0 Then vNewField.WhereOperator = CDBField.FieldWhereOperators.fwoLike
            vNewField.SpecialColumn = vClassField.SpecialColumn
            vFields.Add(vNewField)
          End If
        Next
      Else
        vFields = MyBase.GetUniqueKeyFields(pParams)
      End If
      Return vFields
    End Function

    Public Overrides Sub Update(ByVal pParameterList As CDBParameters)
      If NoUniqueKey Then
        'Check that the values we are trying to update to dont exist in the db
        'Dim vFields As New CDBFields
        'For Each vClassField As ClassField In mvClassFields
        '  If vClassField.Name <> "amended_by" AndAlso vClassField.Name <> "amended_on" AndAlso vClassField.FieldType <> CDBField.FieldTypes.cftMemo Then
        '    Dim vNewField As CDBField
        '    If pParameterList.ContainsKey(vClassField.ProperName) Then
        '      vNewField = New CDBField(vClassField.Name, vClassField.FieldType, pParameterList(vClassField.ProperName).Value)
        '    Else
        '      vNewField = New CDBField(vClassField.Name, vClassField.FieldType, "")
        '    End If
        '    vNewField.SpecialColumn = vClassField.SpecialColumn
        '    vFields.Add(vNewField)
        '  End If
        'Next
        'If mvEnv.Connection.GetCount(DatabaseTableName, vFields) > 0 Then RaiseError(DataAccessErrors.daeDuplicateRecord)
      End If
      MyBase.Update(pParameterList)
      If NoUniqueKey Then
        Dim vCheckExists As Boolean
        For Each vClassField As ClassField In mvClassFields
          If vClassField.UniqueField AndAlso vClassField.ValueChanged Then
            vCheckExists = True
          End If
        Next
        If vCheckExists Then mvClassFields.CheckRecordExists(mvEnv)
      End If
    End Sub

    Public Overrides Function KeyValueRequired(ByVal pField As String) As Boolean
      Dim vRequired As Boolean = True
      If DatabaseTableName = "rate_nominal_accounts" Then
        'nominal_account_suffix Attribute allows nulls and is part of the primary key so set to 'N'
        If pField = "nominal_account_suffix" Then vRequired = False
      ElseIf DatabaseTableName = "config" Then
        'config table can have null values
        If pField = "department" OrElse pField = "logname" OrElse pField = "client" Then vRequired = False
      ElseIf DatabaseTableName = "explorer_link_access_levels" Then
        'explorer link access level table can have null values
        If pField = "department" OrElse pField = "logname" OrElse pField = "client" Then vRequired = False
      ElseIf DatabaseTableName = "membership_type_categories" Then
        'activity value can have null values as restriction can either be done by activity or combination of activity and value
        If pField = "activity_value" Then vRequired = False
      Else
        vRequired = MyBase.KeyValueRequired(pField)
      End If
      Return vRequired
    End Function

    Protected Overrides Sub AddAdditionalFields()
      MyBase.AddAdditionalFields()
      Select Case DatabaseTableName
        Case "legacy_income_stages"
          If mvEnv.GetConfig("opt_lg_income_stage_level") = "TYPE" Then
            mvClassFields("bequest_type").PrimaryKey = True
          Else
            mvClassFields("bequest_sub_type").PrimaryKey = True
          End If
          mvClassFields("stage_months_delay").PrimaryKey = True
        Case "fp_applications"
          mvClassFields("fp_application").PrimaryKey = True
      End Select
      mvClassFields.TableMaintenance = mvTableMaintenance

      Dim vBulkFields As String() = {"document", "standard_document_text", "document_text"}
      'Ignore these attributes as they are bulk values
      For Each vBulkField As String In vBulkFields
        If mvClassFields.ContainsKey(vBulkField) Then mvClassFields.Remove(vBulkField)
      Next
      For Each vClassField As ClassField In mvClassFields
        If mvEnv.Connection.IsSpecialColumn(vClassField.Name) Then vClassField.SpecialColumn = True
      Next

    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Select Case mvMode
        Case MaintenanceTypes.Insert
          Select Case DatabaseTableName
            Case "relationships"
              If Not Me.Field("complimentary_relationship").HasValue Then
                Me.Field("auto_create_complementary").Bool = False
              End If
          End Select
        Case MaintenanceTypes.Update
          Select Case DatabaseTableName
            Case "packages" 'BR16591 - Do not allow ackages.storage_path to be changed, if documents exits for the package. 
              If mvClassFields("storage_type").ValueChanged Then
                Dim vDocWhereFields As CDBFields = New CDBFields() ' Check Communications_log
                vDocWhereFields.Add("package", CDBField.FieldTypes.cftCharacter, mvClassFields("package").Value, CDBField.FieldWhereOperators.fwoEqual)
                Dim vCountCommsLog As Integer = mvEnv.Connection.GetCount("communications_log", vDocWhereFields)
                If vCountCommsLog = 0 Then
                  Dim vStdDocWhereFields As CDBFields = New CDBFields() ' Check Standard Documents, if nothing found in communications_log
                  vStdDocWhereFields.Add("package", CDBField.FieldTypes.cftCharacter, mvClassFields("package").Value, CDBField.FieldWhereOperators.fwoEqual)
                  Dim vCountStdDocs As Integer = mvEnv.Connection.GetCount("standard_documents", vStdDocWhereFields)
                  If vCountStdDocs > 0 Then
                    RaiseError(DataAccessErrors.daePackageStorageTypeCannotChange)
                  End If
                Else
                  RaiseError(DataAccessErrors.daePackageStorageTypeCannotChange)
                End If
              Else
              End If
            Case "activities"
              If mvClassFields("is_historic").ValueChanged AndAlso mvClassFields("is_historic").Value = "Y" Then
                ValidateActivityDependencies()
              End If
            Case "activity_values"
              If mvClassFields("is_historic").ValueChanged AndAlso mvClassFields("is_historic").Value = "Y" Then
                ValidateActivityValueDependencies()
              End If
            Case "mailing_suppressions"
              If mvClassFields("is_historic").ValueChanged AndAlso mvClassFields("is_historic").Value = "Y" Then
                ValidateSuppressionDependencies()
              End If
            Case "relationships"
              If Me.Field("complimentary_relationship").IsNullOrWhitespace Then
                Me.Field("auto_create_complementary").Bool = False
              End If
          End Select
      End Select

      If mvClassFields.DatabaseTableName = "config_names" AndAlso mvClassFields.ContainsKey("amended_by") Then
        Me.mvOverrideAmended = True
        Me.mvOverrideCreated = True
        MyBase.Save(String.Empty, pAudit, pJournalNumber)
      Else
        MyBase.Save(If(mvEnv.InitialisingDatabase AndAlso String.IsNullOrWhiteSpace(pAmendedBy), "dbinit", pAmendedBy), pAudit, pJournalNumber)
      End If

      'Check for any additional items to be performed (topics, activities, printers)
      Select Case mvMode
        Case MaintenanceTypes.Insert
          Select Case DatabaseTableName
            Case "config"
              CheckConfigChanges()
            Case "products", "product_warehouses"
              CheckStockMovement()                   'Check stock movement required
              CheckProductWarehouses()               'Check if product_warehouse needs creating
              CheckProductCosts()                    'Check if ProductCosts need creating
            Case "users"
              Dim vDOD As New DepartmentOwnershipDefault(mvEnv)
              vDOD.AddForUser(mvClassFields("department").Value, mvClassFields("logname").Value)
            Case "topics", "activities"
              Dim vCDBFields As New CDBFields
              If DatabaseTableName = "activities" Then
                'Add activity users record
                vCDBFields.Clear()
                vCDBFields.AddAmendedOnBy(mvEnv.User.Logname)
                vCDBFields.Add("department", mvEnv.User.Department)
                vCDBFields.Add("activity", mvClassFields("activity").Value)
                mvEnv.Connection.InsertRecord("activity_users", vCDBFields)
              End If
            Case "mailing_template_documents"
              Dim vCDBFields As New CDBFields
              vCDBFields.AddAmendedOnBy(mvEnv.User.Logname)
              vCDBFields.Add("explicit_selection", CDBField.FieldTypes.cftCharacter, "Y")

              Dim vWhereFields As CDBFields = New CDBFields
              With vWhereFields
                .Add("mailing_template", CDBField.FieldTypes.cftCharacter, mvClassFields("mailing_template").Value)
                '.Add "explicit_selection", cftCharacter, "N" ab 8/8/05 user may have manually set to 'Y'
              End With
              mvEnv.Connection.UpdateRecords("mailing_templates", vCDBFields, vWhereFields)
            Case "ownership_groups"
              Dim vOGU As New OwnershipGroupUser(mvEnv)
              Dim vRS As CDBRecordSet
              Dim vOALT As CDBEnvironment.OwnershipAccessLevelTypes
              'find all users and add browse access for the new ownership group
              vRS = mvEnv.Connection.GetRecordSet("SELECT logname, department FROM users")
              While vRS.Fetch
                vOGU.Init()
                If vRS.Fields("department").Value = mvClassFields("principal_department").Value Then
                  vOALT = CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite
                Else
                  vOALT = CDBEnvironment.OwnershipAccessLevelTypes.oaltBrowse
                End If
                vOGU.InitFromDepartment(mvClassFields("ownership_group").Value, vRS.Fields("logname").Value, vOALT)
                vOGU.Save()
              End While
              vRS.CloseRecordSet()
            Case "ownership_group_users"
              'find all old records for the user being added if there are any
              Dim vValidTo As New DateTime
              vValidTo = CDate(mvClassFields("valid_from").Value).AddDays(-1)
              Dim vWhereFields As CDBFields = New CDBFields
              With vWhereFields
                .Add("ownership_group", CDBField.FieldTypes.cftCharacter, mvClassFields("ownership_group").Value)
                .Add("logname", CDBField.FieldTypes.cftCharacter, mvClassFields("logname").Value)
                .Add("valid_from", CDBField.FieldTypes.cftDate, mvClassFields("valid_from").Value, CDBField.FieldWhereOperators.fwoNullOrLessThan)
                .Add("valid_to", CDBField.FieldTypes.cftDate, mvClassFields("valid_from").Value, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
              End With
              Dim vSql As New CARE.Data.SQLStatement(mvEnv.Connection, "ownership_group, logname, ownership_access_level,valid_from, valid_to, amended_by, amended_on", "ownership_group_users", vWhereFields)
              Dim vOGU As New OwnershipGroupUser(mvEnv)
              Dim vRS As CDBRecordSet
              vRS = mvEnv.Connection.GetRecordSet(vSql.SQL.ToString)
              'update the valid to date for any existing records to finish when the new one starts
              While vRS.Fetch
                vOGU.InitFromRecordSet(vRS)
                vOGU.ValidTo = vValidTo.ToString
                vOGU.Save()
              End While
              vRS.CloseRecordSet()
            Case "rate_modifiers"
              Dim vCDBFields As New CDBFields
              vCDBFields.Add("use_modifiers", CDBField.FieldTypes.cftCharacter, "Y")
              Dim vWhereFields As CDBFields = New CDBFields
              With vWhereFields
                .Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields("product").Value)
                .Add("rate", CDBField.FieldTypes.cftCharacter, mvClassFields("rate").Value)
              End With
              mvEnv.Connection.UpdateRecords("rates", vCDBFields, vWhereFields)
          End Select

          'TODO:Currently does not support adding of an activityvalue while adding an activity
          'If DatabaseTableName = "activities" OrElse DatabaseTableName = "activity_values" Then
          If DatabaseTableName = "activity_values" Then
            'Add activity value users record
            Dim vCDBFields As New CDBFields
            vCDBFields.AddAmendedOnBy(mvEnv.User.Logname)
            vCDBFields.Add("department", CDBField.FieldTypes.cftCharacter, mvEnv.User.Department)
            vCDBFields.Add("activity", CDBField.FieldTypes.cftCharacter, mvClassFields("activity").Value)
            vCDBFields.Add("activity_value", CDBField.FieldTypes.cftCharacter, mvClassFields("activity_value").Value)
            mvEnv.Connection.InsertRecord("activity_value_users", vCDBFields)
          End If
        Case MaintenanceTypes.Update
          If mvStockFlagAfter Then
            CheckStockMovement()                   'Check stock movement required
            CheckProductWarehouses()               'Check if product_warehouse needs creating
            CheckProductCosts()                    'Check if ProductCosts need creating
          End If

          Select Case DatabaseTableName
            Case "membership_controls", "membership_types"
              UpdateMembershipGroups()

            Case "config"
              CheckConfigChanges()

            Case "relationships"
              'Check if the complementary relationship is set.  If it is then make sure that we don't create a daisy-chain of complementaries by making the two relationships reflexive
              Dim vWhere As New CDBFields()
              Dim vCompField As ClassField = Me.ClassFields("complimentary_relationship")
              If vCompField.HasValue Then
                vWhere.Add("relationship", vCompField.Value)
                Dim vCompRelationship As Relationship = CARERecordFactory.SelectInstance(Of Relationship)(Me.Environment, vWhere)
                If vCompRelationship IsNot Nothing Then
                  vCompRelationship.ComplementaryRelationship = Me.Field("relationship").Value
                  vCompRelationship.Save()
                End If
              End If

          End Select
      End Select
    End Sub

    Private Sub CheckConfigChanges()
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vEvent As New CDBEvent(mvEnv)
      Dim vRS As CDBRecordSet

      Select Case mvClassFields("config_name").Value
        Case "ev_uppercase_references"
          If mvClassFields("config_value").Value.Substring(0, 1).ToUpper = "Y" AndAlso Not mvMode = MaintenanceTypes.Update Then
            'Switch all event references to upper case
            vEvent.Init()
            vWhereFields.Add("event_reference", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoNotEqual)
            vRS = New SQLStatement(mvEnv.Connection, "ev.event_number,event_reference", "events ev", vWhereFields).GetRecordSet
            With vEvent
              While vRS.Fetch
                .InitFromRecordSet(vRS)
                .UpperCaseEventReference()
                .Save()
              End While
            End With
            vRS.CloseRecordSet()
            vWhereFields.Clear()
            vWhereFields.Add("table_name", "events")
            vWhereFields.Add("attribute_name", "event_reference")
            vUpdateFields.Add("case", "U").SpecialColumn = True
            mvEnv.Connection.UpdateRecords("maintenance_attributes", vUpdateFields, vWhereFields)
          End If
        Case "fixed_renewal_M"
          DBSetup.CreateFixedRenewalLookup(mvEnv, mvClassFields("config_value").Value)
      End Select
    End Sub

    Private Sub UpdateMembershipGroups()
      'Update MembershipGroup data - MembershipControls & MembershipTypes tables
      Dim vMember As New Member
      Dim vUseMembershipGroups As Boolean

      vUseMembershipGroups = False
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) Then
        If DatabaseTableName = "membership_types" Then
          vUseMembershipGroups = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMemberOrganisationGroup).Length > 0
        ElseIf DatabaseTableName = "membership_controls" Then
          vUseMembershipGroups = True
        End If
      End If

      If vUseMembershipGroups Then
        If mvSetMembershipGroups Then
          If mvHistoric Then
            vMember.SetMembershipGroupsHistoric(mvEnv, 0, mvMemberTypeCode)
          Else
            vMember.SetMembershipGroups(mvEnv, 0, mvMemberTypeCode, mvOrgGroup)
          End If
        End If
      End If
    End Sub
    ''' <summary>
    ''' CARERecord PreValidateUpdateParameters. Called before Update is performed
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks>Add a case statement and call a subroutine to do the validation, don't bloat this with validation code.</remarks>
    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      PreValidateSurveyUpdateParameters(pParameterList)
    End Sub
    ''' <summary>
    ''' CARERecord PreValidateCreateParameters. Called before Insert is performed
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks>Add a case statement and call a subroutine to do the validation, don't bloat this with validation code.</remarks>
    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      PreValidateSurveyCreateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      Select Case pParameterList("MaintenanceTableName").Value
        Case "organisation_groups"
          PostValidateUpdateOrganisationGroupParameters(pParameterList)
      End Select
    End Sub

    Private Sub CheckProductCosts()
      'If we have just added a new stock Product or ProductWarehouse then add a new ProductCosts record
      Dim vRS As CDBRecordSet
      Dim vFields As New CDBFields
      Dim vAdd As Boolean
      Dim vCostOfSale As Double

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductCosts) Then
        If DatabaseTableName = "products" Then
          If mvClassFields("stock_item").Bool Then
            'Only add for a stock product if we have Warehouse & BinNumber fields (without these ProductWarehouse will not have been created)
            If mvClassFields("warehouse").Value.Length > 0 AndAlso mvClassFields("bin_number").Value.Length > 0 Then vAdd = True
            vCostOfSale = mvClassFields("cost_of_sale").DoubleValue
          End If
        Else
          'ProductWarehouse
          vAdd = True
          vRS = mvEnv.Connection.GetRecordSet("SELECT cost_of_sale FROM products WHERE product = '" & mvClassFields("product").Value & "'")
          If vRS.Fetch Then vCostOfSale = vRS.Fields("cost_of_sale").DoubleValue
          vRS.CloseRecordSet()
        End If
        If vAdd Then
          With vFields
            .Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields("product").Value)
            .Add("warehouse", CDBField.FieldTypes.cftCharacter, mvClassFields("warehouse").Value)
          End With
          If mvEnv.Connection.GetCount("product_costs", vFields) = 0 Then
            With vFields
              .Add("product_cost_number", CDBField.FieldTypes.cftLong, mvEnv.GetControlNumber("PC"))
              .Add("cost_of_sale", CDBField.FieldTypes.cftNumeric, vCostOfSale)
              .Add("original_quantity", CDBField.FieldTypes.cftLong, mvClassFields("last_stock_count").LongValue)
              .Add("last_stock_count", CDBField.FieldTypes.cftLong, mvClassFields("last_stock_count").LongValue)
              .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
              .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate())
            End With
            mvEnv.Connection.InsertRecord("product_costs", vFields)
          End If
        End If
      End If

    End Sub

    Private Sub CheckProductWarehouses()
      'Add product_warehouses record if using product warehouses
      'If adding product_warehouses then increment stock count on product
      Dim vPUFields As New CDBFields
      Dim vPWFields As New CDBFields

      If DatabaseTableName = "products" Then
        If mvClassFields("stock_item").Bool Then
          If mvClassFields("warehouse").Value.Length > 0 AndAlso mvClassFields("bin_number").Value.Length > 0 Then
            With vPWFields
              .Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields("product").Value)
              .Add("warehouse", CDBField.FieldTypes.cftCharacter, mvClassFields("warehouse").Value)
            End With
            If mvEnv.Connection.GetCount("product_warehouses", vPWFields) = 0 Then
              With vPWFields
                .Add("bin_number", CDBField.FieldTypes.cftCharacter, mvClassFields("bin_number").Value)
                .Add("last_stock_count", CDBField.FieldTypes.cftLong, mvClassFields("last_stock_count").LongValue)
              End With
              vPWFields.AddAmendedOnBy(mvEnv.User.Logname) 'NFPCARE-98: not setting the amended_by, amended_on was causing an exception 
              mvEnv.Connection.InsertRecord("product_warehouses", vPWFields)
            End If
          End If
        End If
      ElseIf DatabaseTableName = "product_warehouses" Then
        If mvClassFields("last_stock_count").LongValue > 0 Then
          vPWFields.Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields("product").Value)
          vPUFields.Add("last_stock_count", CDBField.FieldTypes.cftLong, "last_stock_count + " & mvClassFields("last_stock_count").Value)
          mvEnv.Connection.UpdateRecords("products", vPUFields, vPWFields)
        End If
      End If

    End Sub

    Private Sub CheckStockMovement()
      Dim vStockMovement As New StockMovement
      Dim vStockWhere As New CDBFields
      Dim vInitialReason As String
      Dim vAdd As Boolean

      If DatabaseTableName = "products" Then          'Handle stock movements for products
        vAdd = mvClassFields("stock_item").Bool
      ElseIf DatabaseTableName = "product_warehouses" Then
        vAdd = True                             'New product warehouse created, need a stock movement
      End If

      If vAdd Then
        vInitialReason = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlStockReasonInitial)
        If vInitialReason.Length > 0 Then
          vStockWhere.Add("product", CDBField.FieldTypes.cftCharacter, mvClassFields("product").Value)
          vStockWhere.Add("stock_movement_reason", CDBField.FieldTypes.cftCharacter, vInitialReason)
          If mvClassFields("warehouse").Value.Length > 0 Then vStockWhere.Add("warehouse", CDBField.FieldTypes.cftCharacter, mvClassFields("warehouse").Value)
          If mvEnv.Connection.GetCount("stock_movements", vStockWhere) = 0 Then
            vStockMovement.Create(mvEnv, mvClassFields("product").Value, 0, vInitialReason, 0, 0, 0, True, mvClassFields("warehouse").Value)
          End If
        End If
      End If
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      'First we must check to see if a delete from this table is allowed or if any references exist
      If Not CheckUsedElsewhere("D") Then
        If CanDelete() Then
          'Should always delete a row
          CheckDeleteOther()
          MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)

          If DatabaseTableName = "config" Then
            If mvClassFields("config_name").Value = "fixed_renewal_M" Then DBSetup.CreateFixedRenewalLookup(mvEnv, "")
          End If

        End If
      End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pContext">pContext may take values of 'D' (deletion) or 'C' (change). An appropriate
    ''' message will be displayed in each case.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function CheckUsedElsewhere(ByVal pContext As String) As Boolean
      Dim vUsed As Boolean
      Dim vRecordSet As CDBRecordSet

      Select Case DatabaseTableName
        Case "branch_postcodes"
          vUsed = CheckBranchPCUsed()
        Case "rate_nominal_accounts"
          vUsed = CheckRateNominalAccounts()
        Case Else
          Dim vFieldList As String = "primary_attribute_name, related_table_name, related_attribute_name, related_table_desc, related_attribute_desc"
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("primary_table_name", DatabaseTableName)
          vWhereFields.Add("primary_delete_action", "N")
          vRecordSet = New SQLStatement(mvEnv.Connection, vFieldList, "maintenance_relations", vWhereFields, "sequence_number").GetRecordSet
          While vRecordSet.Fetch() AndAlso vUsed = False
            Dim vRelatedWhereFields As CDBFields = GetWhereFields(vRecordSet.Fields("primary_attribute_name").Value, vRecordSet.Fields("related_attribute_name").Value)
            If vRelatedWhereFields.Count > 0 Then
              Dim vRelatedTable As String = vRecordSet.Fields("related_table_name").Value
              If vRelatedTable.Length > 30 Then vRelatedTable = vRelatedTable.Substring(0, 30) 'Limit the table name to 30 chars
              Dim vCount As Integer = mvEnv.Connection.GetCount(vRelatedTable, vRelatedWhereFields)
              If vCount > 0 Then
                If pContext = "C" Then
                  'A record exists and therefore we cannot change
                  RaiseError(DataAccessErrors.daeRecordCannotBeChanged, vRecordSet.Fields("related_table_desc").Value, vRecordSet.Fields("related_attribute_desc").Value)    '%s refer to this %s\r\n\r\nRecord cannot be changed
                Else
                  'A record exists and therefore we cannot delete
                  RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, vRecordSet.Fields("related_table_desc").Value, vRecordSet.Fields("related_attribute_desc").Value)    '%s refer to this %s\r\n\r\nRecord cannot be deleted
                End If
                vUsed = True
              End If
            End If
          End While
          vRecordSet.CloseRecordSet()
      End Select
      Return vUsed
    End Function

    Private Function CheckBranchPCUsed() As Boolean
      Dim vUsed As Boolean
      Dim vCount As Long
      Dim vBranch As String
      Dim vOutPostcode As String
      Dim vWhere As String
      Dim vWhereFields As New CDBFields
      Dim vCheckOutPostcode As String
      Dim vCoveredByOtherBranch As Boolean

      vBranch = mvClassFields("branch").Value
      vOutPostcode = mvClassFields("outward_postcode").Value & "*"

      'TA BR 7831: 1st check no other Branch Postcode records would encompass Addresses
      'covered by this Branch, e.g. GU7 1 would be covered by G, GU or GU7
      'Note, 1st time get rid of asterisk as well as last char
      vCheckOutPostcode = vOutPostcode.Substring(0, vOutPostcode.Length - 2).TrimEnd

      While vCheckOutPostcode.Length > 0 AndAlso Not vCoveredByOtherBranch
        vWhereFields.Clear()
        vWhereFields.Add("outward_postcode", CDBField.FieldTypes.cftCharacter, vCheckOutPostcode, CDBField.FieldWhereOperators.fwoEqual)
        If mvEnv.Connection.GetCount("branch_postcodes", vWhereFields) > 0 Then
          vCoveredByOtherBranch = True
        Else
          vCheckOutPostcode = vCheckOutPostcode.Substring(0, vCheckOutPostcode.Length - 1).TrimEnd
        End If
      End While

      vWhereFields.Clear()
      If Not vCoveredByOtherBranch Then
        vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, vBranch)
        vWhereFields.Add("postcode", CDBField.FieldTypes.cftCharacter, vOutPostcode, CDBField.FieldWhereOperators.fwoLike)
        vCount = mvEnv.Connection.GetCount("addresses", vWhereFields)
        If vCount > 0 Then
          'A record exists and therefore we cannot delete
          RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, "Addresses", "Branch Postcode")    '%s refer to this %s\r\n\r\nRecord cannot be deleted
          vUsed = True
        End If
      End If

      If vUsed = False AndAlso vCoveredByOtherBranch = False Then
        vWhere = "bi.branch_code = '" & vBranch & "' AND o.order_number = bi.order_number AND a.address_number = o.address_number AND a.postcode " & mvEnv.Connection.DBLike(vOutPostcode)
        vCount = mvEnv.Connection.GetCount("branch_income bi,orders o,addresses a", Nothing, vWhere)
        If vCount > 0 Then
          'A record exists and therefore we cannot delete
          RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, "Branch Income", "Branch Postcode")    '%s refer to this %s\r\n\r\nRecord cannot be deleted
          vUsed = True
        End If
      End If

      If vUsed = False AndAlso vCoveredByOtherBranch = False Then
        vWhere = "o.branch = '" & vBranch & "' AND a.address_number = o.address_number AND a.postcode " & mvEnv.Connection.DBLike(vOutPostcode)
        vCount = mvEnv.Connection.GetCount("orders o,addresses a", Nothing, vWhere)
        If vCount > 0 Then
          'A record exists and therefore we cannot delete
          RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, "Payment Plans", "Branch Postcode")    '%s refer to this %s\r\n\r\nRecord cannot be deleted
          vUsed = True
        End If
      End If

      If vUsed = False AndAlso vCoveredByOtherBranch = False Then
        vWhere = "m.branch = '" & vBranch & "' AND a.address_number = m.address_number AND a.postcode " & mvEnv.Connection.DBLike(vOutPostcode)
        vCount = mvEnv.Connection.GetCount("members m,addresses a", Nothing, vWhere)
        If vCount > 0 Then
          'A record exists and therefore we cannot delete
          RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, "Members", "Branch Postcode")    '%s refer to this %s\r\n\r\nRecord cannot be deleted
          vUsed = True
        End If
      End If
      Return vUsed
    End Function

    Private Function CheckRateNominalAccounts() As Boolean
      'Check if nominal_suffix used for any rates
      Dim vWhereFields As New CDBFields
      Dim vNominal As String
      Dim vSuffix As String
      Dim vUsed As Boolean

      vNominal = mvClassFields("product_nominal_account").Value
      vSuffix = mvClassFields("nominal_account_suffix").Value

      '1) Check nominal_account
      With vWhereFields
        .Add("p.nominal_account", CDBField.FieldTypes.cftCharacter, vNominal)
        .Add("r.product", CDBField.FieldTypes.cftLong, "p.product")
        .Add("r.nominal_account_suffix", CDBField.FieldTypes.cftCharacter, vSuffix)
      End With

      If mvEnv.Connection.GetCount("products p, rates r", vWhereFields) > 0 Then vUsed = True

      '2) Check subsequent_nominal_account
      If Not vUsed Then
        vWhereFields = New CDBFields
        With vWhereFields
          .Add("p.subsequent_nominal_account", CDBField.FieldTypes.cftCharacter, vNominal)
          .Add("r.product", CDBField.FieldTypes.cftLong, "p.product")
          .Add("r.subsequent_nominal_suffix", CDBField.FieldTypes.cftCharacter, vSuffix)
        End With

        If mvEnv.Connection.GetCount("products p, rates r", vWhereFields) > 0 Then vUsed = True
      End If

      If vUsed Then RaiseError(DataAccessErrors.daeRecordCannotBeDeleted, "Rates", "Rate Nominal Accounts") '%s refer to this %s\r\n\r\nRecord cannot be deleted
      Return vUsed
    End Function

    Private Function GetWhereFields(ByVal pPrimaryAttrs As String, ByVal pRelatedAttrs As String) As CDBFields
      Dim vPrimaryAttrs() As String
      Dim vRelatedAttrs() As String
      Dim vFound As Boolean
      Dim vWhereFields As New CDBFields

      vPrimaryAttrs = pPrimaryAttrs.Split(","c)
      vRelatedAttrs = pRelatedAttrs.Split(","c)
      If vPrimaryAttrs.Length <> vRelatedAttrs.Length Then
        RaiseError(DataAccessErrors.daePrimaryAndRelatedAttributesDontMatch)
      Else
        For vIndex As Integer = 0 To vPrimaryAttrs.Length - 1
          vFound = False
          If mvClassFields.ContainsKey(vPrimaryAttrs(vIndex).Trim) Then
            vFound = True
            vWhereFields.Add(vRelatedAttrs(vIndex).Trim, mvClassFields(vPrimaryAttrs(vIndex).Trim).FieldType, mvClassFields(vPrimaryAttrs(vIndex).Trim).SetValue)
          End If
          If Not vFound Then Exit For
        Next
      End If

      Return vWhereFields
    End Function

    Private Function CanDelete() As Boolean
      Dim vCanDelete As Boolean
      Dim vMT As MembershipType
      Dim vRate As ProductRate
      Dim vFields As CDBFields

      vCanDelete = True
      Select Case DatabaseTableName
        Case "membership_prices"
          vMT = New MembershipType(mvEnv)
          vMT.Init(mvClassFields("membership_type").Value)

          vRate = New ProductRate(mvEnv)
          vRate.Init(vMT.FirstPeriodsProduct, mvClassFields("rate").Value)

          If Not vMT Is Nothing AndAlso Not vRate Is Nothing Then
            If vMT.Existing AndAlso vRate.Existing Then
              vCanDelete = vRate.Concessionary
              If Not vCanDelete Then RaiseError(DataAccessErrors.daePrimaryRateForMembership)
            End If
          End If

        Case "packed_products"
          vFields = New CDBFields
          vFields.Add("product", mvClassFields("product").Value)
          vFields.Add("rate", mvClassFields("rate").Value)

          If vFields("product").Value.Length > 0 And vFields("rate").Value.Length > 0 Then
            'Once Pack has been sold, can not delete the details
            vCanDelete = (mvEnv.Connection.GetCount("financial_history_details", vFields) = 0)
            If vCanDelete Then vCanDelete = (mvEnv.Connection.GetCount("batch_transaction_analysis", vFields) = 0)
            If Not vCanDelete Then RaiseError(DataAccessErrors.daePackCannotBeDeleted)
          End If

      End Select
      Return vCanDelete
    End Function

    Public Sub CheckDeleteOther()
      'Check for Unique relations (If record in primary unique then ask delete in secondary)
      'Check for cascade deletes - If record in primary delete values in secondary
      Dim vRelatedTable As String
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vWhereFields As CDBFields

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT primary_attribute_name, related_table_name, related_attribute_name, related_table_desc, related_attribute_desc, primary_delete_action FROM maintenance_relations WHERE primary_table_name = '" & DatabaseTableName & "' AND primary_delete_action in ('U','C') ORDER BY primary_delete_action")
      While vRecordSet.Fetch
        vRelatedTable = vRecordSet.Fields("related_table_name").Value
        vWhereFields = GetWhereFields(vRecordSet.Fields("primary_attribute_name").Value, vRecordSet.Fields("related_attribute_name").Value)
        If vWhereFields.Count > 0 Then
          If vRecordSet.Fields("primary_delete_action").Value = "C" Then
            '(C) Cascade delete
            mvEnv.Connection.DeleteRecords(vRelatedTable, vWhereFields, False)
          Else
            '(U) Delete parent if unique entry in primary table
            'Because of using the BuildRelatedWhere function above
            'this will only work if the attributes in the parent and child have the same names!!!!!
            If mvEnv.Connection.GetCount(DatabaseTableName, vWhereFields) = 1 Then
              If Not mvConfirmDelete Then RaiseError(DataAccessErrors.daeDeleteParentIfUniqueEntry, vRecordSet.Fields("related_table_desc").Value, vRecordSet.Fields("related_attribute_desc").Value, vRecordSet.Fields("related_attribute_desc").Value) 'This is the only %s for the %s\r\n\r\nDo you want to delete the %s as well?
              'OK here we have a problem - Since we are going to delete a record from the parent table
              'we should really get recursive and check for values used elsewhere, check it's parent
              'and check for cascade delete. However this is not feasible since we are reading
              'attribute values from the grid on the maintenance form
              'We can however cascade delete if required using the same where clause
              'This will work for the current case of activity_users where activity = 'value'
              'So after the main delete let's do a one time check on the relations table and delete if required
              mvEnv.Connection.DeleteRecords(vRelatedTable, vWhereFields, False)
              vRecordSet2 = mvEnv.Connection.GetRecordSet("SELECT related_table_name FROM maintenance_relations WHERE primary_table_name = '" & vRelatedTable & "' AND primary_delete_action = 'C'")
              While vRecordSet2.Fetch()
                vRelatedTable = vRecordSet2.Fields(1).Value
                mvEnv.Connection.DeleteRecords(vRelatedTable, vWhereFields, False)
              End While
              vRecordSet2.CloseRecordSet()
            End If
          End If
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Function LoadActivityReferenceTables() As XElement

      Dim vDataStore As XElement = <ActivityDependentData>
                                     <!--Tables with no Activity Value column -->
                                     <ActivityDependentTable Name="activity_group_details" Description="Activity Group Details" ActivityColumn="activity"/>
                                     <ActivityDependentTable Name="contact_groups" Description="Contact Groups (Graph Activity)" ActivityColumn="graph_activity"/>
                                     <ActivityDependentTable Name="organisation_groups" Description="Organisation Groups (Graph Activity)" ActivityColumn="graph_activity"/>
                                     <!--Tables with standard Activity and value column-->
                                     <ActivityDependentTable Name="covenant_controls" Description="Covenant Controls" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="exam_prices" Description="" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="legacy_controls" Description="Legacy Controls" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="legacy_statuses" Description="Legacy Statuses" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="membership_prices" Description="Membership Prices" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="membership_type_categories" Description="Membership Type Categories" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="membership_types" Description="Membership Types" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="product_offers" Description="Product Offers" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="products" Description="Products" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="purchase_order_activities" Description="Purchase Order Activities" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="qp_answers" Description="QP Answers" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="rate_modifiers" Description="Rate Modifiers" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="session_activities" Description="Session Activities" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="sub_topics" Description="Sub Topics" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <ActivityDependentTable Name="tick_boxes" Description="Tick Boxes" ActivityColumn="activity" ActivityValueColumn="activity_value"/>
                                     <!--Tables with non-standard Activity and value column-->
                                     <ActivityDependentTable Name="contact_controls" Description="Contact Controls (Current Days Remaining Activity)" ActivityColumn="curr_days_remaining_activity" ActivityValueColumn="curr_days_remaining_act_val"/>
                                     <ActivityDependentTable Name="contact_controls" Description="Contact Controls (Permitted Days Activity)" ActivityColumn="max_permitted_days_activity" ActivityValueColumn="max_permitted_days_act_val"/>
                                     <ActivityDependentTable Name="contact_controls" Description="Contact Controls (Qualifying Position Activity" ActivityColumn="qualifying_position_activity" ActivityValueColumn="qualifying_position_act_val"/>
                                     <ActivityDependentTable Name="events" Description="Events (Deferred Booking Activity)" ActivityColumn="deferred_booking_act" ActivityValueColumn="deferred_booking_act_value"/>
                                     <ActivityDependentTable Name="events" Description="Events (Rejected Booking Activity)" ActivityColumn="rejected_booking_act" ActivityValueColumn="rejected_booking_act_value"/>
                                     <ActivityDependentTable Name="exam_controls" Description="Exam Controls" ActivityColumn="exemption_org_activity" ActivityValueColumn="exemption_org_activity_value"/>
                                     <ActivityDependentTable Name="financial_controls" Description="Financial Controls (CCCA Activity)" ActivityColumn="ccca_activity" ActivityValueColumn="ccca_activity_value"/>
                                     <ActivityDependentTable Name="financial_controls" Description="Financial Controls (Direct Debit Activity)" ActivityColumn="direct_debit_activity" ActivityValueColumn="direct_debit_activity_value"/>
                                     <ActivityDependentTable Name="financial_controls" Description="Financial Controls (Standing Order Activity)" ActivityColumn="standing_order_activity" ActivityValueColumn="standing_order_activity_value"/>
                                     <ActivityDependentTable Name="gaye_controls" Description="GAYE Controls (GAYE Pledge Activity)" ActivityColumn="gaye_pledge_activity" ActivityValueColumn="gaye_pledge_activity_value"/>
                                     <ActivityDependentTable Name="gaye_controls" Description="GAYE Controls (Post Tax Employer Activity)" ActivityColumn="post_tax_employer_activity" ActivityValueColumn="post_tax_employer_act_value"/>
                                     <ActivityDependentTable Name="gaye_controls" Description="GAYE Controls (Post Tax Pledge Activity)" ActivityColumn="post_tax_pledge_activity" ActivityValueColumn="post_tax_pledge_activity_value"/>
                                     <ActivityDependentTable Name="marketing_controls" Description="Marketing Controls (Bank Order Activity)" ActivityColumn="b_order_activity" ActivityValueColumn="b_order_activity_value"/>
                                     <ActivityDependentTable Name="marketing_controls" Description="Marketing Controls (Covenant Activity)" ActivityColumn="covenant_activity" ActivityValueColumn="covenant_activity_value"/>
                                     <ActivityDependentTable Name="membership_controls" Description="Membership Controls (Sponsor Activity)" ActivityColumn="sponsor_activity" ActivityValueColumn="sponsor_activity_value"/>
                                     <ActivityDependentTable Name="service_controls" Description="Service Controls (Modifier Activity)" ActivityColumn="modifier_activity" ActivityValueColumn="modifier_activity_value"/>
                                     <ActivityDependentTable Name="surveys" Description="Responded Activity" ActivityColumn="responded_activity" ActivityValueColumn="responded_activity_value"/>
                                     <ActivityDependentTable Name="surveys" Description="Surveys (Sent Activity)" ActivityColumn="sent_activity" ActivityValueColumn="sent_activity_value"/>
                                   </ActivityDependentData>
      Return vDataStore

    End Function
    Private Function LoadSuppressionReferenceTables() As XElement

      Dim vDataStore As XElement = <SuppressionDependentData>
                                     <!--Tables with no Activity Value column -->
                                     <SuppressionDependentTable Name="contact_controls" Description="Contact Controls (Default Suppression)" SuppressionColumn="default_mailing_suppression"/>
                                     <SuppressionDependentTable Name="contact_controls" Description="Contact Controls (Gone Away Suppression)" SuppressionColumn="gone_away_mailing_suppression"/>
                                     <SuppressionDependentTable Name="fp_applications" Description="FP Applications" SuppressionColumn="mailing_suppression"/>
                                     <SuppressionDependentTable Name="marketing_controls" Description="Marketing Controls (Data Protection Suppression)" SuppressionColumn="data_prot_mailing_supp"/>
                                     <SuppressionDependentTable Name="marketing_controls" Description="Marketing Controls (Derived Contact Suppression)" SuppressionColumn="derived_contact_mailing_supp"/>
                                     <SuppressionDependentTable Name="marketing_controls" Description="Marketing Controls (Gone Away Suppression)" SuppressionColumn="gone_away_mailing_supp"/>
                                     <SuppressionDependentTable Name="marketing_controls" Description="Marketing Controls (Joint Contact Suppression)" SuppressionColumn="joint_contact_mailing_supp"/>
                                     <SuppressionDependentTable Name="membership_types" Description="Membership Types" SuppressionColumn="mailing_suppression"/>
                                     <SuppressionDependentTable Name="suppression_group_details" Description="Suppression Group Details" SuppressionColumn="mailing_suppression"/>
                                     <SuppressionDependentTable Name="tick_boxes" Description="Tick Boxes" SuppressionColumn="mailing_suppression"/>
                                     <SuppressionDependentTable Name="unsubscribe_suppressions" Description="Unsubscribe Suppressions" SuppressionColumn="mailing_suppression"/>
                                   </SuppressionDependentData>
      Return vDataStore

    End Function

    Private Sub ValidateActivityDependencies()
      Dim vData As XElement = LoadActivityReferenceTables()
      Dim vWhereValues As New Dictionary(Of String, String) From {{"ActivityColumn", mvClassFields("activity").Value}}
      Dim vValidators As List(Of DataDependencyValidator) = DataDependencyValidator.FromXElement(mvEnv, vData, "Name", "Description", vWhereValues)
      Dim vFailedValidators As List(Of DataDependencyValidator) = ValidateDataDependencies(vValidators)
      If vFailedValidators IsNot Nothing AndAlso vFailedValidators.Count > 0 Then
        RaiseError(DataAccessErrors.daeRecordCannotBeMadeHistoric, "Activity", vFailedValidators.AsCommaSeperated)
      End If
    End Sub

    Private Sub ValidateActivityValueDependencies()
      Dim vData As XElement = LoadActivityReferenceTables()
      Dim vWhereValues As New Dictionary(Of String, String) From
        {
          {"ActivityColumn", mvClassFields("activity").Value},
          {"ActivityValueColumn", mvClassFields("activity_value").Value}
        }
      Dim vValidators As List(Of DataDependencyValidator) = DataDependencyValidator.FromXElement(mvEnv, vData, "Name", "Description", vWhereValues)
      'vValidators.ForEach(Sub(vValidator) vValidator.WhereClause.Add(
      Dim vFailedValidators As List(Of DataDependencyValidator) = ValidateDataDependencies(vValidators)
      If vFailedValidators IsNot Nothing AndAlso vFailedValidators.Count > 0 Then
        RaiseError(DataAccessErrors.daeRecordCannotBeMadeHistoric, "Activity Value", vFailedValidators.AsCommaSeperated)
      End If
    End Sub

    Private Sub ValidateSuppressionDependencies()
      Dim vData As XElement = LoadSuppressionReferenceTables()
      Dim vWhereValues As New Dictionary(Of String, String) From {{"SuppressionColumn", mvClassFields("mailing_suppression").Value}}
      Dim vValidators As List(Of DataDependencyValidator) = DataDependencyValidator.FromXElement(mvEnv, vData, "Name", "Description", vWhereValues)
      Dim vFailedValidators As List(Of DataDependencyValidator) = ValidateDataDependencies(vValidators)
      If vFailedValidators IsNot Nothing AndAlso vFailedValidators.Count > 0 Then
        RaiseError(DataAccessErrors.daeRecordCannotBeMadeHistoric, "Mailing Suppression", vFailedValidators.AsCommaSeperated)
      End If
    End Sub

    Private Function ValidateDataDependencies(vValidators As List(Of DataDependencyValidator)) As List(Of DataDependencyValidator)
      Dim vRtn As New List(Of DataDependencyValidator)
      If vValidators IsNot Nothing Then
        For Each vEntry In vValidators
          If Not vEntry.Validate Then
            vRtn.Add(vEntry)
          End If
        Next
      End If
      Return vRtn
    End Function

    ''' <summary>When changing a CPD Cycle Type from fixed to flexible or vice versa make sure it's not already in use.</summary>
    ''' <param name="pParams"></param>
    ''' <returns>True if CPD Cycle can be updated, otherwise False</returns>
    ''' <remarks></remarks>
    Private Function CanUpdateCPDCycles(ByVal pParams As CDBParameters) As Boolean
      Dim vCanUpdate As Boolean = True
      If pParams.ContainsKey("CpdCycleType") Then
        Dim vWhereFields As New CDBFields(New CDBField("cct.cpd_cycle_type", CDBField.FieldTypes.cftCharacter, pParams("CpdCycleType").Value))
        If pParams.ParameterExists("StartMonth").Value.Length = 0 Then
          vWhereFields.Add("start_month", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual)
        Else
          vWhereFields.Add("start_month", CDBField.FieldTypes.cftInteger, "")
        End If
        Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("contact_cpd_cycles cpd", "cct.cpd_cycle_type", "cpd.cpd_cycle_type")})
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "", "cpd_cycle_types cct", vWhereFields, "", vAnsiJoins)
        If mvEnv.Connection.GetCountFromStatement(vSQLStatement) > 0 Then vCanUpdate = False
      End If
      Return vCanUpdate
    End Function
#Region "Surveys Validation"
    Public Sub PreValidateSurveyUpdateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.Exists("MaintenanceTableName") Then
        Select Case pParameterList("MaintenanceTableName").Value
          Case "survey_answers"
            ValidateSurveyAnswers(pParameterList)
          Case "survey_questions"
          Case "survey_versions"
          Case "surveys"
          Case "survey_contact_groups"
          Case Else
        End Select
      Else
      End If
    End Sub
    Public Sub PreValidateSurveyCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateCreateParameters(pParameterList)
      If pParameterList.Exists("MaintenanceTableName") Then
        Select Case pParameterList("MaintenanceTableName").Value
          Case "survey_answers"
            ValidateSurveyAnswers(pParameterList)
          Case "survey_questions"
          Case "survey_versions"
          Case "surveys"
          Case "survey_contact_groups"
          Case Else
        End Select
      Else
      End If
    End Sub
    ''' <summary>
    ''' Validate Paramaters for Survey Answers
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks>Validation is the same for Update and Create, so only one validation subroutine.</remarks>
    Private Sub ValidateSurveyAnswers(ByVal pParameterList As CDBParameters)
      Dim vSurveyAnswer As SurveyAnswer = SurveyAnswer.CreateInstance(mvEnv, pParameterList)
      vSurveyAnswer.ValidateSurveyQuestionParameter(pParameterList)
    End Sub

    Private Sub PostValidateUpdateOrganisationGroupParameters(ByVal pParameterList As CDBParameters)
      If mvExisting = True AndAlso pParameterList("MaintenanceTableName").Value.Equals("organisation_groups", StringComparison.InvariantCultureIgnoreCase) Then
        If mvClassFields.ContainsKey("view_in_contact_card") AndAlso mvClassFields.Item("view_in_contact_card").ValueChanged Then
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "organisation_number", "organisations", New CDBFields(New CDBField("organisation_group", mvClassFields.Item("organisation_group").Value)))
          If mvEnv.Connection.GetCountFromStatement(vSQLStatement) > 0 Then
            RaiseError(DataAccessErrors.daeViewInContactCardCannotBeChanged)
          End If
        End If
      End If
    End Sub
#End Region

    Public ReadOnly Property Field(pFieldName As String, Optional pReportError As Boolean = True) As ClassField
      Get
        Dim vResult As New ClassField(pFieldName, CDBField.FieldTypes.cftCharacter)
        If Me.ClassFields.ContainsKey(pFieldName) Then
          vResult = Me.ClassFields(pFieldName)
        ElseIf pReportError Then
          RaiseError(DataAccessErrors.daeIndexNotFound, pFieldName)
        End If
        Return vResult
      End Get
    End Property

  End Class
End Namespace
