

Namespace Access
  Public Class CriteriaSet

    Public Enum CriteriaSetRecordSetTypes 'These are bit values
      csrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CriteriaSetFields
      csfAll = 0
      csfCriteriaSet
      csfUserName
      csfDepartment
      csfCriteriaSetDesc
      csfCriteriaGroup
      csfReportCode
      csfStandardDocument
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvSelectionSteps As Collection
    Private mvCriteriaDetails As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "criteria_sets"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("criteria_set", CDBField.FieldTypes.cftLong)
          .Add("user_name")
          .Add("department")
          .Add("criteria_set_desc")
          .Add("criteria_group")
          .Add("report_code")
          .Add("standard_document")

          .Item(CriteriaSetFields.csfCriteriaSet).SetPrimaryKeyOnly()
          .Item(CriteriaSetFields.csfReportCode).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports)
          .Item(CriteriaSetFields.csfStandardDocument).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMailmergeHeaderOnReports)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvSelectionSteps = Nothing
      mvCriteriaDetails = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CriteriaSetFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(CriteriaSetFields.csfCriteriaSet).IntegerValue = 0 Then mvClassFields.Item(CriteriaSetFields.csfCriteriaSet).IntegerValue = mvEnv.GetControlNumber("CS")
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CriteriaSetRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CriteriaSetRecordSetTypes.csrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cs")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCriteriaSet As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCriteriaSet > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CriteriaSetRecordSetTypes.csrtAll) & " FROM criteria_sets cs WHERE criteria_set = " & pCriteriaSet)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CriteriaSetRecordSetTypes.csrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CriteriaSetRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CriteriaSetFields.csfCriteriaSet, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CriteriaSetRecordSetTypes.csrtAll) = CriteriaSetRecordSetTypes.csrtAll Then
          .SetItem(CriteriaSetFields.csfUserName, vFields)
          .SetItem(CriteriaSetFields.csfDepartment, vFields)
          .SetItem(CriteriaSetFields.csfCriteriaSetDesc, vFields)
          .SetItem(CriteriaSetFields.csfCriteriaGroup, vFields)
          .SetOptionalItem(CriteriaSetFields.csfReportCode, vFields)
          .SetOptionalItem(CriteriaSetFields.csfStandardDocument, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CriteriaSetFields.csfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pCriteriaSet As Integer, ByRef pOwner As String, ByRef pDepartment As String, _
                      ByRef pDescription As String, ByRef pGroup As String, Optional ByVal pReportCode As String = "", _
                      Optional ByVal pStandardDocument As String = "")
      With mvClassFields
        .Item(CriteriaSetFields.csfCriteriaSet).IntegerValue = pCriteriaSet
        .Item(CriteriaSetFields.csfUserName).Value = pOwner
        .Item(CriteriaSetFields.csfDepartment).Value = pDepartment
        .Item(CriteriaSetFields.csfCriteriaSetDesc).Value = pDescription
        .Item(CriteriaSetFields.csfCriteriaGroup).Value = pGroup
        .Item(CriteriaSetFields.csfReportCode).Value = pReportCode
        .Item(CriteriaSetFields.csfStandardDocument).Value = pStandardDocument
      End With
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Sub Clone(ByRef pTargetCSNumber As Integer, Optional ByVal pDetailsOnly As Boolean = False)
      Dim vTargetCriteriaSet As New CriteriaSet
      Dim vSelectionStep As SelectionStep
      Dim vNewSelectionStep As SelectionStep
      Dim vCriteriaDetail As CriteriaDetails
      Dim vNewCriteriaDetail As CriteriaDetails
      Dim vSequenceOffset As Integer
      Dim vDescription As String

      Try

        mvEnv.Connection.StartTransaction()

        'Validate if Source Criteria Set has Selection Steps and Selection Criteria (Criteria Set Details) and raise error if has both
        If SelectionSteps.Count > 0 AndAlso CriteriaSetDetails.Count > 0 Then
          RaiseError(DataAccessErrors.daeCannotCopySourceCriteriaSet)
        End If

        If pTargetCSNumber < 0 Then pTargetCSNumber = 0
        vTargetCriteriaSet.Init(mvEnv, pTargetCSNumber)
        If vTargetCriteriaSet.Existing Then
          'existing
        Else
          vDescription = "Criteria Set "
          If SelectionSteps.Count() > 0 Then vDescription = "List Manager Steps: "
          If pTargetCSNumber = 0 Then pTargetCSNumber = mvEnv.GetControlNumber("CS")
          vTargetCriteriaSet.Create(pTargetCSNumber, UserName, Department, vDescription & pTargetCSNumber, CriteriaGroup)
          If Not pDetailsOnly Then vTargetCriteriaSet.Save()
        End If

        'Validate that we aren't about to create Criteria with Criteria Set Details and Criteria Selection Steps as we can't have both
        If CriteriaSetDetails.Count > 0 AndAlso vTargetCriteriaSet.SelectionSteps.Count > 0 Then
          'Cannot copy criteria with selection criteria onto criteria with selection steps 
          'ImCannotCopyCampaignCriteriaSetSCSS
          RaiseError(DataAccessErrors.daeCannotCopySourceCriteriaSet)
        ElseIf SelectionSteps.Count > 0 AndAlso vTargetCriteriaSet.CriteriaSetDetails.Count > 0 Then
          'Cannot copy criteria with selection steps onto criteria with criteria set details
          RaiseError(DataAccessErrors.daeCannotCopySourceCriteriaSet)
        End If

        vSequenceOffset = CInt(Val(mvEnv.Connection.GetValue("SELECT MAX(sequence_number) FROM selection_steps WHERE criteria_set = " & pTargetCSNumber)))
        For Each vSelectionStep In SelectionSteps
          With vSelectionStep
            vNewSelectionStep = New SelectionStep
            vNewSelectionStep.Init(mvEnv)
            vNewSelectionStep.Create(pTargetCSNumber, .ViewName, vSequenceOffset + .SequenceNumber, .FilterSql, .SelectAction, .RecordCount)
            vNewSelectionStep.Save()
          End With
        Next vSelectionStep
        vSequenceOffset = CInt(Val(mvEnv.Connection.GetValue("SELECT MAX(sequence_number) FROM criteria_set_details WHERE criteria_set = " & pTargetCSNumber)))
        For Each vCriteriaDetail In CriteriaSetDetails
          With vCriteriaDetail
            vNewCriteriaDetail = New CriteriaDetails
            vNewCriteriaDetail.Init(mvEnv)
            vNewCriteriaDetail.Create(pTargetCSNumber, vSequenceOffset + .SequenceNumber, .SearchArea, .IncludeOrExclude, .ContactOrOrganisation, .MainValue, .SubsidiaryValue, .Period, CStr(.Counted), .AndOr, .LeftParenthesis, .RightParenthesis)
            vNewCriteriaDetail.Save()
          End With
        Next vCriteriaDetail
        mvEnv.Connection.CommitTransaction()
      Catch vException As Exception
        If mvEnv.Connection.InTransaction Then mvEnv.Connection.RollbackTransaction()
        Throw vException
      End Try
    End Sub
    Public ReadOnly Property SelectionSteps() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vStep As New SelectionStep

        If mvSelectionSteps Is Nothing Then
          mvSelectionSteps = New Collection
          If mvExisting And mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.CDBDataViewNames) Then
            vStep.Init(mvEnv)
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vStep.GetRecordSetFields(SelectionStep.SelectionStepRecordSetTypes.sstprtAll) & " FROM selection_steps sst WHERE criteria_set = " & CriteriaSetNumber)
            While vRecordSet.Fetch() = True
              vStep = New SelectionStep
              vStep.InitFromRecordSet(mvEnv, vRecordSet, SelectionStep.SelectionStepRecordSetTypes.sstprtAll)
              mvSelectionSteps.Add(vStep)
            End While
            vRecordSet.CloseRecordSet()
          End If
        End If
        SelectionSteps = mvSelectionSteps
      End Get
    End Property

    Public ReadOnly Property CriteriaSetDetails() As Collection
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vCriteriaDetail As New CriteriaDetails

        If mvCriteriaDetails Is Nothing Then
          mvCriteriaDetails = New Collection
          vCriteriaDetail.Init(mvEnv)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vCriteriaDetail.GetRecordSetFields(CriteriaDetails.CriteriaSetDetailRecordSetTypes.csdrtAll) & " FROM criteria_set_details csd WHERE criteria_set = " & CriteriaSetNumber & " ORDER BY sequence_number")
          While vRecordSet.Fetch() = True
            vCriteriaDetail = New CriteriaDetails
            vCriteriaDetail.InitFromRecordSet(mvEnv, vRecordSet, CriteriaDetails.CriteriaSetDetailRecordSetTypes.csdrtAll)
            mvCriteriaDetails.Add(vCriteriaDetail)
          End While
          vRecordSet.CloseRecordSet()
        End If
        CriteriaSetDetails = mvCriteriaDetails
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property CriteriaGroup() As String
      Get
        CriteriaGroup = mvClassFields.Item(CriteriaSetFields.csfCriteriaGroup).Value
      End Get
    End Property

    Public ReadOnly Property CriteriaSetNumber() As Integer
      Get
        CriteriaSetNumber = mvClassFields.Item(CriteriaSetFields.csfCriteriaSet).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CriteriaSetDesc() As String
      Get
        CriteriaSetDesc = mvClassFields.Item(CriteriaSetFields.csfCriteriaSetDesc).Value
      End Get
    End Property

    Public ReadOnly Property Department() As String
      Get
        Department = mvClassFields.Item(CriteriaSetFields.csfDepartment).Value
      End Get
    End Property

    Public ReadOnly Property UserName() As String
      Get
        UserName = mvClassFields.Item(CriteriaSetFields.csfUserName).Value
      End Get
    End Property
    Public ReadOnly Property ReportCode() As String
      Get
        ReportCode = mvClassFields.Item(CriteriaSetFields.csfReportCode).Value
      End Get
    End Property
    Public ReadOnly Property StandardDocument() As String
      Get
        StandardDocument = mvClassFields.Item(CriteriaSetFields.csfStandardDocument).Value
      End Get
    End Property

    Public Sub Update(ByVal pCriteriaSet As String, ByVal pCriteriaSetDescription As String, ByVal pUserName As String, ByVal pDepartment As String)
      With mvClassFields
        .Item(CriteriaSetFields.csfCriteriaSet).Value = pCriteriaSet
        .Item(CriteriaSetFields.csfCriteriaSetDesc).Value = pCriteriaSetDescription
        .Item(CriteriaSetFields.csfUserName).Value = pUserName
        .Item(CriteriaSetFields.csfDepartment).Value = pDepartment
      End With
    End Sub

    Public Sub Delete(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields()
      Dim vCriteriaDetail As CriteriaDetails

      'Delete Selection Steps
      vWhereFields.Add("criteria_set", CriteriaSetNumber)
      mvEnv.Connection.DeleteRecords("selection_steps", vWhereFields, False)

      'Delete Criteria Set Details
      For Each vCriteriaDetail In CriteriaSetDetails
        vCriteriaDetail.Delete()
      Next

      'Delete CriteriaSet
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub
  End Class
End Namespace
