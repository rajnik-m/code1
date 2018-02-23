

Namespace Access
  Public Class EventSubmission

    Public Enum EventSubmissionRecordSetTypes 'These are bit values
      esrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventSubmissionFields
      esfAll = 0
      esfEventNumber
      esfPaperTitle
      esfSubject
      esfSkillLevel
      esfSubmitted
      esfContactNumber
      esfAddressNumber
      esfAssessor
      esfForwarded
      esfReturned
      esfResult
      esfAmendedBy
      esfAmendedOn
      esfSubmissionNumber
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "event_submissions"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("paper_title")
          .Add("subject")
          .Add("skill_level")
          .Add("submitted", CDBField.FieldTypes.cftDate)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("assessor", CDBField.FieldTypes.cftLong)
          .Add("forwarded", CDBField.FieldTypes.cftDate)
          .Add("returned", CDBField.FieldTypes.cftDate)
          .Add("result")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("submission_number", CDBField.FieldTypes.cftLong)
        End With

        mvClassFields.Item(EventSubmissionFields.esfEventNumber).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfSubject).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfSkillLevel).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfContactNumber).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfAddressNumber).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfAmendedBy).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfAmendedOn).PrefixRequired = True
        mvClassFields.Item(EventSubmissionFields.esfSubmissionNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(EventSubmissionFields.esfSubmitted).Value = TodaysDate()
    End Sub

    Private Sub SetValid(ByVal pField As EventSubmissionFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventSubmissionFields.esfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventSubmissionFields.esfAmendedBy).Value = mvEnv.User.Logname
      If mvClassFields.Item(EventSubmissionFields.esfSubmissionNumber).IntegerValue = 0 Then
        mvClassFields.Item(EventSubmissionFields.esfSubmissionNumber).Value = CStr(mvEnv.GetControlNumber("ES"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventSubmissionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventSubmissionRecordSetTypes.esrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "es")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSubmissionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSubmissionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventSubmissionRecordSetTypes.esrtAll) & " FROM event_submissions es WHERE submission_number = " & pSubmissionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventSubmissionRecordSetTypes.esrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventSubmissionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventSubmissionFields.esfSubmissionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventSubmissionRecordSetTypes.esrtAll) = EventSubmissionRecordSetTypes.esrtAll Then
          .SetItem(EventSubmissionFields.esfEventNumber, vFields)
          .SetItem(EventSubmissionFields.esfPaperTitle, vFields)
          .SetItem(EventSubmissionFields.esfSubject, vFields)
          .SetItem(EventSubmissionFields.esfSkillLevel, vFields)
          .SetItem(EventSubmissionFields.esfSubmitted, vFields)
          .SetItem(EventSubmissionFields.esfContactNumber, vFields)
          .SetItem(EventSubmissionFields.esfAddressNumber, vFields)
          .SetItem(EventSubmissionFields.esfAssessor, vFields)
          .SetItem(EventSubmissionFields.esfForwarded, vFields)
          .SetItem(EventSubmissionFields.esfReturned, vFields)
          .SetItem(EventSubmissionFields.esfResult, vFields)
          .SetItem(EventSubmissionFields.esfAmendedBy, vFields)
          .SetItem(EventSubmissionFields.esfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub InitFromEvent(ByVal pEnv As CDBEnvironment, ByVal pEvent As CDBEvent)
      mvEnv = pEnv
      InitClassFields()
      mvClassFields.Item(EventSubmissionFields.esfEventNumber).Value = CStr(pEvent.EventNumber)
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventSubmissionFields.esfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      mvClassFields.Item(EventSubmissionFields.esfEventNumber).Value = pParams("EventNumber").Value
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("PaperTitle") Then .Item(EventSubmissionFields.esfPaperTitle).Value = pParams("PaperTitle").Value
        If pParams.Exists("Subject") Then .Item(EventSubmissionFields.esfSubject).Value = pParams("Subject").Value
        If pParams.Exists("SkillLevel") Then .Item(EventSubmissionFields.esfSkillLevel).Value = pParams("SkillLevel").Value
        If pParams.Exists("Submitted") Then .Item(EventSubmissionFields.esfSubmitted).Value = pParams("Submitted").Value
        If pParams.Exists("ContactNumber") Then .Item(EventSubmissionFields.esfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(EventSubmissionFields.esfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("Assessor") Then .Item(EventSubmissionFields.esfAssessor).Value = pParams("Assessor").Value
        If pParams.Exists("Forwarded") Then .Item(EventSubmissionFields.esfForwarded).Value = pParams("Forwarded").Value
        If pParams.Exists("Returned") Then .Item(EventSubmissionFields.esfReturned).Value = pParams("Returned").Value
        If pParams.Exists("SubmissionResult") Then .Item(EventSubmissionFields.esfResult).Value = pParams("SubmissionResult").Value
      End With
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(EventSubmissionFields.esfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(EventSubmissionFields.esfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(EventSubmissionFields.esfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Assessor() As String
      Get
        Assessor = mvClassFields.Item(EventSubmissionFields.esfAssessor).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(EventSubmissionFields.esfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(EventSubmissionFields.esfEventNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Forwarded() As String
      Get
        Forwarded = mvClassFields.Item(EventSubmissionFields.esfForwarded).Value
      End Get
    End Property

    Public ReadOnly Property PaperTitle() As String
      Get
        PaperTitle = mvClassFields.Item(EventSubmissionFields.esfPaperTitle).Value
      End Get
    End Property

    Public ReadOnly Property Result() As String
      Get
        Result = mvClassFields.Item(EventSubmissionFields.esfResult).Value
      End Get
    End Property

    Public ReadOnly Property Returned() As String
      Get
        Returned = mvClassFields.Item(EventSubmissionFields.esfReturned).Value
      End Get
    End Property

    Public ReadOnly Property SkillLevel() As String
      Get
        SkillLevel = mvClassFields.Item(EventSubmissionFields.esfSkillLevel).Value
      End Get
    End Property

    Public ReadOnly Property Subject() As String
      Get
        Subject = mvClassFields.Item(EventSubmissionFields.esfSubject).Value
      End Get
    End Property

    Public ReadOnly Property SubmissionNumber() As Integer
      Get
        SubmissionNumber = mvClassFields.Item(EventSubmissionFields.esfSubmissionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Submitted() As String
      Get
        Submitted = mvClassFields.Item(EventSubmissionFields.esfSubmitted).Value
      End Get
    End Property

  End Class
End Namespace
