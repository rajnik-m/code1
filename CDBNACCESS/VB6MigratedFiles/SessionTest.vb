

Namespace Access
  Public Class SessionTest

    Public Enum SessionTestRecordSetTypes 'These are bit values
      estrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SessionTestFields
      stfAll = 0
      stfSessionNumber
      stfTestNumber
      stfTestDesc
      stfGradeDataType
      stfMinimumValue
      stfMaximumValue
      stfPattern
      stfAmendedBy
      stfAmendedOn
    End Enum

    Public Enum GradeDataTypes
      gdtCharacter
      gdtInteger
      gdtNumeric
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
          .DatabaseTableName = "session_tests"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("test_number", CDBField.FieldTypes.cftInteger)
          .Add("test_desc")
          .Add("grade_data_type")
          .Add("minimum_value")
          .Add("maximum_value")
          .Add("pattern")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(SessionTestFields.stfAmendedBy).PrefixRequired = True
        mvClassFields.Item(SessionTestFields.stfSessionNumber).PrefixRequired = True
        mvClassFields.Item(SessionTestFields.stfAmendedOn).PrefixRequired = True

        mvClassFields.Item(SessionTestFields.stfSessionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(SessionTestFields.stfTestNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SessionTestFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SessionTestFields.stfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SessionTestFields.stfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SessionTestRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SessionTestRecordSetTypes.estrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "st")
      End If
      Return vFields
    End Function

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
        AmendedBy = mvClassFields.Item(SessionTestFields.stfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SessionTestFields.stfAmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property GradeCodeFromType(ByVal pGradeType As GradeDataTypes) As String
      Get
        Select Case pGradeType
          Case GradeDataTypes.gdtCharacter
            Return "C"
          Case GradeDataTypes.gdtInteger
            Return "I"
          Case GradeDataTypes.gdtNumeric
            Return "N"
          Case Else
            Return ""           'Added fix for compiler warning
        End Select
      End Get
    End Property
    Public ReadOnly Property GradeDataType() As GradeDataTypes
      Get
        GradeDataType = GradeTypeFromCode((mvClassFields.Item(SessionTestFields.stfGradeDataType).Value))
      End Get
    End Property

    Public ReadOnly Property MaximumValue() As String
      Get
        MaximumValue = mvClassFields.Item(SessionTestFields.stfMaximumValue).Value
      End Get
    End Property

    Public ReadOnly Property MinimumValue() As String
      Get
        MinimumValue = mvClassFields.Item(SessionTestFields.stfMinimumValue).Value
      End Get
    End Property

    Public ReadOnly Property Pattern() As String
      Get
        Pattern = mvClassFields.Item(SessionTestFields.stfPattern).Value
      End Get
    End Property

    Public Property SessionNumber() As Integer
      Get
        SessionNumber = mvClassFields.Item(SessionTestFields.stfSessionNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionTestFields.stfSessionNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property TestDesc() As String
      Get
        TestDesc = mvClassFields.Item(SessionTestFields.stfTestDesc).Value
      End Get
    End Property

    Public ReadOnly Property TestNumber() As Integer
      Get
        TestNumber = mvClassFields.Item(SessionTestFields.stfTestNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ChangesAllowed() As Boolean
      Get
        Dim vWhereFields As CDBFields

        vWhereFields = New CDBFields
        vWhereFields.Add("test_number", CDBField.FieldTypes.cftLong, TestNumber)
        vWhereFields.Add("session_number", CDBField.FieldTypes.cftLong, SessionNumber)
        If mvEnv.Connection.GetCount("session_test_results", vWhereFields) > 0 Then
          ChangesAllowed = False ' results exist
        Else
          ChangesAllowed = True
        End If
      End Get
    End Property
    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSessionNumber As Integer = 0, Optional ByRef pTestNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSessionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SessionTestRecordSetTypes.estrtAll) & " FROM session_tests st WHERE session_number = " & pSessionNumber & " AND test_number = " & pTestNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SessionTestRecordSetTypes.estrtAll)
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
    Friend Sub InitFromSessionTest(ByVal pOriginalEvent As CDBEvent, ByRef pSessionTest As SessionTest, ByRef pNewEvent As CDBEvent)
      With pSessionTest
        'mvClassFields.Item(SessionTestFields.stfSessionNumber).Value = pNewEvent    'CStr(pNewEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfSessionNumber))     'CStr(pNewEvent.BaseItemNumber + (.SessionNumber Mod 10000))
        mvClassFields.Item(SessionTestFields.stfSessionNumber).IntegerValue = pNewEvent.Sessions(pOriginalEvent.Sessions.IndexOf(pOriginalEvent.Sessions.Item(.SessionNumber.ToString))).SessionNumber
        mvClassFields.Item(SessionTestFields.stfTestNumber).Value = CStr(.TestNumber)
        mvClassFields.Item(SessionTestFields.stfTestDesc).Value = .TestDesc
        mvClassFields.Item(SessionTestFields.stfGradeDataType).Value = .GradeCodeFromType(.GradeDataType)
        mvClassFields.Item(SessionTestFields.stfMinimumValue).Value = .MinimumValue
        mvClassFields.Item(SessionTestFields.stfMaximumValue).Value = .MaximumValue
        mvClassFields.Item(SessionTestFields.stfPattern).Value = .Pattern
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SessionTestRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SessionTestFields.stfSessionNumber, vFields)
        .SetItem(SessionTestFields.stfTestNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SessionTestRecordSetTypes.estrtAll) = SessionTestRecordSetTypes.estrtAll Then
          .SetItem(SessionTestFields.stfTestDesc, vFields)
          .SetItem(SessionTestFields.stfGradeDataType, vFields)
          .SetItem(SessionTestFields.stfMinimumValue, vFields)
          .SetItem(SessionTestFields.stfMaximumValue, vFields)
          .SetItem(SessionTestFields.stfPattern, vFields)
          .SetItem(SessionTestFields.stfAmendedBy, vFields)
          .SetItem(SessionTestFields.stfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(SessionTestFields.stfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(SessionTestFields.stfSessionNumber).Value = pParams("SessionNumber").Value
        .Item(SessionTestFields.stfTestNumber).Value = pParams("TestNumber").Value
        .Item(SessionTestFields.stfTestDesc).Value = pParams("TestDesc").Value
        .Item(SessionTestFields.stfGradeDataType).Value = pParams("GradeDataType").Value
        If pParams.Exists("MinimumValue") Then .Item(SessionTestFields.stfMinimumValue).Value = pParams("MinimumValue").Value
        If pParams.Exists("MaximumValue") Then .Item(SessionTestFields.stfMaximumValue).Value = pParams("MaximumValue").Value
        If pParams.Exists("Pattern") Then .Item(SessionTestFields.stfPattern).Value = pParams("Pattern").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("TestDesc") Then .Item(SessionTestFields.stfTestDesc).Value = pParams("TestDesc").Value
        If pParams.Exists("GradeDataType") Then .Item(SessionTestFields.stfGradeDataType).Value = pParams("GradeDataType").Value
        If pParams.Exists("MinimumValue") Then .Item(SessionTestFields.stfMinimumValue).Value = pParams("MinimumValue").Value
        If pParams.Exists("MaximumValue") Then .Item(SessionTestFields.stfMaximumValue).Value = pParams("MaximumValue").Value
        If pParams.Exists("Pattern") Then .Item(SessionTestFields.stfPattern).Value = pParams("Pattern").Value
      End With
    End Sub
    Public Function GradeTypeFromCode(ByRef pGradeTypeCode As String) As GradeDataTypes
      Select Case pGradeTypeCode
        Case "C"
          GradeTypeFromCode = GradeDataTypes.gdtCharacter
        Case "I"
          GradeTypeFromCode = GradeDataTypes.gdtInteger
        Case "N"
          GradeTypeFromCode = GradeDataTypes.gdtNumeric
      End Select
    End Function
    Public Function IsValid(ByVal pResult As String) As Boolean
      'Validate whether given Result according to Session Test definitions
      Dim vValid As Boolean

      vValid = True
      If vValid Then
        Select Case GradeDataType
          Case GradeDataTypes.gdtCharacter
            If Pattern.Length > 0 Then
              If InStr(Pattern, pResult) = 0 Then vValid = False
            End If
            If vValid Then
              If MinimumValue.Length > 0 And pResult < MinimumValue Then vValid = False
            End If
            If vValid Then
              If MaximumValue.Length > 0 And pResult > MaximumValue Then vValid = False
            End If
          Case GradeDataTypes.gdtInteger, GradeDataTypes.gdtNumeric
            If Not IsNumeric(pResult) Then
              vValid = False
            Else
              If GradeDataType = GradeDataTypes.gdtInteger And System.Math.Abs(Val(pResult)) <> System.Math.Round(Val(pResult)) Then vValid = False
            End If
            If vValid Then
              If MinimumValue.Length > 0 And Val(pResult) < Val(MinimumValue) Then vValid = False
            End If
            If vValid Then
              If MaximumValue.Length > 0 And Val(pResult) > Val(MaximumValue) Then vValid = False
            End If
        End Select
      End If

      IsValid = vValid
    End Function

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub
  End Class
End Namespace
