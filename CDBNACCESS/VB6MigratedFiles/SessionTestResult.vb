

Namespace Access
  Public Class SessionTestResult

    Public Enum SessionTestResultRecordSetTypes 'These are bit values
      strrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SessionTestResultFields
      strfAll = 0
      strfSessionNumber
      strfContactNumber
      strfTestNumber
      strfTestResult
      strfCertificateNumber
      strfNotes
      strfAmendedBy
      strfAmendedOn
    End Enum

    Private mvSessionTest As SessionTest
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
          .DatabaseTableName = "session_test_results"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("session_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("test_number", CDBField.FieldTypes.cftInteger)
          .Add("test_result")
          .Add("certificate_number")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(SessionTestResultFields.strfSessionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(SessionTestResultFields.strfContactNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(SessionTestResultFields.strfTestNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

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
        AmendedBy = mvClassFields.Item(SessionTestResultFields.strfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(SessionTestResultFields.strfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CertificateNumber() As String
      Get
        CertificateNumber = mvClassFields.Item(SessionTestResultFields.strfCertificateNumber).Value
      End Get
    End Property
    Public ReadOnly Property SessionTest() As SessionTest
      Get
        If mvSessionTest Is Nothing Then
          If SessionNumber > 0 And TestNumber > 0 Then
            mvSessionTest = New SessionTest
            mvSessionTest.Init(mvEnv, SessionNumber, TestNumber)
          End If
        End If
        SessionTest = mvSessionTest
      End Get
    End Property
    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(SessionTestResultFields.strfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionTestResultFields.strfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(SessionTestResultFields.strfNotes).MultiLineValue
      End Get
    End Property

    Public Property SessionNumber() As Integer
      Get
        SessionNumber = mvClassFields.Item(SessionTestResultFields.strfSessionNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SessionTestResultFields.strfSessionNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property TestNumber() As Integer
      Get
        TestNumber = mvClassFields.Item(SessionTestResultFields.strfTestNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TestResult() As String
      Get
        TestResult = mvClassFields.Item(SessionTestResultFields.strfTestResult).Value
      End Get
    End Property
    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SessionTestResultFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(SessionTestResultFields.strfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(SessionTestResultFields.strfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SessionTestResultRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SessionTestResultRecordSetTypes.strrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "str")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pSessionNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0, Optional ByRef pTestNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSessionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SessionTestResultRecordSetTypes.strrtAll) & " FROM session_test_results str WHERE session_number = " & pSessionNumber & " AND contact_number = " & pContactNumber & " AND test_number = " & pTestNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SessionTestResultRecordSetTypes.strrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SessionTestResultRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SessionTestResultFields.strfSessionNumber, vFields)
        .SetItem(SessionTestResultFields.strfContactNumber, vFields)
        .SetItem(SessionTestResultFields.strfTestNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SessionTestResultRecordSetTypes.strrtAll) = SessionTestResultRecordSetTypes.strrtAll Then
          .SetItem(SessionTestResultFields.strfTestResult, vFields)
          .SetItem(SessionTestResultFields.strfCertificateNumber, vFields)
          .SetItem(SessionTestResultFields.strfNotes, vFields)
          .SetItem(SessionTestResultFields.strfAmendedBy, vFields)
          .SetItem(SessionTestResultFields.strfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(SessionTestResultFields.strfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(SessionTestResultFields.strfSessionNumber).Value = pParams("SessionNumber").Value
        .Item(SessionTestResultFields.strfContactNumber).Value = pParams("ContactNumber").Value
        .Item(SessionTestResultFields.strfTestNumber).Value = pParams("TestNumber").Value
        .Item(SessionTestResultFields.strfTestResult).Value = pParams("TestResult").Value
        If pParams.Exists("CertificateNumber") Then .Item(SessionTestResultFields.strfCertificateNumber).Value = pParams("CertificateNumber").Value
        If pParams.Exists("Notes") Then .Item(SessionTestResultFields.strfNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("TestResult") Then .Item(SessionTestResultFields.strfTestResult).Value = pParams("TestResult").Value
        If pParams.Exists("CertificateNumber") Then .Item(SessionTestResultFields.strfCertificateNumber).Value = pParams("CertificateNumber").Value
        If pParams.Exists("Notes") Then .Item(SessionTestResultFields.strfNotes).Value = pParams("Notes").Value
      End With
    End Sub
  End Class
End Namespace
