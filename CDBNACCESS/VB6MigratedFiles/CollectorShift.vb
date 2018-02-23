

Namespace Access
  Public Class CollectorShift

    Public Enum CollectorShiftRecordSetTypes 'These are bit values
      csrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectorShiftFields
      csfAll = 0
      csfCollectorShiftNumber
      csfCollectorNumber
      csfCollectionPointNumber
      csfStartTime
      csfEndTime
      csfNotes
      csfAmendedBy
      csfAmendedOn
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
          .DatabaseTableName = "collector_shifts"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collector_shift_number", CDBField.FieldTypes.cftLong)
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("collection_point_number", CDBField.FieldTypes.cftLong)
          .Add("start_time", CDBField.FieldTypes.cftCharacter)
          .Add("end_time", CDBField.FieldTypes.cftCharacter)
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

        End With

        mvClassFields.Item(CollectorShiftFields.csfCollectorShiftNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectorShiftFields)
      'Add code here to ensure all values are valid before saving
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vContactNumber As String = ""
      Dim vCollectionNumber As String = ""

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT contact_number, collection_number FROM manned_collectors WHERE collector_number = " & CollectorNumber)
      If vRecordSet.Fetch() = True Then
        vContactNumber = vRecordSet.Fields("contact_number").Value
        vCollectionNumber = vRecordSet.Fields("collection_number").Value
      Else
        RaiseError(DataAccessErrors.daeCollectorNotFound, CStr(CollectorNumber))
      End If
      With vWhereFields
        .Add("mc.contact_number", CDBField.FieldTypes.cftLong, vContactNumber)
        .Add("cs.collector_number", CDBField.FieldTypes.cftLong, "mc.collector_number")
        '.Add "cs.start_time", cftCharacter, StartTime, fwoLessThanEqual + fwoOpenBracket + fwoOpenBracket

        '(vNewStartTime >= vStartTime And vNewStartTime <= vEndTime) Or (vNewEndTime >= vStartTime And vNewEndTime <= vEndTime)
        '.Add "cs.end_time", cftCharacter, EndTime, fwoGreaterThanEqual
        .Add("mco.collection_number", CDBField.FieldTypes.cftLong, "mc.collection_number")
        .Add("mco.collection_date", CDBField.FieldTypes.cftLong, "(SELECT collection_date FROM manned_collections WHERE collection_number = " & vCollectionNumber & ")")
        If CollectorShiftNumber > 0 Then .Add("collector_shift_number", CollectorShiftNumber, CDBField.FieldWhereOperators.fwoNotEqual)
      End With
      If mvEnv.Connection.GetCount("manned_collectors mc, collector_shifts cs, manned_collections mco", Nothing, mvEnv.Connection.WhereClause(vWhereFields) & " AND (( cs.start_time" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, StartTime) & " AND cs.end_time" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartTime) & ") OR (cs.start_time" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, EndTime) & " AND cs.end_time" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, EndTime) & "))") > 0 Then
        RaiseError(DataAccessErrors.daeCollectorShiftsOverlap)
      End If
      If mvClassFields.Item(CollectorShiftFields.csfCollectorShiftNumber).IntegerValue = 0 Then mvClassFields.Item(CollectorShiftFields.csfCollectorShiftNumber).IntegerValue = mvEnv.GetControlNumber("CH")
      mvClassFields.Item(CollectorShiftFields.csfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectorShiftFields.csfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectorShiftRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CollectorShiftRecordSetTypes.csrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cs")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectorShiftNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectorShiftNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectorShiftRecordSetTypes.csrtAll) & " FROM collector_shifts cs WHERE collector_shift_number = " & pCollectorShiftNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectorShiftRecordSetTypes.csrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectorShiftRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectorShiftFields.csfCollectorShiftNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectorShiftRecordSetTypes.csrtAll) = CollectorShiftRecordSetTypes.csrtAll Then
          .SetItem(CollectorShiftFields.csfCollectionPointNumber, vFields)
          .SetItem(CollectorShiftFields.csfStartTime, vFields)
          .SetItem(CollectorShiftFields.csfEndTime, vFields)
          .SetItem(CollectorShiftFields.csfNotes, vFields)
          .SetItem(CollectorShiftFields.csfAmendedBy, vFields)
          .SetItem(CollectorShiftFields.csfAmendedOn, vFields)
          .SetItem(CollectorShiftFields.csfCollectorNumber, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectorShiftFields.csfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    'Public Sub Create(ByVal pCollectorNumber As Long, ByVal pCollectionPointNumber As Long, ByVal pStartTime As String, ByVal pEndTime As String, pNotes As String)
    '  With mvClassFields
    '    .Item(csfCollectorNumber).IntegerValue = pCollectorNumber
    '    .Item(csfCollectionPointNumber).IntegerValue = pCollectionPointNumber
    '    .Item(csfStartTime).Value = pStartTime
    '    .Item(csfEndTime).Value = pEndTime
    '    .Item(csfNotes).Value = pNotes
    '  End With
    'End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(CollectorShiftFields.csfCollectorNumber).Value = pParams("CollectorNumber").Value
        .Item(CollectorShiftFields.csfCollectionPointNumber).IntegerValue = pParams("CollectionPointNumber").IntegerValue
        If pParams.Exists("StartTime") Then .Item(CollectorShiftFields.csfStartTime).Value = pParams("StartTime").Value
        If pParams.Exists("EndTime") Then .Item(CollectorShiftFields.csfEndTime).Value = pParams("EndTime").Value
        If pParams.Exists("Notes") Then .Item(CollectorShiftFields.csfNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("StartTime") Then .Item(CollectorShiftFields.csfStartTime).Value = pParams("StartTime").Value
        If pParams.Exists("EndTime") Then .Item(CollectorShiftFields.csfEndTime).Value = pParams("EndTime").Value
        If pParams.Exists("Notes") Then .Item(CollectorShiftFields.csfNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(CollectorShiftFields.csfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectorShiftFields.csfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CollectionPointNumber() As Integer
      Get
        CollectionPointNumber = CInt(mvClassFields.Item(CollectorShiftFields.csfCollectionPointNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        CollectorNumber = CInt(mvClassFields.Item(CollectorShiftFields.csfCollectorNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectorShiftNumber() As Integer
      Get
        CollectorShiftNumber = mvClassFields.Item(CollectorShiftFields.csfCollectorShiftNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EndTime() As String
      Get
        EndTime = mvClassFields.Item(CollectorShiftFields.csfEndTime).Value
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(CollectorShiftFields.csfNotes).Value
      End Get
    End Property

    Public ReadOnly Property StartTime() As String
      Get
        StartTime = mvClassFields.Item(CollectorShiftFields.csfStartTime).Value
      End Get
    End Property
  End Class
End Namespace
