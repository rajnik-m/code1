

Namespace Access
  Public Class EventBookingOption

    Public Enum EventBookingOptionRecordSetTypes 'These are bit values
      ebortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum EventBookingOptionFields
      ebofAll = 0
      ebofEventNumber
      ebofOptionNumber
      ebofOptionDesc
      ebofPickSessions
      ebofNumberOfSessions
      ebofDeductFromEvent
      ebofMaximumBookings
      ebofMinimumBookings
      ebofProduct
      ebofRate
      ebofLongDescription
      ebofFreeOfCharge
      ebofAmendedBy
      ebofAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvOptionSessions As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "event_booking_options"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("event_number", CDBField.FieldTypes.cftInteger)
          .Add("option_number", CDBField.FieldTypes.cftLong)
          .Add("option_desc")
          .Add("pick_sessions")
          .Add("number_of_sessions", CDBField.FieldTypes.cftInteger)
          .Add("deduct_from_event")
          .Add("maximum_bookings", CDBField.FieldTypes.cftInteger)
          .Add("minimum_bookings", CDBField.FieldTypes.cftInteger)
          .Add("product")
          .Add("rate")
          .Add("long_description")
          .Add("free_of_charge")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With
        mvClassFields.Item(EventBookingOptionFields.ebofOptionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(EventBookingOptionFields.ebofLongDescription).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLongDescription)
        mvClassFields.Item(EventBookingOptionFields.ebofMinimumBookings).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEventMinimumBookings)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvClassFields.Item(EventBookingOptionFields.ebofMaximumBookings).IntegerValue = 1
      mvClassFields.Item(EventBookingOptionFields.ebofMinimumBookings).IntegerValue = 1
    End Sub

    Private Sub SetValid(ByRef pField As EventBookingOptionFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(EventBookingOptionFields.ebofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(EventBookingOptionFields.ebofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As EventBookingOptionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = EventBookingOptionRecordSetTypes.ebortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ebo")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pOptionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pOptionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventBookingOptionRecordSetTypes.ebortAll) & " FROM event_booking_options WHERE option_number = " & pOptionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, EventBookingOptionRecordSetTypes.ebortAll)
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

    Public Sub InitFromBookingOption(ByRef pBookingOption As EventBookingOption, ByRef pNewEvent As CDBEvent, Optional ByRef pCopyOfBookingOption As Boolean = False)
      With pBookingOption
        mvClassFields.Item(EventBookingOptionFields.ebofEventNumber).Value = CStr(pNewEvent.EventNumber)
        If pCopyOfBookingOption Then
          mvClassFields.Item(EventBookingOptionFields.ebofOptionNumber).Value = CStr(pNewEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfOptionNumber))
          mvClassFields.Item(EventBookingOptionFields.ebofOptionDesc).Value = .OptionDesc
          If Len(OptionDesc) < 33 Then mvClassFields.Item(EventBookingOptionFields.ebofOptionDesc).Value = OptionDesc & (ProjectText.String25512) ' (copy)
        Else
          mvClassFields.Item(EventBookingOptionFields.ebofOptionNumber).Value = CStr(pNewEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfOptionNumber)) 'CStr(pNewEvent.BaseItemNumber + (.OptionNumber Mod 10000))
          mvClassFields.Item(EventBookingOptionFields.ebofOptionDesc).Value = .OptionDesc
        End If
        mvClassFields.Item(EventBookingOptionFields.ebofPickSessions).Bool = .PickSessions
        mvClassFields.Item(EventBookingOptionFields.ebofNumberOfSessions).Value = CStr(.NumberOfSessions)
        mvClassFields.Item(EventBookingOptionFields.ebofDeductFromEvent).Bool = .DeductFromEvent
        mvClassFields.Item(EventBookingOptionFields.ebofMaximumBookings).Value = CStr(.MaximumBookings)
        mvClassFields.Item(EventBookingOptionFields.ebofMinimumBookings).Value = CStr(.MinimumBookings)
        mvClassFields.Item(EventBookingOptionFields.ebofProduct).Value = .ProductCode
        mvClassFields.Item(EventBookingOptionFields.ebofRate).Value = .RateCode
        mvClassFields.Item(EventBookingOptionFields.ebofLongDescription).Value = .LongDescription
        mvClassFields.Item(EventBookingOptionFields.ebofFreeOfCharge).Value = .FreeOfCharge
      End With
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventBookingOptionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(EventBookingOptionFields.ebofOptionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And EventBookingOptionRecordSetTypes.ebortAll) = EventBookingOptionRecordSetTypes.ebortAll Then
          .SetItem(EventBookingOptionFields.ebofEventNumber, vFields)
          .SetItem(EventBookingOptionFields.ebofOptionDesc, vFields)
          .SetItem(EventBookingOptionFields.ebofPickSessions, vFields)
          .SetItem(EventBookingOptionFields.ebofNumberOfSessions, vFields)
          .SetItem(EventBookingOptionFields.ebofDeductFromEvent, vFields)
          .SetItem(EventBookingOptionFields.ebofMaximumBookings, vFields)
          .SetItem(EventBookingOptionFields.ebofProduct, vFields)
          .SetItem(EventBookingOptionFields.ebofRate, vFields)
          .SetItem(EventBookingOptionFields.ebofAmendedBy, vFields)
          .SetItem(EventBookingOptionFields.ebofAmendedOn, vFields)
          .SetItem(EventBookingOptionFields.ebofFreeOfCharge, vFields)
        End If
        .SetOptionalItem(EventBookingOptionFields.ebofLongDescription, vFields)
        .SetOptionalItem(EventBookingOptionFields.ebofMinimumBookings, vFields)
      End With
    End Sub

    Public Sub InitOptionSessions()
      Dim vOptionSession As New EventOptionSession
      Dim vRecordSet As CDBRecordSet

      mvOptionSessions = New Collection
      vOptionSession.Init(mvEnv)
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vOptionSession.GetRecordSetFields(EventOptionSession.EventOptionSessionRecordSetTypes.osrtAll) & " FROM option_sessions os WHERE option_number = " & mvClassFields(EventBookingOptionFields.ebofOptionNumber).IntegerValue & " ORDER BY session_number, option_number")
      Do While vRecordSet.Fetch() = True
        vOptionSession = New EventOptionSession
        vOptionSession.InitFromRecordSet(mvEnv, vRecordSet, EventOptionSession.EventOptionSessionRecordSetTypes.osrtAll)
        mvOptionSessions.Add(vOptionSession)
      Loop
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("option_number", CDBField.FieldTypes.cftLong, OptionNumber)
      mvEnv.Connection.StartTransaction()
      mvEnv.Connection.DeleteRecords("option_sessions", vWhereFields, False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
      mvEnv.Connection.CommitTransaction()
      InitClassFields()
      SetDefaults()
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(EventBookingOptionFields.ebofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pEvent As CDBEvent, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      With mvClassFields
        .Item(EventBookingOptionFields.ebofOptionNumber).IntegerValue = pEvent.AllocateNextNumber(CDBEvent.EventNumberFields.enfOptionNumber)
        .Item(EventBookingOptionFields.ebofEventNumber).Value = pParams("EventNumber").Value
      End With
      Update(pEvent, pParams)
    End Sub

    Public Sub Update(ByRef pEvent As CDBEvent, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("OptionDesc") Then .Item(EventBookingOptionFields.ebofOptionDesc).Value = pParams("OptionDesc").Value
        If pParams.Exists("PickSessions") Then .Item(EventBookingOptionFields.ebofPickSessions).Value = pParams("PickSessions").Value
        If pParams.Exists("NumberOfSessions") Then .Item(EventBookingOptionFields.ebofNumberOfSessions).Value = pParams("NumberOfSessions").Value
        If pParams.Exists("DeductFromEvent") Then .Item(EventBookingOptionFields.ebofDeductFromEvent).Value = pParams("DeductFromEvent").Value
        If pParams.Exists("MaximumBookings") Then .Item(EventBookingOptionFields.ebofMaximumBookings).Value = pParams("MaximumBookings").Value
        If pParams.Exists("MinimumBookings") Then .Item(EventBookingOptionFields.ebofMinimumBookings).Value = pParams("MinimumBookings").Value
        If pParams.Exists("Product") Then .Item(EventBookingOptionFields.ebofProduct).Value = pParams("Product").Value
        If pParams.Exists("Rate") Then .Item(EventBookingOptionFields.ebofRate).Value = pParams("Rate").Value
        If pParams.Exists("LongDescription") Then .Item(EventBookingOptionFields.ebofLongDescription).Value = pParams("LongDescription").Value
        If pParams.Exists("FreeOfCharge") Then .Item(EventBookingOptionFields.ebofFreeOfCharge).Value = pParams("FreeOfCharge").Value
      End With
      If Not pEvent.MultiSession Then
        If PickSessions Then RaiseError(DataAccessErrors.daeEventParameterError, "PickSessions")
        If NumberOfSessions > 1 Then RaiseError(DataAccessErrors.daeEventParameterError, "NumberOfSessions")
      End If
      If pEvent.EligibilityCheckRequired And MaximumBookings > 1 Then RaiseError(DataAccessErrors.daeEventParameterError, "MaximumBookings")
      If MinimumBookings > MaximumBookings Then RaiseError(DataAccessErrors.daeEventParameterError, "MinimumBookings")
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
        Return mvClassFields.Item(EventBookingOptionFields.ebofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property DeductFromEvent() As Boolean
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofDeductFromEvent).Bool
      End Get
    End Property

    Public Property EventNumber() As Integer
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofEventNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(EventBookingOptionFields.ebofEventNumber).Value = CStr(Value)
      End Set
    End Property

    Public Property IssueEventResources() As Boolean
      Get
        Dim vEvent As New CDBEvent(mvEnv)
        Dim vOptionSession As EventOptionSession
        vEvent.Init(EventNumber)

        IssueEventResources = False
        For Each vOptionSession In OptionSessions
          If vOptionSession.SessionNumber = vEvent.LowestSessionNumber Then
            IssueEventResources = True
            Exit For
          End If
        Next vOptionSession
      End Get
      Set(ByVal Value As Boolean)
        Dim vInsert As New CDBFields
        Dim vWhere As New CDBFields
        Dim vEvent As New CDBEvent(mvEnv)
        vEvent.Init(EventNumber)
        If Value Then
          If Not IssueEventResources Then
            vInsert.AddAmendedOnBy(mvEnv.User.Logname)
            vInsert.Add("option_number", CDBField.FieldTypes.cftLong, OptionNumber)
            vInsert.Add("session_number", CDBField.FieldTypes.cftLong, vEvent.LowestSessionNumber) ' Lowest session number
            vInsert.Add("allocation", CDBField.FieldTypes.cftNumeric, 0)
            mvEnv.Connection.InsertRecord("option_sessions", vInsert)
          End If
        Else
          'Don't issue resources - check if there was an entry
          If IssueEventResources Then
            vWhere.Add("option_number", CDBField.FieldTypes.cftLong, OptionNumber)
            vWhere.Add("session_number", CDBField.FieldTypes.cftLong, vEvent.LowestSessionNumber) ' Lowest session number
            mvEnv.Connection.DeleteRecords("option_sessions", vWhere, True)
          End If
        End If
      End Set
    End Property

    Public ReadOnly Property MaximumBookings() As Integer
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofMaximumBookings).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MinimumBookings() As Integer
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofMinimumBookings).IntegerValue
      End Get
    End Property

    Public ReadOnly Property NumberOfSessions() As Integer
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofNumberOfSessions).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OptionDesc() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofOptionDesc).Value
      End Get
    End Property

    Public ReadOnly Property OptionNumber() As Integer
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofOptionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property OptionSessions() As Collection
      Get
        If mvOptionSessions Is Nothing Then InitOptionSessions()
        OptionSessions = mvOptionSessions
      End Get
    End Property

    Public ReadOnly Property PickSessions() As Boolean
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofPickSessions).Bool
      End Get
    End Property

    Public ReadOnly Property ProductCode() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofProduct).Value
      End Get
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofRate).Value
      End Get
    End Property

    Public ReadOnly Property FreeOfCharge() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofFreeOfCharge).Value
      End Get
    End Property

    Public ReadOnly Property RateCurrencyCode() As String
      Get
        Return mvEnv.Connection.GetValue("SELECT currency_code FROM rates WHERE product = '" & ProductCode & "' AND rate = '" & RateCode & "'")
      End Get
    End Property

    Public ReadOnly Property ChangesAllowed() As Boolean
      Get
        Dim vWhereFields As CDBFields

        vWhereFields = New CDBFields
        vWhereFields.Add("option_number", CDBField.FieldTypes.cftLong, OptionNumber)
        If mvEnv.Connection.GetCount("event_bookings", vWhereFields) > 0 Then
          ChangesAllowed = False ' bookings exist
        Else
          ChangesAllowed = True
        End If
      End Get
    End Property
    Public ReadOnly Property LongDescription() As String
      Get
        Return mvClassFields.Item(EventBookingOptionFields.ebofLongDescription).Value
      End Get
    End Property
  End Class
End Namespace
