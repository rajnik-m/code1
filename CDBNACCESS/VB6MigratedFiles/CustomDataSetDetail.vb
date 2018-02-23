

Namespace Access
  Public Class CustomDataSetDetail

    Public Enum CustomDataSetDetailRecordSetTypes 'These are bit values
      cdsdrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CustomDataSetDetailFields
      cdsdfAll = 0
      cdsdfCustomDataSet
      cdsdfSequenceNumber
      cdsdfDbName
      cdsdfSelectSql
      cdsdfAttributeNames
      cdsdfAttributeCaptions
      cdsdfAmendedBy
      cdsdfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvRecordSet As CDBRecordSet
    Private mvContactNumber As Integer
    Private mvStatus As Boolean
    Private mvSelectedAttributes As String
    Private mvSelectedCaptions As String

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "custom_data_set_details"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("custom_data_set")
          .Add("sequence_number", CDBField.FieldTypes.cftInteger)
          .Add("db_name")
          .Add("select_sql")
          .Add("attribute_names")
          .Add("attribute_captions")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CustomDataSetDetailFields.cdsdfCustomDataSet).SetPrimaryKeyOnly()
        mvClassFields.Item(CustomDataSetDetailFields.cdsdfSequenceNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvStatus = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CustomDataSetDetailFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CustomDataSetDetailRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CustomDataSetDetailRecordSetTypes.cdsdrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cdsd")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCustomDataSet As String = "", Optional ByRef pSequenceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCustomDataSet) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CustomDataSetDetailRecordSetTypes.cdsdrtAll) & " FROM custom_data_set_details cdsd WHERE custom_data_set = '" & pCustomDataSet & "' AND sequence_number = " & pSequenceNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CustomDataSetDetailRecordSetTypes.cdsdrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CustomDataSetDetailRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CustomDataSetDetailFields.cdsdfCustomDataSet, vFields)
        .SetItem(CustomDataSetDetailFields.cdsdfSequenceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CustomDataSetDetailRecordSetTypes.cdsdrtAll) = CustomDataSetDetailRecordSetTypes.cdsdrtAll Then
          .SetItem(CustomDataSetDetailFields.cdsdfDbName, vFields)
          .SetItem(CustomDataSetDetailFields.cdsdfSelectSql, vFields)
          .SetItem(CustomDataSetDetailFields.cdsdfAttributeNames, vFields)
          .SetItem(CustomDataSetDetailFields.cdsdfAttributeCaptions, vFields)
          .SetItem(CustomDataSetDetailFields.cdsdfAmendedBy, vFields)
          .SetItem(CustomDataSetDetailFields.cdsdfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CustomDataSetDetailFields.cdsdfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub InitDefault(ByVal pEnv As CDBEnvironment)
      Init(pEnv)
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfDbName).Value = "DATA"
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfSelectSql).Value = "SELECT sc.contact_number, position, name, dialling_code, std_code, telephone FROM selected_contacts sc, contact_positions cp, organisations o WHERE sc.selection_set = # AND sc.contact_number = cp.contact_number AND sc.address_number = cp.address_number AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND cp.organisation_number = o.organisation_number ORDER BY sc.contact_number"
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfAttributeNames).Value = "position,name"
      mvClassFields.Item(CustomDataSetDetailFields.cdsdfAttributeCaptions).Value = "Position,Organisation"
    End Sub

    Public Sub SetSelectionSet(ByRef pSelectionSetNumber As Integer)
      Dim vSQL As String

      vSQL = SelectSql
      If InStr(vSQL, "#") > 0 Then mvClassFields(CustomDataSetDetailFields.cdsdfSelectSql).Value = Replace(vSQL, "#", CStr(pSelectionSetNumber))
    End Sub

    Public Sub SelectData(ByRef pSelect As Boolean)
      If pSelect Then
        mvRecordSet = mvEnv.GetConnection(DBName).GetRecordSet(SelectSql)
        If mvRecordSet.Fetch() = True Then
          mvContactNumber = mvRecordSet.Fields("contact_number").IntegerValue
          If mvContactNumber <= 0 Then
            mvStatus = False
          Else
            mvStatus = True
          End If
        Else
          mvStatus = False
        End If
      Else
        mvStatus = False
      End If
    End Sub

    Public Sub GetNext()
      If mvStatus = True Then
        mvStatus = mvRecordSet.Fetch
        If mvStatus = True Then mvContactNumber = mvRecordSet.Fields("contact_number").IntegerValue
      End If
    End Sub

    Public Sub CloseRecordSet()
      If Not mvRecordSet Is Nothing Then
        mvRecordSet.CloseRecordSet()
        mvRecordSet = Nothing
      End If
    End Sub

    Public Function GetValue(ByRef pAttr As String) As String
      Dim vPhoneNumber As String = ""
      If pAttr = "PHONE" Then
        If Len(mvRecordSet.Fields("dialling_code").Value) > 0 Then vPhoneNumber = "(" & mvRecordSet.Fields("dialling_code").Value & ") "
        If Len(mvRecordSet.Fields("std_code").Value) > 0 Then vPhoneNumber = vPhoneNumber & mvRecordSet.Fields("std_code").Value & " "
        GetValue = vPhoneNumber & mvRecordSet.Fields("telephone").Value
      Else
        GetValue = mvRecordSet.Fields(pAttr).Value
      End If
    End Function

    Public ReadOnly Property Status() As Boolean
      Get
        Status = mvStatus
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvContactNumber
      End Get
    End Property

    Public Property SelectedAttributes() As String
      Get
        SelectedAttributes = mvSelectedAttributes
      End Get
      Set(ByVal Value As String)
        mvSelectedAttributes = Value
      End Set
    End Property

    Public Property SelectedCaptions() As String
      Get
        SelectedCaptions = mvSelectedCaptions
      End Get
      Set(ByVal Value As String)
        mvSelectedCaptions = Value
      End Set
    End Property

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
        AmendedBy = mvClassFields.Item(CustomDataSetDetailFields.cdsdfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CustomDataSetDetailFields.cdsdfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AttributeCaptions() As String
      Get
        AttributeCaptions = mvClassFields.Item(CustomDataSetDetailFields.cdsdfAttributeCaptions).Value
      End Get
    End Property

    Public ReadOnly Property AttributeNames() As String
      Get
        AttributeNames = mvClassFields.Item(CustomDataSetDetailFields.cdsdfAttributeNames).Value
      End Get
    End Property

    Public ReadOnly Property CustomDataSetCode() As String
      Get
        CustomDataSetCode = mvClassFields.Item(CustomDataSetDetailFields.cdsdfCustomDataSet).Value
      End Get
    End Property

    Public ReadOnly Property DBName() As String
      Get
        DBName = mvClassFields.Item(CustomDataSetDetailFields.cdsdfDbName).Value
      End Get
    End Property

    Public ReadOnly Property SelectSql() As String
      Get
        SelectSql = mvClassFields.Item(CustomDataSetDetailFields.cdsdfSelectSql).Value
      End Get
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvClassFields.Item(CustomDataSetDetailFields.cdsdfSequenceNumber).IntegerValue
      End Get
    End Property
  End Class
End Namespace
