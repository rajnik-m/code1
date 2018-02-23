

Namespace Access
  Public Class SelectionSet

    Public Enum SelectionSetRecordSetTypes 'These are bit values
      sstrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum SelectionSetFields
      ssfAll = 0
      ssfSelectionSet
      ssfUserName
      ssfDepartment
      ssfSelectionSetDesc
      ssfNumberInSet
      ssfSelectionGroup
      ssfCustomData
      ssfAttributeCaptions
      ssfShowOwner
      ssfShowPosition
      ssfCheckBoxCaptions
      ssfSource
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
          .DatabaseTableName = "selection_sets"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("selection_set", CDBField.FieldTypes.cftLong)
          .Add("user_name")
          .Add("department")
          .Add("selection_set_desc")
          .Add("number_in_set", CDBField.FieldTypes.cftLong)
          .Add("selection_group")
          .Add("custom_data")
          .Add("attribute_captions")
          .Add("show_owner")
          .Add("show_position")
          .Add("check_box_captions")
          .Add("source")

          .Item(SelectionSetFields.ssfSelectionSet).SetPrimaryKeyOnly()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As SelectionSetFields)
      'Add code here to ensure all values are valid before saving
      If SelectionSetNumber = 0 Then mvClassFields(SelectionSetFields.ssfSelectionSet).IntegerValue = mvEnv.GetControlNumber("SS")
      If Len(Department) = 0 Then mvClassFields(SelectionSetFields.ssfDepartment).Value = mvEnv.User.Department
      If Len(UserName) = 0 Then mvClassFields(SelectionSetFields.ssfUserName).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As SelectionSetRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = SelectionSetRecordSetTypes.sstrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ss")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pSelectionSet As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pSelectionSet > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(SelectionSetRecordSetTypes.sstrtAll) & " FROM selection_sets ss WHERE selection_set = " & pSelectionSet)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, SelectionSetRecordSetTypes.sstrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
        mvClassFields(SelectionSetFields.ssfSelectionSet).Value = pSelectionSet.ToString
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As SelectionSetRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(SelectionSetFields.ssfSelectionSet, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And SelectionSetRecordSetTypes.sstrtAll) = SelectionSetRecordSetTypes.sstrtAll Then
          .SetItem(SelectionSetFields.ssfUserName, vFields)
          .SetItem(SelectionSetFields.ssfDepartment, vFields)
          .SetItem(SelectionSetFields.ssfSelectionSetDesc, vFields)
          .SetItem(SelectionSetFields.ssfNumberInSet, vFields)
          .SetItem(SelectionSetFields.ssfSelectionGroup, vFields)
          .SetItem(SelectionSetFields.ssfCustomData, vFields)
          .SetItem(SelectionSetFields.ssfAttributeCaptions, vFields)
          .SetItem(SelectionSetFields.ssfShowOwner, vFields)
          .SetItem(SelectionSetFields.ssfShowPosition, vFields)
          .SetItem(SelectionSetFields.ssfCheckBoxCaptions, vFields)
          .SetItem(SelectionSetFields.ssfSource, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(SelectionSetFields.ssfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pDescription As String, ByRef pNumberInSet As Integer, ByRef pGroup As String, Optional ByRef pSSNumber As Integer = 0, Optional ByRef pOwner As String = "", Optional ByRef pSourceCode As String = "", Optional ByRef pDepartment As String = "")
      With mvClassFields
        .Item(SelectionSetFields.ssfSelectionSetDesc).Value = pDescription
        .Item(SelectionSetFields.ssfNumberInSet).IntegerValue = pNumberInSet
        .Item(SelectionSetFields.ssfSelectionGroup).Value = pGroup
        If pSSNumber > 0 Then .Item(SelectionSetFields.ssfSelectionSet).Value = CStr(pSSNumber)
        If Len(pOwner) > 0 Then .Item(SelectionSetFields.ssfUserName).Value = pOwner
        If Len(pDepartment) > 0 Then .Item(SelectionSetFields.ssfDepartment).Value = pDepartment
        If Len(pSourceCode) > 0 Then .Item(SelectionSetFields.ssfSource).Value = pSourceCode
      End With
      SetValid(SelectionSetFields.ssfAll)
    End Sub

    'UPGRADE_NOTE: Rename was upgraded to Rename_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Sub Rename_Renamed(ByRef pNewDescription As String)
      mvClassFields.Item(SelectionSetFields.ssfSelectionSetDesc).Value = pNewDescription
    End Sub

    Public Sub CopyData(ByRef pSetNo As Integer, ByRef pGroup As String)
      Dim vSQL As String
      Dim vWhereFields As New CDBFields
      'First insert the new rows into the temporary table if required
      'This is because we cannot do an insert into select from on the same table
      If pGroup = "GM" Then
        vSQL = "INSERT INTO selected_contacts_temp (selection_set,revision,contact_number,address_number) SELECT selection_set,1,contact_number,address_number FROM selected_contacts WHERE selection_set = " & pSetNo
        mvEnv.Connection.ExecuteSQL(vSQL)
      End If
      vSQL = "INSERT INTO selected_contacts (selection_set,revision,contact_number,address_number) SELECT " & SelectionSetNumber & ",1,contact_number,address_number FROM selected_contacts_temp WHERE selection_set = " & pSetNo
      mvClassFields.Item(SelectionSetFields.ssfNumberInSet).IntegerValue = mvEnv.Connection.ExecuteSQL(vSQL) 'Insert the rows into the real destination table
      If pGroup = "GM" Then 'Now delete the rows from the temporary table
        vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSetNo)
        mvEnv.Connection.DeleteRecords("selected_contacts_temp", vWhereFields, False)
      End If
    End Sub

    Public Sub Delete()
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, SelectionSetNumber)
      With mvEnv.Connection
        .DeleteRecords("selected_contacts", vWhereFields, False)
        .DeleteRecords("selected_contacts_temp", vWhereFields, False)
        .DeleteRecords("selection_set_data", vWhereFields, False)
      End With
      mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Sub DeleteCustomData()
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, SelectionSetNumber)
      mvEnv.Connection.DeleteRecords("selection_set_data", vWhereFields, False)
    End Sub

    Public Sub ChangeSelectionGroup(ByRef pNewGroup As String)
      Dim vSource As String
      Dim vDest As String
      Dim vSQL As String
      Dim vWhereFields As New CDBFields

      If pNewGroup <> SelectionGroup Then
        vSource = "selected_contacts"
        vDest = "selected_contacts"
        If pNewGroup = "AU" Then
          vDest = vDest & "_temp"
        Else
          vSource = vSource & "_temp"
        End If
        vSQL = "INSERT INTO " & vDest & " (selection_set,revision,contact_number,address_number) SELECT selection_set,1,contact_number,address_number FROM " & vSource & " WHERE selection_set = " & SelectionSetNumber
        mvEnv.Connection.ExecuteSQL(vSQL) 'Insert the rows into the destination table
        vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, SelectionSetNumber) 'Now delete the rows from the source table
        mvEnv.Connection.DeleteRecords(vSource, vWhereFields, False)
        mvClassFields(SelectionSetFields.ssfSelectionGroup).Value = pNewGroup
        Save()
      End If
    End Sub

    Public Sub SetCustomData(ByRef pCustomDataSet As CustomDataSet, ByRef pCaptions As String)
      mvClassFields.Item(SelectionSetFields.ssfCustomData).Bool = True
      mvClassFields.Item(SelectionSetFields.ssfAttributeCaptions).Value = pCaptions
      mvClassFields.Item(SelectionSetFields.ssfShowOwner).Bool = pCustomDataSet.ShowOwner
      mvClassFields.Item(SelectionSetFields.ssfShowPosition).Bool = pCustomDataSet.ShowPosition
      mvClassFields.Item(SelectionSetFields.ssfCheckBoxCaptions).Value = pCustomDataSet.CheckBoxCaptions
    End Sub

    Private ReadOnly Property ContactsTableName() As String
      Get
        If Existing Then
          If SelectionGroup = "AU" Then
            Return "selected_contacts_temp"
          Else
            Return "selected_contacts"
          End If
        Else
          Return "smcam_smapp_" & SelectionSetNumber.ToString
        End If
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AttributeCaptions() As String
      Get
        AttributeCaptions = mvClassFields.Item(SelectionSetFields.ssfAttributeCaptions).Value
      End Get
    End Property

    Public ReadOnly Property CheckBoxCaptions() As String
      Get
        CheckBoxCaptions = mvClassFields.Item(SelectionSetFields.ssfCheckBoxCaptions).Value
      End Get
    End Property

    Public ReadOnly Property CustomData() As Boolean
      Get
        CustomData = mvClassFields.Item(SelectionSetFields.ssfCustomData).Bool
      End Get
    End Property

    Public ReadOnly Property Department() As String
      Get
        Department = mvClassFields.Item(SelectionSetFields.ssfDepartment).Value
      End Get
    End Property

    Public Property NumberInSet() As Integer
      Get
        NumberInSet = mvClassFields.Item(SelectionSetFields.ssfNumberInSet).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(SelectionSetFields.ssfNumberInSet).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property SelectionGroup() As String
      Get
        SelectionGroup = mvClassFields.Item(SelectionSetFields.ssfSelectionGroup).Value
      End Get
    End Property

    Public ReadOnly Property SelectionSetNumber() As Integer
      Get
        SelectionSetNumber = mvClassFields.Item(SelectionSetFields.ssfSelectionSet).IntegerValue
      End Get
    End Property

    Public ReadOnly Property SelectionSetDesc() As String
      Get
        SelectionSetDesc = mvClassFields.Item(SelectionSetFields.ssfSelectionSetDesc).Value
      End Get
    End Property

    Public ReadOnly Property ShowOwner() As Boolean
      Get
        ShowOwner = mvClassFields.Item(SelectionSetFields.ssfShowOwner).Bool
      End Get
    End Property

    Public ReadOnly Property ShowPosition() As Boolean
      Get
        ShowPosition = mvClassFields.Item(SelectionSetFields.ssfShowPosition).Bool
      End Get
    End Property

    Public ReadOnly Property Source() As String
      Get
        Source = mvClassFields.Item(SelectionSetFields.ssfSource).Value
      End Get
    End Property

    Public ReadOnly Property UserName() As String
      Get
        UserName = mvClassFields.Item(SelectionSetFields.ssfUserName).Value
      End Get
    End Property

    Public Sub AddContact(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer)
      Dim vFields As New CDBFields

      vFields.Add("selection_set", CDBField.FieldTypes.cftLong, SelectionSetNumber)
      vFields.Add("revision", CDBField.FieldTypes.cftLong, 1)
      vFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      vFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
      mvEnv.Connection.InsertRecord(ContactsTableName, vFields, True)
      If Not mvEnv.Connection.IsLastErrorDuplicate() Then
        mvClassFields(SelectionSetFields.ssfNumberInSet).IntegerValue = NumberInSet + 1
        'Do this via direct SQL in case of multi-user changes
        mvEnv.Connection.ExecuteSQL("UPDATE selection_sets SET number_in_set = number_in_set + 1 WHERE selection_set = " & SelectionSetNumber)
      End If
    End Sub

    Public Function DeleteContact(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer) As Boolean
      Dim vFields As New CDBFields

      vFields.Add("selection_set", CDBField.FieldTypes.cftLong, SelectionSetNumber)
      vFields.Add("revision", CDBField.FieldTypes.cftLong, 1)
      vFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      If pAddressNumber > 0 Then vFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
      If mvEnv.Connection.DeleteRecords(ContactsTableName, vFields, False) > 0 Then
        If NumberInSet > 0 Then
          mvClassFields(SelectionSetFields.ssfNumberInSet).IntegerValue = NumberInSet - 1
          'Do this via direct SQL in case of multi-user changes
          mvEnv.Connection.ExecuteSQL("UPDATE selection_sets SET number_in_set = number_in_set - 1 WHERE selection_set = " & SelectionSetNumber)
        End If
        Return True
      End If
    End Function

    Public Function TemporaryTableExists(ByVal pSelectionSetNumber As Integer) As Boolean
      mvClassFields(SelectionSetFields.ssfSelectionSet).Value = pSelectionSetNumber.ToString
      Return mvEnv.Connection.TableExists(ContactsTableName)
    End Function
  End Class
End Namespace
