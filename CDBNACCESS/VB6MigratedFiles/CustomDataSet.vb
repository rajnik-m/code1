

Namespace Access
  Public Class CustomDataSet

    Public Enum CustomDataSetRecordSetTypes 'These are bit values
      cdsrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CustomDataSetFields
      cdsfAll = 0
      cdsfCustomDataSet
      cdsfCustomDataSetDesc
      cdsfClient
      cdsfShowOwner
      cdsfShowPosition
      cdsfOrganisationTelephone
      cdsfCheckBoxCaptions
      cdsfAmendedBy
      cdsfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvDetails As Collection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "custom_data_sets"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("custom_data_set")
          .Add("custom_data_set_desc")
          .Add("client")
          .Add("show_owner")
          .Add("show_position")
          .Add("organisation_telephone")
          .Add("check_box_captions")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CustomDataSetFields.cdsfCustomDataSet).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvDetails = Nothing
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CustomDataSetFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(CustomDataSetFields.cdsfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CustomDataSetFields.cdsfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CustomDataSetRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CustomDataSetRecordSetTypes.cdsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cds")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCustomDataSet As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pCustomDataSet) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CustomDataSetRecordSetTypes.cdsrtAll) & " FROM custom_data_sets cds WHERE custom_data_set = '" & pCustomDataSet & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CustomDataSetRecordSetTypes.cdsrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CustomDataSetRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CustomDataSetFields.cdsfCustomDataSet, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CustomDataSetRecordSetTypes.cdsrtAll) = CustomDataSetRecordSetTypes.cdsrtAll Then
          .SetItem(CustomDataSetFields.cdsfCustomDataSetDesc, vFields)
          .SetItem(CustomDataSetFields.cdsfClient, vFields)
          .SetItem(CustomDataSetFields.cdsfShowOwner, vFields)
          .SetItem(CustomDataSetFields.cdsfShowPosition, vFields)
          .SetItem(CustomDataSetFields.cdsfOrganisationTelephone, vFields)
          .SetItem(CustomDataSetFields.cdsfCheckBoxCaptions, vFields)
          .SetItem(CustomDataSetFields.cdsfAmendedBy, vFields)
          .SetItem(CustomDataSetFields.cdsfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CustomDataSetFields.cdsfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public ReadOnly Property Details() As Collection
      Get
        Dim vCDSD As New CustomDataSetDetail
        Dim vRS As CDBRecordSet

        If mvDetails Is Nothing Then
          mvDetails = New Collection
          vCDSD.InitDefault(mvEnv)
          mvDetails.Add(vCDSD, "0")
          vRS = mvEnv.Connection.GetRecordSet("SELECT " & vCDSD.GetRecordSetFields(CustomDataSetDetail.CustomDataSetDetailRecordSetTypes.cdsdrtAll) & " FROM custom_data_set_details WHERE custom_data_set = '" & CustomDataSetCode & "' ORDER BY sequence_number")
          While vRS.Fetch() = True
            vCDSD = New CustomDataSetDetail
            vCDSD.InitFromRecordSet(mvEnv, vRS, CustomDataSetDetail.CustomDataSetDetailRecordSetTypes.cdsdrtAll)
            mvDetails.Add(vCDSD, CStr(vCDSD.SequenceNumber))
          End While
          vRS.CloseRecordSet()
        End If
        Details = mvDetails
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

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CustomDataSetFields.cdsfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CustomDataSetFields.cdsfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CheckBoxCaptions() As String
      Get
        CheckBoxCaptions = mvClassFields.Item(CustomDataSetFields.cdsfCheckBoxCaptions).Value
      End Get
    End Property

    Public ReadOnly Property ClientCode() As String
      Get
        ClientCode = mvClassFields.Item(CustomDataSetFields.cdsfClient).Value
      End Get
    End Property

    Public ReadOnly Property CustomDataSetCode() As String
      Get
        CustomDataSetCode = mvClassFields.Item(CustomDataSetFields.cdsfCustomDataSet).Value
      End Get
    End Property

    Public ReadOnly Property CustomDataSetDesc() As String
      Get
        CustomDataSetDesc = mvClassFields.Item(CustomDataSetFields.cdsfCustomDataSetDesc).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationTelephone() As Boolean
      Get
        OrganisationTelephone = mvClassFields.Item(CustomDataSetFields.cdsfOrganisationTelephone).Bool
      End Get
    End Property

    Public ReadOnly Property ShowOwner() As Boolean
      Get
        ShowOwner = mvClassFields.Item(CustomDataSetFields.cdsfShowOwner).Bool
      End Get
    End Property

    Public ReadOnly Property ShowPosition() As Boolean
      Get
        ShowPosition = mvClassFields.Item(CustomDataSetFields.cdsfShowPosition).Bool
      End Get
    End Property
  End Class
End Namespace
