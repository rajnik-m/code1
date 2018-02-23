

Namespace Access
  Public Class UnmannedCollection

    Public Enum UnmannedCollectionRecordSetTypes 'These are bit values
      ucrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum UnmannedCollectionFields
      ucfAll = 0
      ucfCollectionNumber
      ucfOrganisationNumber
      ucfAddressNumber
      ucfContactNumber
      ucfStartDate
      ucfEndDate
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvContact As Contact
    Private mvOrganisation As Organisation

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "unmanned_collections"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("end_date", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(UnmannedCollectionFields.ucfCollectionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(UnmannedCollectionFields.ucfCollectionNumber).PrefixRequired = True
        mvClassFields.Item(UnmannedCollectionFields.ucfOrganisationNumber).PrefixRequired = True
        mvClassFields.Item(UnmannedCollectionFields.ucfContactNumber).PrefixRequired = True
        mvClassFields.Item(UnmannedCollectionFields.ucfAddressNumber).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      'UPGRADE_NOTE: Object mvOrganisation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvOrganisation = Nothing
      'UPGRADE_NOTE: Object mvContact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mvContact = Nothing
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As UnmannedCollectionFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As UnmannedCollectionRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = UnmannedCollectionRecordSetTypes.ucrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "uc")
      End If
      GetRecordSetFields = vFields
    End Function

    Friend Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(UnmannedCollectionRecordSetTypes.ucrtAll) & " FROM unmanned_collections uc WHERE collection_number = " & pCollectionNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, UnmannedCollectionRecordSetTypes.ucrtAll)
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

    Friend Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As UnmannedCollectionRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(UnmannedCollectionFields.ucfCollectionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And UnmannedCollectionRecordSetTypes.ucrtAll) = UnmannedCollectionRecordSetTypes.ucrtAll Then
          .SetItem(UnmannedCollectionFields.ucfOrganisationNumber, vFields)
          .SetItem(UnmannedCollectionFields.ucfAddressNumber, vFields)
          .SetItem(UnmannedCollectionFields.ucfContactNumber, vFields)
          .SetItem(UnmannedCollectionFields.ucfStartDate, vFields)
          .SetItem(UnmannedCollectionFields.ucfEndDate, vFields)
        End If
      End With
    End Sub

    Friend Sub Save(Optional ByRef pAudit As Boolean = False, Optional ByRef pCollectionNumber As Integer = 0)
      SetValid(UnmannedCollectionFields.ucfAll)
      If pCollectionNumber > 0 Then mvClassFields(UnmannedCollectionFields.ucfCollectionNumber).IntegerValue = pCollectionNumber
      mvClassFields.Save(mvEnv, mvExisting, "", pAudit)
    End Sub

    'Friend Sub Create(ByVal pCollectionNumber As Long, ByVal pOrganisationNumber As Long, ByVal pAddressNumber As Long, ByVal pContactNumber As Long, ByVal pStartDate As String, ByVal pEndDate As String)
    '  With mvClassFields
    '    .Item(ucfCollectionNumber).IntegerValue = pCollectionNumber
    '    .Item(ucfOrganisationNumber).IntegerValue = pOrganisationNumber
    '    .Item(ucfAddressNumber).IntegerValue = pAddressNumber
    '    .Item(ucfContactNumber).IntegerValue = pContactNumber
    '    .Item(ucfStartDate).Value = pStartDate
    '    .Item(ucfEndDate).Value = pEndDate
    '  End With
    'End Sub
    '
    'Friend Sub Update(ByVal pOrganisationNumber As Long, ByVal pAddressNumber As Long, ByVal pContactNumber As Long, ByVal pStartDate As String, ByVal pEndDate As String)
    '  With mvClassFields
    '    .Item(ucfOrganisationNumber).IntegerValue = pOrganisationNumber
    '    .Item(ucfAddressNumber).IntegerValue = pAddressNumber
    '    .Item(ucfContactNumber).IntegerValue = pContactNumber
    '    .Item(ucfStartDate).Value = pStartDate
    '    .Item(ucfEndDate).Value = pEndDate
    '  End With
    'End Sub

    Friend Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcUnmannedCollection As UnmannedCollection)
      Init(pEnv)

      With mvClassFields
        .Item(UnmannedCollectionFields.ucfOrganisationNumber).Value = CStr(pSrcUnmannedCollection.OrganisationNumber)
        .Item(UnmannedCollectionFields.ucfAddressNumber).Value = CStr(pSrcUnmannedCollection.AddressNumber)
        .Item(UnmannedCollectionFields.ucfContactNumber).Value = CStr(pSrcUnmannedCollection.ContactNumber)
        .Item(UnmannedCollectionFields.ucfStartDate).Value = pSrcUnmannedCollection.StartDate
        .Item(UnmannedCollectionFields.ucfEndDate).Value = pSrcUnmannedCollection.EndDate
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
        Return mvClassFields.Item(UnmannedCollectionFields.ucfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(UnmannedCollectionFields.ucfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Contact() As Contact
      Get
        If mvContact Is Nothing Then
          mvContact = New Contact(mvEnv)
          mvContact.Init((mvClassFields.Item(UnmannedCollectionFields.ucfContactNumber).IntegerValue))
        End If
        Contact = mvContact
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields.Item(UnmannedCollectionFields.ucfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property EndDate() As String
      Get
        EndDate = mvClassFields.Item(UnmannedCollectionFields.ucfEndDate).Value
      End Get
    End Property

    Public ReadOnly Property Organisation() As Organisation
      Get
        'Init the Organisation for the OrganisationNumber and AddressNumber
        If mvOrganisation Is Nothing Then
          mvOrganisation = New Organisation(mvEnv)
          mvOrganisation.InitWithAddress(mvEnv, (mvClassFields.Item(UnmannedCollectionFields.ucfOrganisationNumber).IntegerValue), (mvClassFields.Item(UnmannedCollectionFields.ucfAddressNumber).IntegerValue))
        End If
        Organisation = mvOrganisation
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        Return mvClassFields.Item(UnmannedCollectionFields.ucfOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property StartDate() As String
      Get
        StartDate = mvClassFields.Item(UnmannedCollectionFields.ucfStartDate).Value
      End Get
    End Property

    Friend Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(UnmannedCollectionFields.ucfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        .Item(UnmannedCollectionFields.ucfAddressNumber).Value = pParams("AddressNumber").Value
        .Item(UnmannedCollectionFields.ucfContactNumber).Value = pParams("ContactNumber").Value
        .Item(UnmannedCollectionFields.ucfStartDate).Value = pParams("StartDate").Value
        .Item(UnmannedCollectionFields.ucfEndDate).Value = pParams("EndDate").Value
      End With
    End Sub

    Friend Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("OrganisationNumber") Then .Item(UnmannedCollectionFields.ucfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(UnmannedCollectionFields.ucfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("ContactNumber") Then .Item(UnmannedCollectionFields.ucfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("StartDate") Then .Item(UnmannedCollectionFields.ucfStartDate).Value = pParams("StartDate").Value
        If pParams.Exists("EndDate") Then .Item(UnmannedCollectionFields.ucfEndDate).Value = pParams("EndDate").Value
      End With
    End Sub
  End Class
End Namespace
