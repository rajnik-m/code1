

Namespace Access
  Public Class CollectionResource

    Public Enum CollectionResourceRecordSetTypes 'These are bit values
      crsrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionResourceFields
      crfAll = 0
      crfCollectionResourceNumber
      crfCollectionNumber
      crfAppealResourceNumber
      crfRate
      crfQuantity
      crfDespatchOn
      crfDespatchMethod
      crfAmendedBy
      crfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvAppealCollection As AppealCollection

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "collection_resources"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_resource_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("appeal_resource_number", CDBField.FieldTypes.cftLong)
          .Add("rate")
          .Add("quantity", CDBField.FieldTypes.cftLong)
          .Add("despatch_on", CDBField.FieldTypes.cftDate)
          .Add("despatch_method")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CollectionResourceFields.crfCollectionResourceNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionResourceFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(CollectionResourceFields.crfCollectionResourceNumber).IntegerValue = 0 Then mvClassFields.Item(CollectionResourceFields.crfCollectionResourceNumber).IntegerValue = mvEnv.GetControlNumber("CU")
      If mvClassFields.FieldsChanged And mvExisting Then
        If IsDate(AppealCollection.ResourcesProducedOn) Then
          RaiseError(DataAccessErrors.daeCannotChangeCollResAsFulfilled)
        End If
      End If
      mvClassFields.Item(CollectionResourceFields.crfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectionResourceFields.crfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionResourceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CollectionResourceRecordSetTypes.crsrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cr")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionResourceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionResourceNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionResourceRecordSetTypes.crsrtAll) & " FROM collection_resources cr WHERE collection_resource_number = " & pCollectionResourceNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionResourceRecordSetTypes.crsrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionResourceRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionResourceFields.crfCollectionResourceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionResourceRecordSetTypes.crsrtAll) = CollectionResourceRecordSetTypes.crsrtAll Then
          .SetItem(CollectionResourceFields.crfCollectionNumber, vFields)
          .SetItem(CollectionResourceFields.crfAppealResourceNumber, vFields)
          .SetItem(CollectionResourceFields.crfRate, vFields)
          .SetItem(CollectionResourceFields.crfQuantity, vFields)
          .SetItem(CollectionResourceFields.crfDespatchOn, vFields)
          .SetItem(CollectionResourceFields.crfDespatchMethod, vFields)
          .SetItem(CollectionResourceFields.crfAmendedBy, vFields)
          .SetItem(CollectionResourceFields.crfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vQuantityIssued As Integer
      Dim vAppResource As New AppealResource
      Dim vSaveAppealResource As Boolean

      SetValid(CollectionResourceFields.crfAll)
      If mvClassFields("quantity").ValueChanged Then
        vQuantityIssued = Quantity - IntegerValue(mvClassFields(CollectionResourceFields.crfQuantity).SetValue)
        vAppResource.Init(mvEnv, AppealResourceNumber)
        vAppResource.IssueQuantity(vQuantityIssued)
        vSaveAppealResource = True
        mvEnv.Connection.StartTransaction()
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      If vSaveAppealResource Then
        vAppResource.Save(pAmendedBy, pAudit)
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(CollectionResourceFields.crfCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(CollectionResourceFields.crfAppealResourceNumber).IntegerValue = pParams("AppealResourceNumber").IntegerValue
        Update(pParams)
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("Rate") Then .Item(CollectionResourceFields.crfRate).Value = pParams("Rate").Value
        If pParams.Exists("Quantity") Then .Item(CollectionResourceFields.crfQuantity).Value = pParams("Quantity").Value
        If pParams.Exists("DespatchOn") Then .Item(CollectionResourceFields.crfDespatchOn).Value = pParams("DespatchOn").Value
        If pParams.Exists("DespatchMethod") Then .Item(CollectionResourceFields.crfDespatchMethod).Value = pParams("DespatchMethod").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vAppResource As New AppealResource
      If DeleteAllowed() Then
        vAppResource.Init(mvEnv, AppealResourceNumber)
        vAppResource.IssueQuantity(Quantity * -1)
        mvEnv.Connection.StartTransaction()
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
        vAppResource.Save()
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Private Function DeleteAllowed() As Boolean
      Dim vDeleteAllowed As Boolean

      vDeleteAllowed = True
      If IsDate(AppealCollection.ResourcesProducedOn) Then
        RaiseError(DataAccessErrors.daeCannotDeleteCollResourceAsFulfilled)
      End If
      DeleteAllowed = vDeleteAllowed
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Private ReadOnly Property AppealCollection() As AppealCollection
      Get
        If mvAppealCollection Is Nothing Then
          mvAppealCollection = New AppealCollection(mvEnv)
          mvAppealCollection.Init(CollectionNumber)
        End If
        AppealCollection = mvAppealCollection

      End Get
    End Property
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CollectionResourceFields.crfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionResourceFields.crfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AppealResourceNumber() As Integer
      Get
        AppealResourceNumber = mvClassFields.Item(CollectionResourceFields.crfAppealResourceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = CInt(mvClassFields.Item(CollectionResourceFields.crfCollectionNumber).Value)
      End Get
    End Property

    Public ReadOnly Property CollectionResourceNumber() As Integer
      Get
        CollectionResourceNumber = CInt(mvClassFields.Item(CollectionResourceFields.crfCollectionResourceNumber).Value)
      End Get
    End Property

    Public ReadOnly Property DespatchMethod() As String
      Get
        DespatchMethod = mvClassFields.Item(CollectionResourceFields.crfDespatchMethod).Value
      End Get
    End Property

    Public ReadOnly Property DespatchOn() As String
      Get
        DespatchOn = mvClassFields.Item(CollectionResourceFields.crfDespatchOn).Value
      End Get
    End Property

    Public ReadOnly Property Quantity() As Integer
      Get
        Quantity = mvClassFields.Item(CollectionResourceFields.crfQuantity).IntegerValue
      End Get
    End Property

    Public ReadOnly Property RateCode() As String
      Get
        RateCode = mvClassFields.Item(CollectionResourceFields.crfRate).Value
      End Get
    End Property
  End Class
End Namespace
