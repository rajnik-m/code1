

Namespace Access
  Public Class ExternalResource

    Public Enum ExternalResourceRecordSetTypes 'These are bit values
      exrrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ExternalResourceFields
      erfAll = 0
      erfResourceNumber
      erfEventNumber
      erfResourceDesc
      erfExternalResourceType
      erfOrganisationNumber
      erfAddressNumber
      erfContactNumber
      erfDepositAmount
      erfDepositDate
      erfFullAmount
      erfFullPaymentDate
      erfObtainedOn
      erfReturnBy
      erfReturnedOn
      erfAmendedBy
      erfAmendedOn
    End Enum

    Public Enum ExternalTypes
      etHire
      etPurchase
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
          .DatabaseTableName = "external_resources"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("resource_number", CDBField.FieldTypes.cftLong)
          .Add("event_number", CDBField.FieldTypes.cftLong)
          .Add("resource_desc")
          .Add("external_resource_type")
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("deposit_amount", CDBField.FieldTypes.cftNumeric)
          .Add("deposit_date", CDBField.FieldTypes.cftDate)
          .Add("full_amount", CDBField.FieldTypes.cftNumeric)
          .Add("full_payment_date", CDBField.FieldTypes.cftDate)
          .Add("obtained_on", CDBField.FieldTypes.cftDate)
          .Add("return_by", CDBField.FieldTypes.cftDate)
          .Add("returned_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(ExternalResourceFields.erfResourceNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ExternalResourceFields)
      'Add code here to ensure all values are valid before saving
      If ResourceNumber = 0 Then
        mvClassFields.Item(ExternalResourceFields.erfResourceNumber).Value = CStr(mvEnv.GetControlNumber("RN"))
      End If
      mvClassFields.Item(ExternalResourceFields.erfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ExternalResourceFields.erfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ExternalResourceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ExternalResourceRecordSetTypes.exrrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "er")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pResourceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pResourceNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ExternalResourceRecordSetTypes.exrrtAll) & " FROM external_resources er WHERE resource_number = " & pResourceNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ExternalResourceRecordSetTypes.exrrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ExternalResourceRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ExternalResourceFields.erfResourceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ExternalResourceRecordSetTypes.exrrtAll) = ExternalResourceRecordSetTypes.exrrtAll Then
          .SetItem(ExternalResourceFields.erfEventNumber, vFields)
          .SetItem(ExternalResourceFields.erfResourceDesc, vFields)
          .SetItem(ExternalResourceFields.erfExternalResourceType, vFields)
          .SetItem(ExternalResourceFields.erfOrganisationNumber, vFields)
          .SetItem(ExternalResourceFields.erfAddressNumber, vFields)
          .SetItem(ExternalResourceFields.erfContactNumber, vFields)
          .SetItem(ExternalResourceFields.erfDepositAmount, vFields)
          .SetItem(ExternalResourceFields.erfDepositDate, vFields)
          .SetItem(ExternalResourceFields.erfFullAmount, vFields)
          .SetItem(ExternalResourceFields.erfFullPaymentDate, vFields)
          .SetItem(ExternalResourceFields.erfObtainedOn, vFields)
          .SetItem(ExternalResourceFields.erfReturnBy, vFields)
          .SetItem(ExternalResourceFields.erfReturnedOn, vFields)
          .SetItem(ExternalResourceFields.erfAmendedBy, vFields)
          .SetItem(ExternalResourceFields.erfAmendedOn, vFields)
        End If
      End With
    End Sub
    Friend Sub InitFromExternalResource(ByRef pExternalResource As ExternalResource, ByRef pNewEvent As CDBEvent)

      With pExternalResource
        mvClassFields.Item(ExternalResourceFields.erfEventNumber).Value = CStr(pNewEvent.EventNumber)
        mvClassFields.Item(ExternalResourceFields.erfResourceDesc).Value = .ResourceDesc
        mvClassFields.Item(ExternalResourceFields.erfExternalResourceType).Value = .ExternalResourceType
        If .OrganisationNumber > 0 Then mvClassFields.Item(ExternalResourceFields.erfOrganisationNumber).Value = CStr(.OrganisationNumber)
        If .AddressNumber > 0 Then mvClassFields.Item(ExternalResourceFields.erfAddressNumber).Value = CStr(.AddressNumber)
        If .ContactNumber > 0 Then mvClassFields.Item(ExternalResourceFields.erfContactNumber).Value = CStr(.ContactNumber)
        mvClassFields.Item(ExternalResourceFields.erfDepositAmount).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfDepositDate).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfFullAmount).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfFullPaymentDate).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfObtainedOn).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfReturnBy).Value = ""
        mvClassFields.Item(ExternalResourceFields.erfReturnedOn).Value = ""
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      SetValid(ExternalResourceFields.erfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      Init(pEnv)
      mvClassFields.Item(ExternalResourceFields.erfEventNumber).Value = pParams("EventNumber").Value
      Update(pParams)
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services
      With mvClassFields
        If pParams.Exists("ResourceDesc") Then .Item(ExternalResourceFields.erfResourceDesc).Value = pParams("ResourceDesc").Value
        If pParams.Exists("OrganisationNumber") Then .Item(ExternalResourceFields.erfOrganisationNumber).Value = pParams("OrganisationNumber").Value
        If pParams.Exists("AddressNumber") Then .Item(ExternalResourceFields.erfAddressNumber).Value = pParams("AddressNumber").Value
        If pParams.Exists("ContactNumber") Then .Item(ExternalResourceFields.erfContactNumber).Value = pParams("ContactNumber").Value
        If pParams.Exists("ExternalResourceType") Then .Item(ExternalResourceFields.erfExternalResourceType).Value = pParams("ExternalResourceType").Value
        If pParams.Exists("ObtainedOn") Then .Item(ExternalResourceFields.erfObtainedOn).Value = pParams("ObtainedOn").Value
        If pParams.Exists("ReturnBy") Then .Item(ExternalResourceFields.erfReturnBy).Value = pParams("ReturnBy").Value
        If pParams.Exists("ReturnedOn") Then .Item(ExternalResourceFields.erfReturnedOn).Value = pParams("ReturnedOn").Value
        If pParams.Exists("Deposit") Then .Item(ExternalResourceFields.erfDepositAmount).Value = pParams("Deposit").Value
        If pParams.Exists("DepositPaidDate") Then .Item(ExternalResourceFields.erfDepositDate).Value = pParams("DepositPaidDate").Value
        If pParams.Exists("TotalAmount") Then .Item(ExternalResourceFields.erfFullAmount).Value = pParams("TotalAmount").Value
        If pParams.Exists("DueDate") Then .Item(ExternalResourceFields.erfFullPaymentDate).Value = pParams("DueDate").Value
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
        AddressNumber = mvClassFields.Item(ExternalResourceFields.erfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(ExternalResourceFields.erfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ExternalResourceFields.erfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ExternalResourceFields.erfContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExternalCodeFromType(ByVal pExternalType As ExternalTypes) As String
      Get
        Select Case pExternalType
          Case ExternalTypes.etHire
            ExternalCodeFromType = "H"
          Case ExternalTypes.etPurchase
            ExternalCodeFromType = "P"
          Case Else
            ExternalCodeFromType = ""
        End Select
      End Get
    End Property
    Public ReadOnly Property ExternalTypeFromCode(ByVal pExternalCode As String) As ExternalTypes
      Get
        Select Case pExternalCode
          Case "H"
            ExternalTypeFromCode = ExternalTypes.etHire
          Case "P"
            ExternalTypeFromCode = ExternalTypes.etPurchase
        End Select
      End Get
    End Property

    Public ReadOnly Property DepositAmount() As Double
      Get
        DepositAmount = mvClassFields.Item(ExternalResourceFields.erfDepositAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property DepositDate() As String
      Get
        DepositDate = mvClassFields.Item(ExternalResourceFields.erfDepositDate).Value
      End Get
    End Property

    Public Property EventNumber() As Integer
      Get
        EventNumber = mvClassFields.Item(ExternalResourceFields.erfEventNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(ExternalResourceFields.erfEventNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property ExternalResourceType() As String
      Get
        ExternalResourceType = mvClassFields.Item(ExternalResourceFields.erfExternalResourceType).Value
      End Get
    End Property

    Public ReadOnly Property FullAmount() As Double
      Get
        FullAmount = mvClassFields.Item(ExternalResourceFields.erfFullAmount).DoubleValue
      End Get
    End Property

    Public ReadOnly Property FullPaymentDate() As String
      Get
        FullPaymentDate = mvClassFields.Item(ExternalResourceFields.erfFullPaymentDate).Value
      End Get
    End Property

    Public ReadOnly Property ObtainedOn() As String
      Get
        ObtainedOn = mvClassFields.Item(ExternalResourceFields.erfObtainedOn).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(ExternalResourceFields.erfOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ResourceDesc() As String
      Get
        ResourceDesc = mvClassFields.Item(ExternalResourceFields.erfResourceDesc).Value
      End Get
    End Property

    Public Property ResourceNumber() As Integer
      Get
        ResourceNumber = mvClassFields.Item(ExternalResourceFields.erfResourceNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(ExternalResourceFields.erfResourceNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property ReturnBy() As String
      Get
        ReturnBy = mvClassFields.Item(ExternalResourceFields.erfReturnBy).Value
      End Get
    End Property

    Public ReadOnly Property ReturnedOn() As String
      Get
        ReturnedOn = mvClassFields.Item(ExternalResourceFields.erfReturnedOn).Value
      End Get
    End Property
    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub
  End Class
End Namespace
