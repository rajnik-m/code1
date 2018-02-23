Namespace Access
  Public Class AddressData

    Public Enum AddressDataRecordSetTypes 'These are bit values
      adrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum AddressDataFields
      adfAll = 0
      adfAddressNumber
      adfLEACode
      adfLEADesc
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
          .DatabaseTableName = "address_data"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("lea_code")
          .Add("lea_name")
        End With
        mvClassFields.Item(AddressDataFields.adfAddressNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As AddressDataFields)
      'Add code here to ensure all values are valid before saving
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As AddressDataRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AddressDataRecordSetTypes.adrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ad")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAddressNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pAddressNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AddressDataRecordSetTypes.adrtAll) & " FROM address_data WHERE address_number = " & pAddressNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, AddressDataRecordSetTypes.adrtAll)
        Else
          InitClassFields()
          SetDefaults()
          mvClassFields(AddressDataFields.adfAddressNumber).IntegerValue = pAddressNumber
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AddressDataRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AddressDataFields.adfAddressNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AddressDataRecordSetTypes.adrtAll) = AddressDataRecordSetTypes.adrtAll Then
          .SetItem(AddressDataFields.adfLEACode, vFields)
          .SetItem(AddressDataFields.adfLEADesc, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(AddressDataFields.adfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pAddressNumber As Integer)
      mvClassFields.Item(AddressDataFields.adfAddressNumber).Value = CStr(pAddressNumber)
    End Sub

    Public Sub SetLEAData(ByRef pLEACode As String, ByRef pLEAName As String)
      mvClassFields.Item(AddressDataFields.adfLEACode).Value = pLEACode
      mvClassFields.Item(AddressDataFields.adfLEADesc).Value = pLEAName
    End Sub

    Public Sub Delete()
      mvClassFields.Delete(mvEnv.Connection)
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
        AddressNumber = mvClassFields.Item(AddressDataFields.adfAddressNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property LEACode() As String
      Get
        LEACode = mvClassFields.Item(AddressDataFields.adfLEACode).Value
      End Get
    End Property

    Public ReadOnly Property LEAName() As String
      Get
        LEAName = mvClassFields.Item(AddressDataFields.adfLEADesc).Value
      End Get
    End Property
  End Class
End Namespace
