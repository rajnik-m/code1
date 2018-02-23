

Namespace Access
  Public Class ContactSearchName

    Public Enum ContactSearchNameRecordSetTypes 'These are bit values
      csnrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ContactSearchNameFields
      csnfAll = 0
      csnfContactNumber
      csnfSearchName
      csnfSoundexCode
      csnfIsActive
      csnfCreatedOn
      csnfAmendedBy
      csnfAmendedOn
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
          .DatabaseTableName = "contact_search_names"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("search_name", CDBField.FieldTypes.cftUnicode)
          .Add("soundex_code")
          .Add("is_active")
          .Add("created_on", CDBField.FieldTypes.cftTime)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(ContactSearchNameFields.csnfContactNumber).SetPrimaryKeyOnly()
          .Item(ContactSearchNameFields.csnfCreatedOn).SetPrimaryKeyOnly()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ContactSearchNameFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ContactSearchNameFields.csnfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ContactSearchNameFields.csnfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ContactSearchNameRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ContactSearchNameRecordSetTypes.csnrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "csn")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pContactNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pContactNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ContactSearchNameRecordSetTypes.csnrtAll) & " FROM contact_search_names csn WHERE contact_number = " & pContactNumber & " AND is_active = 'Y'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ContactSearchNameRecordSetTypes.csnrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactSearchNameRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ContactSearchNameFields.csnfContactNumber, vFields)
        .SetItem(ContactSearchNameFields.csnfCreatedOn, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ContactSearchNameRecordSetTypes.csnrtAll) = ContactSearchNameRecordSetTypes.csnrtAll Then
          .SetItem(ContactSearchNameFields.csnfSearchName, vFields)
          .SetItem(ContactSearchNameFields.csnfSoundexCode, vFields)
          .SetItem(ContactSearchNameFields.csnfIsActive, vFields)
          .SetItem(ContactSearchNameFields.csnfAmendedBy, vFields)
          .SetItem(ContactSearchNameFields.csnfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ContactSearchNameFields.csnfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Update(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByVal pName As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      If pEnv.GetConfigOption("cd_advanced_name_searching", False) Then
        Init(pEnv)
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        vUpdateFields.Add("is_active", CDBField.FieldTypes.cftCharacter, "N")
        vUpdateFields.AddAmendedOnBy(mvEnv.User.Logname)
        mvEnv.Connection.UpdateRecords((mvClassFields.DatabaseTableName), vUpdateFields, vWhereFields, False)
        Create(mvEnv, pContactNumber, pName, False)
      End If
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByVal pName As String, Optional ByVal pCheckConfig As Boolean = True, Optional ByVal pActive As Boolean = True)
      Dim vAddRecord As Boolean

      If pCheckConfig Then
        vAddRecord = pEnv.GetConfigOption("cd_advanced_name_searching", False)
      Else
        vAddRecord = True
      End If
      If vAddRecord Then
        If mvClassFields Is Nothing Then Init(pEnv)
        With mvClassFields
          .Item(ContactSearchNameFields.csnfContactNumber).IntegerValue = pContactNumber
          .Item(ContactSearchNameFields.csnfSearchName).Value = LCase(pName)
          .Item(ContactSearchNameFields.csnfSoundexCode).Value = GetSoundexCode(pName)
          .Item(ContactSearchNameFields.csnfCreatedOn).Value = TodaysDateAndTime()
          .Item(ContactSearchNameFields.csnfIsActive).Bool = pActive
        End With
        Save()
      End If
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
        AmendedBy = mvClassFields.Item(ContactSearchNameFields.csnfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ContactSearchNameFields.csnfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(ContactSearchNameFields.csnfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(ContactSearchNameFields.csnfCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property IsActive() As Boolean
      Get
        IsActive = mvClassFields.Item(ContactSearchNameFields.csnfIsActive).Bool
      End Get
    End Property

    Public ReadOnly Property SearchName() As String
      Get
        SearchName = mvClassFields.Item(ContactSearchNameFields.csnfSearchName).Value
      End Get
    End Property

    Public ReadOnly Property SoundexCode() As String
      Get
        SoundexCode = mvClassFields.Item(ContactSearchNameFields.csnfSoundexCode).Value
      End Get
    End Property
  End Class
End Namespace
