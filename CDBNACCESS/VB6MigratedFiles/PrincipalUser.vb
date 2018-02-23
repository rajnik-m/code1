

Namespace Access
  Public Class PrincipalUser

    Public Enum PrincipalUserRecordSetTypes 'These are bit values
      purtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum PrincipalUserFields
      pufAll = 0
      pufContactNumber
      pufPrincipalUser
      pufPrincipalUserReason
      pufAmendedBy
      pufAmendedOn
    End Enum

    Private mvUser As CDBUser
    Private mvUserContact As Contact

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
          .DatabaseTableName = "principal_users"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("principal_user")
          .Add("principal_user_reason")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(PrincipalUserFields.pufContactNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As PrincipalUserFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(PrincipalUserFields.pufAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PrincipalUserFields.pufAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As PrincipalUserRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PrincipalUserRecordSetTypes.purtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pu")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pContactNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pContactNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet(pEnv.Connection.GetSelectSQLCSC & GetRecordSetFields(PrincipalUserRecordSetTypes.purtAll) & " FROM principal_users pu WHERE contact_number = " & pContactNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PrincipalUserRecordSetTypes.purtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PrincipalUserRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PrincipalUserFields.pufContactNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PrincipalUserRecordSetTypes.purtAll) = PrincipalUserRecordSetTypes.purtAll Then
          .SetItem(PrincipalUserFields.pufPrincipalUser, vFields)
          .SetItem(PrincipalUserFields.pufPrincipalUserReason, vFields)
          .SetItem(PrincipalUserFields.pufAmendedBy, vFields)
          .SetItem(PrincipalUserFields.pufAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PrincipalUserFields.pufAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
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
        AmendedBy = mvClassFields.Item(PrincipalUserFields.pufAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PrincipalUserFields.pufAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(PrincipalUserFields.pufContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PrincipalUserLogname() As String
      Get
        PrincipalUserLogname = mvClassFields.Item(PrincipalUserFields.pufPrincipalUser).Value
      End Get
    End Property

    Public ReadOnly Property PrincipalUserReason() As String
      Get
        PrincipalUserReason = mvClassFields.Item(PrincipalUserFields.pufPrincipalUserReason).Value
      End Get
    End Property

    Public ReadOnly Property PrincipalUserName() As String
      Get
        Dim vName As String = ""
        If mvExisting Then
          InitUserAndContact()
          vName = mvUserContact.LabelName
          If Len(vName) = 0 Then vName = mvUser.FullName
        End If
        PrincipalUserName = vName
      End Get
    End Property

    Public ReadOnly Property PrincipalUserContact() As Contact
      Get
        InitUserAndContact()
        PrincipalUserContact = mvUserContact
      End Get
    End Property

    Public ReadOnly Property PrincipalUserCDBUser() As CDBUser
      Get
        InitUserAndContact()
        PrincipalUserCDBUser = mvUser
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByVal pUserName As String, ByVal pUserReason As String)
      Init(pEnv)
      With mvClassFields
        .Item(PrincipalUserFields.pufContactNumber).Value = CStr(pContactNumber)
        .Item(PrincipalUserFields.pufPrincipalUser).Value = pUserName
        .Item(PrincipalUserFields.pufPrincipalUserReason).Value = pUserReason
      End With
      Save(mvEnv.User.UserID, True)
    End Sub
    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
    End Sub

    Public Sub UpdateUserAndReason(ByVal pUser As String, ByVal pUserReason As String)
      With mvClassFields
        .Item(PrincipalUserFields.pufPrincipalUser).Value = pUser
        .Item(PrincipalUserFields.pufPrincipalUserReason).Value = pUserReason
      End With
      Save(mvEnv.User.UserID, True)
    End Sub

    Private Sub InitUserAndContact()
      If mvUser Is Nothing Then
        mvUser = New CDBUser(mvEnv)
        mvUser.Init(PrincipalUserLogname)
        mvUserContact = New Contact(mvEnv)
        mvUserContact.Init((mvUser.ContactNumber))
      End If
    End Sub
  End Class
End Namespace
