Namespace Access
  Public Class Branch

    Public Enum BranchRecordSetTypes 'These are bit values
      bratAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BranchFields
      bfAll = 0
      bfBranch
      bfOrganisationNumber
      bfNominalAccount
      bfCharityNumber
      bfHistorical
      bfAmendedBy
      bfAmendedOn
      bfOwnershipGroup
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvOrganisation As Organisation
    Private mvTableList As String
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "branches"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("branch")
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("nominal_account")
          .Add("charity_number")
          .Add("historical")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("ownership_group")
        End With

        mvClassFields.Item(BranchFields.bfBranch).SetPrimaryKeyOnly()
        mvClassFields.Item(BranchFields.bfOwnershipGroup).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBranchOwnershipGroup)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
      mvTableList = "addresses,members,orders,branch_postcodes,branch_income" 'branch_income should always be the last table in this list
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As BranchFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(BranchFields.bfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(BranchFields.bfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As BranchRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BranchRecordSetTypes.bratAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "b")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pBranch As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pBranch) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(BranchRecordSetTypes.bratAll) & " FROM branches b WHERE branch = '" & pBranch & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BranchRecordSetTypes.bratAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BranchRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BranchFields.bfBranch, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BranchRecordSetTypes.bratAll) = BranchRecordSetTypes.bratAll Then
          .SetItem(BranchFields.bfOrganisationNumber, vFields)
          .SetItem(BranchFields.bfNominalAccount, vFields)
          .SetItem(BranchFields.bfCharityNumber, vFields)
          .SetItem(BranchFields.bfHistorical, vFields)
          .SetItem(BranchFields.bfAmendedBy, vFields)
          .SetItem(BranchFields.bfAmendedOn, vFields)
          .SetOptionalItem(BranchFields.bfOwnershipGroup, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(BranchFields.bfAll)
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
        AmendedBy = mvClassFields.Item(BranchFields.bfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(BranchFields.bfAmendedOn).Value
      End Get
    End Property

    Public Property BranchCode() As String
      Get
        BranchCode = mvClassFields.Item(BranchFields.bfBranch).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BranchFields.bfBranch).Value = Value
      End Set
    End Property

    Public ReadOnly Property CharityNumber() As String
      Get
        CharityNumber = mvClassFields.Item(BranchFields.bfCharityNumber).Value
      End Get
    End Property

    Public Property Historical() As Boolean
      Get
        Historical = mvClassFields.Item(BranchFields.bfHistorical).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(BranchFields.bfHistorical).Bool = Value
      End Set
    End Property

    Public ReadOnly Property NominalAccount() As String
      Get
        NominalAccount = mvClassFields.Item(BranchFields.bfNominalAccount).Value
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        OrganisationNumber = mvClassFields.Item(BranchFields.bfOrganisationNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Organisation() As Organisation
      Get
        If mvOrganisation Is Nothing Then
          If OrganisationNumber > 0 Then
            mvOrganisation = New Organisation(mvEnv)
            mvOrganisation.Init(OrganisationNumber)
          End If
        End If
        Organisation = mvOrganisation
      End Get
    End Property
    Public ReadOnly Property NeedToPromptForNewBranch() As Boolean
      Get
        'This method is used to determine whether the branch exists in the database.
        'It should be used before the MakeHistorical method.

        Dim vTables() As String
        Dim vIndex As Integer
        Dim vPrompt As Boolean
        Dim vWhereFields As New CDBFields

        vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, BranchCode)
        vTables = Split(mvTableList, ",")
        For vIndex = 0 To UBound(vTables)
          If vTables(vIndex) = "branch_income" Then 'branch_income should always be the last table in the list
            vWhereFields = New CDBFields
            vWhereFields.Add("branch_code", CDBField.FieldTypes.cftCharacter, BranchCode)
          End If
          vPrompt = mvEnv.Connection.GetCount(vTables(vIndex), vWhereFields, "") > 0
          If vPrompt Then Exit For
        Next
        NeedToPromptForNewBranch = vPrompt
      End Get
    End Property

    Public ReadOnly Property OwnershipGroup() As String
      Get
        OwnershipGroup = mvClassFields.Item(BranchFields.bfOwnershipGroup).Value
      End Get
    End Property

    Public Sub MakeHistorical(ByVal pNewBranch As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vTables() As String
      Dim vIndex As Integer

      If Len(pNewBranch) > 0 Then
        vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, BranchCode)
        vUpdateFields.Add("branch", CDBField.FieldTypes.cftCharacter, pNewBranch)
        vUpdateFields.AddAmendedOnBy((mvEnv.User.Logname))
        vTables = Split(mvTableList, ",")
        For vIndex = 0 To UBound(vTables)
          If vTables(vIndex) = "branch_income" Then 'branch_income should always be the last table in the list
            vWhereFields.Clear()
            vUpdateFields.Clear()
            vWhereFields.Add("branch_code", CDBField.FieldTypes.cftCharacter, BranchCode)
            vUpdateFields.Add("branch_code", CDBField.FieldTypes.cftCharacter, pNewBranch)
          End If
          mvEnv.Connection.UpdateRecords(vTables(vIndex), vUpdateFields, vWhereFields, False)
        Next
      End If
      mvClassFields.Item(BranchFields.bfHistorical).Bool = True
      Save()
    End Sub
  End Class
End Namespace
