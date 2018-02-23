Namespace Access
  Public Class Bank

    Public Enum BankRecordSetTypes 'These are bit values
      bkrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum BankFields
      bfAll = 0
      bfSortCode
      bfBank
      bfBranchName
      bfAddress
      bfTown
      bfCounty
      bfPostcode
      bfAmendedBy
      bfAmendedOn
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
          .DatabaseTableName = "banks"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("sort_code")
          .Add("bank")
          .Add("branch_name")
          .Add("address", CDBField.FieldTypes.cftMemo)
          .Add("town")
          .Add("county")
          .Add("postcode")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(BankFields.bfSortCode).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As BankFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(BankFields.bfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(BankFields.bfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByRef pRSType As BankRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = BankRecordSetTypes.bkrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "b")
      End If
      Return vFields
    End Function

    Public Sub Init(ByRef pEnv As CDBEnvironment, Optional ByRef pSortCode As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pSortCode) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(BankRecordSetTypes.bkrtAll) & " FROM banks WHERE sort_code = '" & pSortCode & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, BankRecordSetTypes.bkrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As BankRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(BankFields.bfSortCode, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And BankRecordSetTypes.bkrtAll) = BankRecordSetTypes.bkrtAll Then
          .SetItem(BankFields.bfBank, vFields)
          .SetItem(BankFields.bfBranchName, vFields)
          .SetItem(BankFields.bfAddress, vFields)
          .SetItem(BankFields.bfTown, vFields)
          .SetItem(BankFields.bfCounty, vFields)
          .SetItem(BankFields.bfPostcode, vFields)
          .SetItem(BankFields.bfAmendedBy, vFields)
          .SetItem(BankFields.bfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Create(ByRef pSortCode As String, ByRef pBranchName As String)
      mvClassFields.Item(BankFields.bfSortCode).Value = pSortCode
      mvClassFields.Item(BankFields.bfBranchName).Value = pBranchName
    End Sub
    Public Sub Create(ByVal pParams As CDBParameters)
      For Each vClassField As ClassField In mvClassFields
        If pParams.ContainsKey(vClassField.ProperName) Then vClassField.Value = pParams(vClassField.ProperName).Value
      Next
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(BankFields.bfAll)
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

    Public Property Address() As String
      Get
        Address = mvClassFields.Item(BankFields.bfAddress).MultiLineValue
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfAddress).Value = Value
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(BankFields.bfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(BankFields.bfAmendedOn).Value
      End Get
    End Property

    Public Property BankName() As String
      Get
        BankName = mvClassFields.Item(BankFields.bfBank).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfBank).Value = Value
      End Set
    End Property

    Public Property BranchName() As String
      Get
        BranchName = mvClassFields.Item(BankFields.bfBranchName).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfBranchName).Value = Value
      End Set
    End Property

    Public Property County() As String
      Get
        County = mvClassFields.Item(BankFields.bfCounty).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfCounty).Value = Value
      End Set
    End Property

    Public Property Postcode() As String
      Get
        Postcode = mvClassFields.Item(BankFields.bfPostcode).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfPostcode).Value = Value
      End Set
    End Property

    Public Property SortCode() As String
      Get
        SortCode = mvClassFields.Item(BankFields.bfSortCode).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfSortCode).Value = Value
      End Set
    End Property

    Public Property Town() As String
      Get
        Town = mvClassFields.Item(BankFields.bfTown).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(BankFields.bfTown).Value = Value
      End Set
    End Property
  End Class
End Namespace
