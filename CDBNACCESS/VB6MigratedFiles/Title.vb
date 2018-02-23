

Namespace Access
  Public Class Title

    Public Enum TitleRecordSetTypes 'These are bit values
      trtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum TitleFields
      tfAll = 0
      tfTitle
      tfSex
      tfSalutation
      tfAmendedBy
      tfAmendedOn
      tfShortcutKey
      tfJointTitle
      tfInformalSalutation
      tfLabelName
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
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "titles"
          .Add("title")
          .Add("sex")
          .Add("salutation")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("shortcut_key")
          .Add("joint_title")
          .Add("informal_salutation")
          .Add("label_name")
        End With
        mvClassFields.Item(TitleFields.tfTitle).SetPrimaryKeyOnly()
        mvClassFields.Item(TitleFields.tfInformalSalutation).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport)
        mvClassFields.Item(TitleFields.tfLabelName).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataLabelName)
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByRef pField As TitleFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(TitleFields.tfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(TitleFields.tfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As TitleRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = TitleRecordSetTypes.trtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "t")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pTitle As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pTitle.Length > 0 Then
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, GetRecordSetFields(TitleRecordSetTypes.trtAll), "titles", New CDBField("title", pTitle))
        vRecordSet = vSQLStatement.GetRecordSet
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, TitleRecordSetTypes.trtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As TitleRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(TitleFields.tfTitle, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And TitleRecordSetTypes.trtAll) = TitleRecordSetTypes.trtAll Then
          .SetItem(TitleFields.tfSex, vFields)
          .SetItem(TitleFields.tfSalutation, vFields)
          .SetItem(TitleFields.tfAmendedBy, vFields)
          .SetItem(TitleFields.tfAmendedOn, vFields)
          .SetOptionalItem(TitleFields.tfShortcutKey, vFields)
          .SetOptionalItem(TitleFields.tfJointTitle, vFields)
          .SetOptionalItem(TitleFields.tfInformalSalutation, vFields)
          .SetOptionalItem(TitleFields.tfLabelName, vFields)
        End If
      End With
    End Sub

    Public Sub Save()
      SetValid(TitleFields.tfAll)
      mvClassFields.Save(mvEnv, mvExisting)
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
        AmendedBy = mvClassFields.Item(TitleFields.tfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(TitleFields.tfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property JointTitle() As Boolean
      Get
        JointTitle = mvClassFields.Item(TitleFields.tfJointTitle).Bool
      End Get
    End Property

    Public ReadOnly Property Salutation() As String
      Get
        Salutation = mvClassFields.Item(TitleFields.tfSalutation).Value
      End Get
    End Property
    Public ReadOnly Property LabelName() As String
      Get
        LabelName = mvClassFields.Item(TitleFields.tfLabelName).Value
      End Get
    End Property
    Public ReadOnly Property InformalSalutation() As String
      Get
        InformalSalutation = mvClassFields.Item(TitleFields.tfInformalSalutation).Value
      End Get
    End Property
    Public ReadOnly Property Sex() As String
      Get
        Sex = mvClassFields.Item(TitleFields.tfSex).Value
      End Get
    End Property

    Public ReadOnly Property ShortcutKey() As String
      Get
        ShortcutKey = mvClassFields.Item(TitleFields.tfShortcutKey).Value
      End Get
    End Property

    Public ReadOnly Property TitleName() As String
      Get
        TitleName = mvClassFields.Item(TitleFields.tfTitle).Value
      End Get
    End Property
  End Class
End Namespace
