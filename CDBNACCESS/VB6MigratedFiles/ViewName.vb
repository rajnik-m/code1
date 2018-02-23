

Namespace Access
  Public Class ViewName

    Public Enum ViewNameRecordSetTypes 'These are bit values
      vnrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ViewNameFields
      vnfAll = 0
      vnfViewName
      vnfViewNameDesc
      vnfViewType
      vnfVersionNumber
      vnfTableNames
      vnfClient
      vnfDepartment
      vnfLogname
      vnfNotes
      vnfAmendedBy
      vnfAmendedOn
      vnfDashboardGeneralView
    End Enum

    Public Enum ViewTypes
      vtSelection
      vtInternal
      vtReporting
      vtContact
      vtOrganisation
      tvEvent
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
          .DatabaseTableName = "view_names"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("view_name")
          .Add("view_name_desc")
          .Add("view_type")
          .Add("version_number")
          .Add("table_names")
          .Add("client")
          .Add("department")
          .Add("logname")
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("dashboard_general_view")
          .Item(ViewNameFields.vnfViewName).SetPrimaryKeyOnly()
          .Item(ViewNameFields.vnfDashboardGeneralView).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDashboardViewNames)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDashboardViewNames) = False Then
            .Item(ViewNameFields.vnfDashboardGeneralView).InDatabase = mvEnv.Connection.AttributeExists("view_names", "dashboard_general_view")
            'As this gets used by dbUpgrade before maintenance data is upgraded, check to see if the attribute does actually exit
          End If
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ViewNameFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ViewNameFields.vnfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ViewNameFields.vnfAmendedBy).Value = If(mvEnv.InitialisingDatabase AndAlso String.IsNullOrWhiteSpace(mvEnv.User.Logname), "dbinit", mvEnv.User.Logname)
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ViewNameRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ViewNameRecordSetTypes.vnrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "vn")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pViewName As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pViewName) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ViewNameRecordSetTypes.vnrtAll) & " FROM view_names vn WHERE view_name = '" & pViewName & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ViewNameRecordSetTypes.vnrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ViewNameRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And ViewNameRecordSetTypes.vnrtAll) = ViewNameRecordSetTypes.vnrtAll Then
          .SetItem(ViewNameFields.vnfViewName, vFields)
          .SetItem(ViewNameFields.vnfViewNameDesc, vFields)
          .SetItem(ViewNameFields.vnfViewType, vFields)
          .SetItem(ViewNameFields.vnfVersionNumber, vFields)
          .SetItem(ViewNameFields.vnfTableNames, vFields)
          .SetItem(ViewNameFields.vnfClient, vFields)
          .SetItem(ViewNameFields.vnfDepartment, vFields)
          .SetItem(ViewNameFields.vnfLogname, vFields)
          .SetItem(ViewNameFields.vnfNotes, vFields)
          .SetItem(ViewNameFields.vnfAmendedOn, vFields)
          .SetItem(ViewNameFields.vnfAmendedOn, vFields)
          .SetOptionalItem(ViewNameFields.vnfDashboardGeneralView, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ViewNameFields.vnfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByRef pEnv As CDBEnvironment, ByVal pViewName As String, ByVal pViewDesc As String, ByVal pViewType As String, ByVal pVersion As Integer, ByVal pDashboardGeneralView As Boolean)
      Init(pEnv)
      With mvClassFields
        .Item(ViewNameFields.vnfViewName).Value = pViewName
        .Item(ViewNameFields.vnfViewNameDesc).Value = pViewDesc
        .Item(ViewNameFields.vnfViewType).Value = pViewType
        .Item(ViewNameFields.vnfVersionNumber).Value = CStr(pVersion)
        .Item(ViewNameFields.vnfDashboardGeneralView).Bool = pDashboardGeneralView
      End With
    End Sub

    Public Sub Update(ByVal pViewDesc As String, ByVal pVersion As Integer)
      With mvClassFields
        .Item(ViewNameFields.vnfViewNameDesc).Value = pViewDesc
        .Item(ViewNameFields.vnfVersionNumber).Value = CStr(pVersion)
      End With
    End Sub

    Public Function GetSelectionViews() As Collection
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vViewName As ViewName
      Dim vColl As New Collection

      vSQL = "SELECT " & GetRecordSetFields(ViewNameRecordSetTypes.vnrtAll) & " FROM view_names vn WHERE view_type = 'S'"
      If Len(mvEnv.ClientCode) = 0 Then
        vSQL = vSQL & "AND ((client is null)"
      Else
        vSQL = vSQL & "AND ((client = '" & mvEnv.ClientCode & "' OR client IS NULL)"
      End If
      vSQL = vSQL & " AND (department = '" & mvEnv.User.Department & "' OR department IS NULL) AND (logname = '" & mvEnv.User.Logname & "' OR logname IS NULL)) "
      If mvEnv.Connection.NullsSortAtEnd Then
        vSQL = vSQL & "ORDER BY logname, department, client, view_name_desc"
      Else
        vSQL = vSQL & "ORDER BY logname DESC, department DESC, client DESC, view_name_desc"
      End If
      vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      While vRecordSet.Fetch() = True
        vViewName = New ViewName
        vViewName.InitFromRecordSet(mvEnv, vRecordSet, ViewNameRecordSetTypes.vnrtAll)
        vColl.Add(vViewName, vViewName.ViewNameCode)
      End While
      vRecordSet.CloseRecordSet()
      GetSelectionViews = vColl
    End Function

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
        AmendedBy = mvClassFields.Item(ViewNameFields.vnfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ViewNameFields.vnfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property ViewNameCode() As String
      Get
        ViewNameCode = mvClassFields.Item(ViewNameFields.vnfViewName).Value
      End Get
    End Property

    Public ReadOnly Property ViewNameDesc() As String
      Get
        ViewNameDesc = mvClassFields.Item(ViewNameFields.vnfViewNameDesc).Value
      End Get
    End Property

    Public ReadOnly Property Client() As String
      Get
        Client = mvClassFields.Item(ViewNameFields.vnfClient).Value
      End Get
    End Property

    Public ReadOnly Property Department() As String
      Get
        Department = mvClassFields.Item(ViewNameFields.vnfDepartment).Value
      End Get
    End Property

    Public ReadOnly Property Logname() As String
      Get
        Logname = mvClassFields.Item(ViewNameFields.vnfLogname).Value
      End Get
    End Property

    Public ReadOnly Property ViewType() As ViewTypes
      Get
        Select Case mvClassFields.Item(ViewNameFields.vnfViewType).Value
          Case "S"
            ViewType = ViewTypes.vtSelection
          Case "I"
            ViewType = ViewTypes.vtInternal
          Case "R"
            ViewType = ViewTypes.vtSelection
          Case "C"
            ViewType = ViewTypes.vtContact
          Case "O"
            ViewType = ViewTypes.vtOrganisation
          Case "E"
            ViewType = ViewTypes.tvEvent
        End Select
      End Get
    End Property

    Public ReadOnly Property VersionNumber() As Integer
      Get
        VersionNumber = mvClassFields.Item(ViewNameFields.vnfVersionNumber).IntegerValue
      End Get
    End Property

    Public Property TableNames() As String
      Get
        TableNames = mvClassFields.Item(ViewNameFields.vnfTableNames).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ViewNameFields.vnfTableNames).Value = Value
      End Set
    End Property

    Public Property Notes() As String
      Get
        Notes = mvClassFields.Item(ViewNameFields.vnfNotes).MultiLineValue
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(ViewNameFields.vnfNotes).Value = Value
      End Set
    End Property

    Public ReadOnly Property DashboardGeneralView() As Boolean
      Get
        DashboardGeneralView = mvClassFields.Item(ViewNameFields.vnfDashboardGeneralView).Bool
      End Get
    End Property

    Public Overrides Function ToString() As String
      Return Me.ViewNameCode
    End Function
  End Class
End Namespace
