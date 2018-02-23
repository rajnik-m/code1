Namespace Access

  Public Class DisplayListItem
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum DisplayListItemFields
      AllFields = 0
      DisplayList
      AccessMethod
      Client
      Department
      Logname
      ContactGroup
      DisplayTitle
      DisplayItems
      DisplayHeadings
      HeadingLines
      DisplaySizes
      WebPageItemNumber
      MaintenanceDesc
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("display_list")
        .Add("access_method")
        .Add("client")
        .Add("department")
        .Add("logname")
        .Add("contact_group")
        .Add("display_title")
        .Add("display_items")
        .Add("display_headings")
        .Add("heading_lines", CDBField.FieldTypes.cftInteger)
        .Add("display_sizes")
        .Add("web_page_item_number", CDBField.FieldTypes.cftLong)
        .Add("maintenance_desc")

        .Item(DisplayListItemFields.DisplayList).PrimaryKey = True
        .Item(DisplayListItemFields.AccessMethod).PrimaryKey = True
        .Item(DisplayListItemFields.Client).PrimaryKey = True
        .Item(DisplayListItemFields.Department).PrimaryKey = True
        .Item(DisplayListItemFields.Logname).PrimaryKey = True
        .Item(DisplayListItemFields.ContactGroup).PrimaryKey = True
        .Item(DisplayListItemFields.WebPageItemNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "dli"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "display_list_items"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property DisplayList() As String
      Get
        Return mvClassFields(DisplayListItemFields.DisplayList).Value
      End Get
    End Property
    Public ReadOnly Property AccessMethod() As String
      Get
        Return mvClassFields(DisplayListItemFields.AccessMethod).Value
      End Get
    End Property
    Public ReadOnly Property Client() As String
      Get
        Return mvClassFields(DisplayListItemFields.Client).Value
      End Get
    End Property
    Public ReadOnly Property Department() As String
      Get
        Return mvClassFields(DisplayListItemFields.Department).Value
      End Get
    End Property
    Public ReadOnly Property Logname() As String
      Get
        Return mvClassFields(DisplayListItemFields.Logname).Value
      End Get
    End Property
    Public ReadOnly Property ContactGroup() As String
      Get
        Return mvClassFields(DisplayListItemFields.ContactGroup).Value
      End Get
    End Property
    Public ReadOnly Property DisplayTitle() As String
      Get
        Return mvClassFields(DisplayListItemFields.DisplayTitle).Value
      End Get
    End Property
    Public ReadOnly Property DisplayItems() As String
      Get
        Return mvClassFields(DisplayListItemFields.DisplayItems).Value
      End Get
    End Property
    Public ReadOnly Property DisplayHeadings() As String
      Get
        Return mvClassFields(DisplayListItemFields.DisplayHeadings).Value
      End Get
    End Property
    Public ReadOnly Property HeadingLines() As Integer
      Get
        Return mvClassFields(DisplayListItemFields.HeadingLines).IntegerValue
      End Get
    End Property
    Public ReadOnly Property DisplaySizes() As String
      Get
        Return mvClassFields(DisplayListItemFields.DisplaySizes).Value
      End Get
    End Property
    Public ReadOnly Property WebPageItemNumber() As String
      Get
        Return mvClassFields(DisplayListItemFields.WebPageItemNumber).Value
      End Get
    End Property
    Public ReadOnly Property Maintenancedesc() As String
      Get
        Return mvClassFields.Item(DisplayListItemFields.MaintenanceDesc).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(DisplayListItemFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(DisplayListItemFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"
    Private mvName As String
    Private mvHeading As String
    Private mvWidth As Integer
    Private mvUserHeading As String
    Private mvUserWidth As Integer
    Private mvColumnWidth As Integer
    Private mvRequired As Boolean
    Private mvSystemHeading As Boolean
    Private mvReadOnly As Boolean

    Public Sub New(ByVal pEnv As CDBEnvironment, ByVal pName As String, ByVal pHeading As String, ByVal pWidth As Integer, ByVal pMandatory As Boolean)
      MyBase.New(pEnv)
      Dim vIndex As Integer
      Dim vChar As String

      mvName = pName
      If pHeading.Length = 0 Then
        mvHeading = pName.Substring(0, 1)
        For vIndex = 2 To pName.Length
          vChar = pName.Substring(vIndex - 1, 1)
          If vChar.ToUpper() = vChar Then mvHeading = mvHeading & " "
          mvHeading = mvHeading & vChar
        Next
        mvUserHeading = mvHeading
        mvSystemHeading = True
      Else
        mvHeading = pHeading
        mvUserHeading = pHeading
      End If
      If pWidth = 0 Then pWidth = 1200
      mvWidth = pWidth
      mvUserWidth = pWidth
      mvRequired = pMandatory
    End Sub

    Public ReadOnly Property Name() As String
      Get
        Return mvName
      End Get
    End Property

    Public Property UserHeading() As String
      Get
        Return mvUserHeading
      End Get
      Set(ByVal pValue As String)
        mvUserHeading = pValue
      End Set
    End Property
    Public Property UserWidth() As Integer
      Get
        Return mvUserWidth
      End Get
      Set(ByVal pValue As Integer)
        mvUserWidth = pValue
      End Set
    End Property

    Public Property IsReadOnly() As Boolean
      Get
        Return mvReadOnly
      End Get
      Set(ByVal pValue As Boolean)
        mvReadOnly = pValue
      End Set
    End Property
    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateCreateParameters(pParameterList)

      Dim vDataSelection As DataSelection
      If pParameterList("DataSelectionType").IntegerValue > ExamDataSelection.ExamDataSelectionTypes.dstNone AndAlso
         pParameterList("DataSelectionType").IntegerValue < WorkstreamDataSelection.DataSelectionTypes.dstNone Then
        Dim vEDST As ExamDataSelection.ExamDataSelectionTypes = CType(pParameterList("DataSelectionType").Value, ExamDataSelection.ExamDataSelectionTypes)
        vDataSelection = New ExamDataSelection(mvEnv, vEDST, Nothing, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      ElseIf pParameterList("DataSelectionType").IntegerValue > WorkstreamDataSelection.DataSelectionTypes.dstNone AndAlso
        pParameterList("DataSelectionType").IntegerValue < Access.Deduplication.DedupDataSelection.DataSelectionTypes.None Then
        Dim vWDST As WorkstreamDataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, WorkstreamDataSelection.DataSelectionTypes)
        vDataSelection = New WorkstreamDataSelection(mvEnv, vWDST, Nothing, WorkstreamDataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      ElseIf pParameterList("DataSelectionType").IntegerValue > Access.Deduplication.DedupDataSelection.DataSelectionTypes.None Then
        Dim vDDDST As Access.Deduplication.DedupDataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, Access.Deduplication.DedupDataSelection.DataSelectionTypes)
        vDataSelection = New Deduplication.DedupDataSelection(mvEnv, Nothing, Deduplication.DedupDataSelection.DataSelectionListType.dsltDefault, vDDDST)
      Else
        Dim vDST As DataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, DataSelection.DataSelectionTypes)
        vDataSelection = New DataSelection(mvEnv, vDST, DataSelection.DataSelectionListType.dsltDefault)
      End If
      If Not pParameterList.ContainsKey("HeadingLines") Then pParameterList.Add("HeadingLines", 0)
      pParameterList.Add("DisplayList", vDataSelection.DisplayListCode)
      If Not pParameterList.ContainsKey("Client") Then pParameterList.Add("Client", "")
      If Not pParameterList.ContainsKey("Department") Then pParameterList.Add("Department", "")
      If Not pParameterList.ContainsKey("Logname") Then pParameterList.Add("Logname", "")
      If pParameterList.ContainsKey("EventGroup") Then
        pParameterList.Add("ContactGroup", pParameterList("EventGroup").Value)
      ElseIf pParameterList.ContainsKey("WorkstreamGroup") Then
        pParameterList.Add("ContactGroup", pParameterList("WorkstreamGroup").Value)
      ElseIf pParameterList.ContainsKey("OrganisationGroup") Then
        pParameterList.Add("ContactGroup", pParameterList("OrganisationGroup").Value)
      ElseIf Not pParameterList.ContainsKey("ContactGroup") Then
        pParameterList.Add("ContactGroup", "")
      End If
      If Not pParameterList.ContainsKey("WebPageItemNumber") Then pParameterList.Add("WebPageItemNumber", "")
      Me.Init(pParameterList) 'used to initialise record for update
    End Sub

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return "DataSelectionType"
    End Function

    Public Function GetDisplayListCode(ByVal pParameterList As CDBParameters) As String
      Dim vDataSelection As DataSelection
      If pParameterList("DataSelectionType").IntegerValue > ExamDataSelection.ExamDataSelectionTypes.dstNone AndAlso
         pParameterList("DataSelectionType").IntegerValue < WorkstreamDataSelection.DataSelectionTypes.dstNone Then
        Dim vEDST As ExamDataSelection.ExamDataSelectionTypes = CType(pParameterList("DataSelectionType").Value, ExamDataSelection.ExamDataSelectionTypes)
        vDataSelection = New ExamDataSelection(mvEnv, vEDST, Nothing, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      ElseIf pParameterList("DataSelectionType").IntegerValue > WorkstreamDataSelection.DataSelectionTypes.dstNone AndAlso
        pParameterList("DataSelectionType").IntegerValue < Deduplication.DedupDataSelection.DataSelectionTypes.None Then
        Dim vWDST As WorkstreamDataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, WorkstreamDataSelection.DataSelectionTypes)
        vDataSelection = New WorkstreamDataSelection(mvEnv, vWDST, Nothing, DataSelection.DataSelectionListType.dsltDefault, DataSelection.DataSelectionUsages.dsuSmartClient)
      ElseIf pParameterList("DataSelectionType").IntegerValue > Deduplication.DedupDataSelection.DataSelectionTypes.None Then
        Dim vDDDST As Deduplication.DedupDataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, Deduplication.DedupDataSelection.DataSelectionTypes)
        vDataSelection = New Deduplication.DedupDataSelection(mvEnv, Nothing, DataSelection.DataSelectionListType.dsltDefault, vDDDST)
      Else
        Dim vDST As DataSelection.DataSelectionTypes = CType(pParameterList("DataSelectionType").Value, DataSelection.DataSelectionTypes)
        vDataSelection = New DataSelection(mvEnv, vDST, DataSelection.DataSelectionListType.dsltDefault)
      End If
      Return vDataSelection.DisplayListCode
    End Function

    Public Overrides Sub PreValidateParameterList(ByVal pType As MaintenanceTypes, ByVal pParameterList As CDBParameters)
      If pType = MaintenanceTypes.Delete Then
        pParameterList.Add("AccessMethod", "W")
        CheckClassFields()
        With mvClassFields
          .Item(DisplayListItemFields.DisplayList).PrimaryKey = False 'so no longer mandatory fields
          .Item(DisplayListItemFields.Client).PrimaryKey = False
          .Item(DisplayListItemFields.Department).PrimaryKey = False
          .Item(DisplayListItemFields.Logname).PrimaryKey = False
          .Item(DisplayListItemFields.ContactGroup).PrimaryKey = False
        End With
      End If
    End Sub

#End Region
  End Class


End Namespace
