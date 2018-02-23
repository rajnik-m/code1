Namespace Access

  Public Class WebMenuItem
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum WebMenuItemFields
      AllFields = 0
      WebMenuItemNumber
      ParentItemNumber
      WebMenuNumber
      WebPageNumber
      MenuTitle
      MenuDesc
      SequenceNumber
      WebUrl
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("web_menu_item_number", CDBField.FieldTypes.cftLong)
        .Add("parent_item_number", CDBField.FieldTypes.cftLong)
        .Add("web_menu_number", CDBField.FieldTypes.cftLong)
        .Add("web_page_number", CDBField.FieldTypes.cftLong)
        .Add("menu_title")
        .Add("menu_desc")
        .Add("sequence_number", CDBField.FieldTypes.cftLong)
        .Add("web_url")

        .Item(WebMenuItemFields.WebMenuItemNumber).PrimaryKey = True

        .Item(WebMenuItemFields.WebMenuNumber).PrefixRequired = True
        .Item(WebMenuItemFields.WebUrl).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataWebFriendlyUrl)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "wmi"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "web_menu_items"
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
    Public ReadOnly Property WebMenuItemNumber() As Integer
      Get
        Return mvClassFields(WebMenuItemFields.WebMenuItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ParentItemNumber() As Integer
      Get
        Return mvClassFields(WebMenuItemFields.ParentItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebMenuNumber() As Integer
      Get
        Return mvClassFields(WebMenuItemFields.WebMenuNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebPageNumber() As Integer
      Get
        Return mvClassFields(WebMenuItemFields.WebPageNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MenuTitle() As String
      Get
        Return mvClassFields(WebMenuItemFields.MenuTitle).Value
      End Get
    End Property
    Public ReadOnly Property MenuDesc() As String
      Get
        Return mvClassFields(WebMenuItemFields.MenuDesc).Value
      End Get
    End Property
    Public ReadOnly Property SequenceNumber() As Integer
      Get
        Return mvClassFields(WebMenuItemFields.SequenceNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebUrl() As String
      Get
        Return mvClassFields(WebMenuItemFields.WebUrl).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(WebMenuItemFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(WebMenuItemFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Sub CheckParentLink()
      'If we are attaching this menu item to another one as it's parent then make sure the parent has no link to a web page
      'as it should just be a header item for a set of sub-items
      If ParentItemNumber > 0 And mvClassFields(WebMenuItemFields.ParentItemNumber).ValueChanged Then
        Dim vWhereFields As New CDBFields(New CDBField("web_menu_item_number", ParentItemNumber))
        Dim vUpdateFields As New CDBFields(New CDBField("web_page_number", CDBField.FieldTypes.cftInteger))
        mvEnv.Connection.UpdateRecords(DatabaseTableName, vUpdateFields, vWhereFields, False)
      End If
    End Sub

    Protected Overrides Sub SetValid()
      CheckParentLink()
      MyBase.SetValid()
    End Sub

#End Region

  End Class
End Namespace

 
