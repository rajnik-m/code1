Namespace Access

  Public Class WebPageItemLink
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum WebPageItemLinkFields
      AllFields = 0
      WebPageItemLinkNumber
      WebPageNumber
      WebPageItemNumber
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("web_page_item_link_number", CDBField.FieldTypes.cftLong)
        .Add("web_page_number", CDBField.FieldTypes.cftLong)
        .Add("web_page_item_number", CDBField.FieldTypes.cftLong)

        .Item(WebPageItemLinkFields.WebPageItemLinkNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "wpil"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "web_page_item_links"
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
    Public ReadOnly Property WebPageItemLinkNumber() As Integer
      Get
        Return mvClassFields(WebPageItemLinkFields.WebPageItemLinkNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebPageNumber() As Integer
      Get
        Return mvClassFields(WebPageItemLinkFields.WebPageNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebPageItemNumber() As Integer
      Get
        Return mvClassFields(WebPageItemLinkFields.WebPageItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(WebPageItemLinkFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(WebPageItemLinkFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"
#End Region

  End Class
End Namespace