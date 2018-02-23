Namespace Access

  Public Class WebMenu
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum WebMenuFields
      AllFields = 0
      WebMenuNumber
      WebMenuName
      WebMenuStyle
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("web_menu_number", CDBField.FieldTypes.cftLong)
        .Add("web_menu_name")
        .Add("web_menu_style")

        .Item(WebMenuFields.WebMenuNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "wm"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "web_menus"
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
    Public ReadOnly Property WebMenuNumber() As Integer
      Get
        Return mvClassFields(WebMenuFields.WebMenuNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property WebMenuName() As String
      Get
        Return mvClassFields(WebMenuFields.WebMenuName).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(WebMenuFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(WebMenuFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property WebMenuStyle() As String
      Get
        Return mvClassFields(WebMenuFields.WebMenuStyle).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      mvEnv.Connection.StartTransaction()
      Dim vWMI As New WebMenuItem(mvEnv)
      vWMI.DeleteByForeignKey(New CDBField("web_menu_number", WebMenuNumber))
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Public Function GetNextMenuItemNumber() As Integer
      Dim vWC As New WebControl(mvEnv)
      vWC.Init(WebMenuNumber \ WebControl.ItemsMultiplier)
      If vWC.Existing = False Then RaiseError(DataAccessErrors.daeInvalidCode, "WebNumber")
      Return vWC.AllocateNextNumber(WebControl.WebNumberFields.wnfMenuItemNumber)
    End Function

#End Region
  End Class
End Namespace

