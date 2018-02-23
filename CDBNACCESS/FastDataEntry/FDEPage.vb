Namespace Access

  Public Class FDEPage
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum FdePageFields
      AllFields = 0
      FdePageNumber
      FdePageName
      FdePageTitle
      FdePageHeight
      FdePageWidth
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("fde_page_number", CDBField.FieldTypes.cftLong)
        .Add("fde_page_name")
        .Add("fde_page_title")
        .Add("fde_page_height", CDBField.FieldTypes.cftInteger)
        .Add("fde_page_width", CDBField.FieldTypes.cftInteger)

        .SetControlNumberField(FdePageFields.FdePageNumber, "FP")

        .Item(FdePageFields.FdePageNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "fdep"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "fde_pages"
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
    Public ReadOnly Property FdePageNumber() As Integer
      Get
        Return mvClassFields(FdePageFields.FdePageNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property FdePageName() As String
      Get
        Return mvClassFields(FdePageFields.FdePageName).Value
      End Get
    End Property
    Public ReadOnly Property FdePageTitle() As String
      Get
        Return mvClassFields(FdePageFields.FdePageTitle).Value
      End Get
    End Property
    Public ReadOnly Property FdePageHeight() As Integer
      Get
        Return mvClassFields(FdePageFields.FdePageHeight).IntegerValue
      End Get
    End Property
    Public ReadOnly Property FdePageWidth() As Integer
      Get
        Return mvClassFields(FdePageFields.FdePageWidth).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(FdePageFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(FdePageFields.AmendedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace
