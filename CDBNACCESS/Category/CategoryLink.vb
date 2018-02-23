Namespace Access

  Public Class CategoryLink
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum CategoryLinkFields
      AllFields = 0
      CategoryLinkId
      ExamUnitLinkId
      ExamCentreId
      ExamCentreUnitId
      CategoryId
      WorkstreamId
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("category_link_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("category_id", CDBField.FieldTypes.cftInteger)
        .Add("workstream_id", CDBField.FieldTypes.cftInteger)

        .Item(CategoryLinkFields.CategoryLinkId).PrimaryKey = True
        .Item(CategoryLinkFields.CategoryLinkId).PrefixRequired = True
        .SetControlNumberField(CategoryLinkFields.CategoryLinkId, "CLI")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "catl"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "category_links"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'AddDeleteCheckItems
    '--------------------------------------------------
    Public Overrides Sub AddDeleteCheckItems()
      AddCascadeDeleteItem("categories", "category_link_id")
    End Sub
'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property CategoryLinkId() As Integer
      Get
        Return mvClassFields(CategoryLinkFields.CategoryLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(CategoryLinkFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(CategoryLinkFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreUnitId() As Integer
      Get
        Return mvClassFields(CategoryLinkFields.ExamCentreUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CategoryId() As Integer
      Get
        Return mvClassFields(CategoryLinkFields.CategoryId).IntegerValue
      End Get
    End Property
#End Region

  End Class
End Namespace