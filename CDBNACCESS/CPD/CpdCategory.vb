Namespace Access

  Public Class CpdCategory
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum CpdCategoryFields
      AllFields = 0
      CpdCategory
      CpdCategoryDesc
      CpdCategoryType
      ValidFrom
      ValidTo
      CpdPoints
      PointsOverride
      DateMandatory
      Approved
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("cpd_category")
        .Add("cpd_category_desc")
        .Add("cpd_category_type")
        .Add("valid_from")
        .Add("valid_to")
        .Add("cpd_points")
        .Add("points_override")
        .Add("date_mandatory")
        .Add("approved")

        .Item(CpdCategoryFields.CpdCategory).PrimaryKey = True
        .Item(CpdCategoryFields.ValidFrom).InDatabase = True
        .Item(CpdCategoryFields.ValidTo).InDatabase = True
        .Item(CpdCategoryFields.CpdPoints).InDatabase = True
        .Item(CpdCategoryFields.PointsOverride).InDatabase = True
        .Item(CpdCategoryFields.DateMandatory).InDatabase = True
        .Item(CpdCategoryFields.Approved).InDatabase = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cc"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "cpd_categories"
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
    Public ReadOnly Property CpdCategoryCode() As String
      Get
        Return mvClassFields(CpdCategoryFields.CpdCategory).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategoryDesc() As String
      Get
        Return mvClassFields(CpdCategoryFields.CpdCategoryDesc).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategoryType() As String
      Get
        Return mvClassFields(CpdCategoryFields.CpdCategoryType).Value
      End Get
    End Property
    Public ReadOnly Property ValidFrom As String
      Get
        Return mvClassFields(CpdCategoryFields.ValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property ValidTo As String
      Get
        Return mvClassFields(CpdCategoryFields.ValidTo).Value
      End Get
    End Property

    Public ReadOnly Property CpdPoints() As String
      Get
        Return mvClassFields(CpdCategoryFields.CpdPoints).Value
      End Get
    End Property
    Public ReadOnly Property PointsOverride() As String
      Get
        Return mvClassFields(CpdCategoryFields.PointsOverride).Value
      End Get
    End Property
    Public ReadOnly Property DateMandatory() As String
      Get
        Return mvClassFields(CpdCategoryFields.DateMandatory).Value
      End Get
    End Property
    Public ReadOnly Property Approved() As Boolean
      Get
        If mvClassFields(CpdCategoryFields.Approved).Value.Length > 0 Then
          Return mvClassFields(CpdCategoryFields.Approved).Bool
        Else
          Return True
        End If
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CpdCategoryFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CpdCategoryFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    ''' <summary>Validates that the Category is approved and valid based upon it's valid from and to dates.</summary>
    ''' <returns>True if it is valid, otherwise False.</returns>
    Friend Function IsValidForPointsEntry(ByVal pPointsDate As Date) As Boolean
      Dim vIsValid As Boolean = Approved
      If Approved = True Then
        Dim vValidFrom As Nullable(Of Date)
        Dim vValidTo As Nullable(Of Date)
        If IsDate(ValidFrom) Then vValidFrom = Date.Parse(ValidFrom)
        If IsDate(vValidTo) Then vValidTo = Date.Parse(ValidTo)

        If vValidFrom.HasValue AndAlso vValidFrom.Value.CompareTo(pPointsDate) > 0 Then vIsValid = False 'ValidFrom > Today
        If vValidTo.HasValue AndAlso vValidTo.Value.CompareTo(pPointsDate) < 0 Then vIsValid = False 'ValidTo < Today
      End If

      Return vIsValid

    End Function

#End Region

  End Class
End Namespace
