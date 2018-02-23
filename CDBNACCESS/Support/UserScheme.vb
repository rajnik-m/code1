Namespace Access

  Public Class UserScheme
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum UserSchemeFields
      AllFields = 0
      UserSchemeId
      UserSchemeDesc
      UserSchemeCode
      AppearanceXmlItemNumber
      FontXmlItemNumber
      IsDefault
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("user_scheme_id", CDBField.FieldTypes.cftInteger)
        .Add("user_scheme_desc")
        .Add("user_scheme_code")
        .Add("appearance_xml_item_number", CDBField.FieldTypes.cftInteger)
        .Add("font_xml_item_number", CDBField.FieldTypes.cftInteger)
        .Add("is_default")
        .Item(UserSchemeFields.UserSchemeId).PrimaryKey = True
        .SetControlNumberField(UserSchemeFields.UserSchemeId, "USN")
      End With

    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "us"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "user_schemes"
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
    Public ReadOnly Property UserSchemeId() As Integer
      Get
        Return mvClassFields(UserSchemeFields.UserSchemeId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property UserSchemeDesc() As String
      Get
        Return mvClassFields(UserSchemeFields.UserSchemeDesc).Value
      End Get
    End Property
    Public ReadOnly Property UserSchemeCode() As String
      Get
        Return mvClassFields(UserSchemeFields.UserSchemeCode).Value
      End Get
    End Property
    Public ReadOnly Property AppearanceXmlItemNumber() As Integer
      Get
        Return mvClassFields(UserSchemeFields.AppearanceXmlItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property FontXmlItemNumber() As Integer
      Get
        Return mvClassFields(UserSchemeFields.FontXmlItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property IsDefault() As String
      Get
        Return mvClassFields(UserSchemeFields.IsDefault).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(UserSchemeFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(UserSchemeFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non Auto Generated"
    Public Function ItemExists(ByVal pSchemeCode As String, ByVal pSchemeDesc As String) As Boolean
      Init()
      Dim vWhereFields As New CDBFields
      If pSchemeCode IsNot Nothing AndAlso pSchemeCode.Length > 0 Then vWhereFields.Add(mvClassFields(UserSchemeFields.UserSchemeCode).Name, pSchemeCode)
      If pSchemeDesc IsNot Nothing AndAlso pSchemeDesc.Length > 0 Then vWhereFields.Add(mvClassFields(UserSchemeFields.UserSchemeDesc).Name, pSchemeDesc)
      InitWithPrimaryKey(vWhereFields)
      Return Me.Existing
    End Function

#End Region

  End Class
End Namespace
