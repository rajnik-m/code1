Namespace Access

  Public Class CustomMergeData
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum CustomMergeDataFields
      AllFields = 0
      Client
      DbName
      SelectSql
      SequenceNumber
      AttributeNames
      AttributeCaptions
      UsageCode
      MultiLine
      ContactGroup
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("client")
        .Add("db_name")
        .Add("select_sql", CDBField.FieldTypes.cftMemo)
        .Add("sequence_number", CDBField.FieldTypes.cftInteger)
        .Add("attribute_names", CDBField.FieldTypes.cftMemo)
        .Add("attribute_captions", CDBField.FieldTypes.cftMemo)
        .Add("usage_code")
        .Add("multi_line")
        .Add("contact_group")

        .Item(CustomMergeDataFields.Client).PrimaryKey = True

        .Item(CustomMergeDataFields.DbName).PrimaryKey = True

        .Item(CustomMergeDataFields.SequenceNumber).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cmd"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "custom_merge_data"
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
    Public ReadOnly Property Client() As String
      Get
        Return mvClassFields(CustomMergeDataFields.Client).Value
      End Get
    End Property
    Public ReadOnly Property DbName() As String
      Get
        Return mvClassFields(CustomMergeDataFields.DbName).Value
      End Get
    End Property
    Public ReadOnly Property SelectSql() As String
      Get
        Return mvClassFields(CustomMergeDataFields.SelectSql).Value
      End Get
    End Property
    Public ReadOnly Property SequenceNumber() As Integer
      Get
        Return mvClassFields(CustomMergeDataFields.SequenceNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AttributeNames() As String
      Get
        Return mvClassFields(CustomMergeDataFields.AttributeNames).Value
      End Get
    End Property
    Public ReadOnly Property AttributeCaptions() As String
      Get
        Return mvClassFields(CustomMergeDataFields.AttributeCaptions).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CustomMergeDataFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CustomMergeDataFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property UsageCode() As String
      Get
        Return mvClassFields(CustomMergeDataFields.UsageCode).Value
      End Get
    End Property
    Public ReadOnly Property MultiLine() As String
      Get
        Return mvClassFields(CustomMergeDataFields.MultiLine).Value
      End Get
    End Property
    Public ReadOnly Property ContactGroup() As String
      Get
        Return mvClassFields(CustomMergeDataFields.ContactGroup).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Public Function StandardAttributeNames() As String
      Dim vNames As New CDBParameters
      vNames.InitFromUniqueList(AttributeNames)
      Return vNames.StandardColumnNameList()
    End Function

    Public Function GetRecordSet(ByVal pValues As String) As CDBRecordSet
      Dim vSQL As String = SelectSql
      If pValues.Contains(",") Then vSQL = SelectSql.Replace(" = #", " IN (#)")
      Return New SQLStatement(mvEnv.Connection, vSQL.Replace("#", pValues)).GetRecordSet
    End Function

    Public Overloads Sub SetDataRow(ByVal pRow As CDBDataRow, ByVal pRS As CDBRecordSet, ByVal pColumns As CollectionList(Of CDBDataColumn))
      Dim vNames As New CDBParameters
      vNames.InitFromUniqueList(AttributeNames)
      For Each vParam As CDBParameter In vNames
        pRow.Item(ProperName(vParam.Name)) = pRS.Fields(vParam.Name).Value
        pColumns(ProperName(vParam.Name)).FieldType = pRS.Fields(vParam.Name).FieldType
      Next
    End Sub

#End Region
  End Class
End Namespace