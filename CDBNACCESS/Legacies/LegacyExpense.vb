Namespace Access

  Public Class LegacyExpense
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum LegacyExpenseFields
      AllFields = 0
      LegacyNumber
      DateReceived
      Value
      Notes
      BequestNumber
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("legacy_number", CDBField.FieldTypes.cftLong)
        .Add("date_received", CDBField.FieldTypes.cftDate)
        .Add("value", CDBField.FieldTypes.cftNumeric)
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("bequest_number", CDBField.FieldTypes.cftLong)

        .Item(LegacyExpenseFields.LegacyNumber).PrimaryKey = True

        .Item(LegacyExpenseFields.DateReceived).PrimaryKey = True

        .Item(LegacyExpenseFields.Value).PrimaryKey = True
        .Item(LegacyExpenseFields.Value).SpecialColumn = True

        .SetUniqueField(LegacyExpenseFields.BequestNumber)
        .SetUniqueField(LegacyExpenseFields.DateReceived)
        .SetUniqueField(LegacyExpenseFields.Value)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "le"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "legacy_expenses"
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
    Public ReadOnly Property LegacyNumber() As Integer
      Get
        Return mvClassFields(LegacyExpenseFields.LegacyNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property DateReceived() As String
      Get
        Return mvClassFields(LegacyExpenseFields.DateReceived).Value
      End Get
    End Property
    Public ReadOnly Property Value() As Double
      Get
        Return mvClassFields(LegacyExpenseFields.Value).DoubleValue
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(LegacyExpenseFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(LegacyExpenseFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(LegacyExpenseFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property BequestNumber() As Integer
      Get
        Return mvClassFields(LegacyExpenseFields.BequestNumber).IntegerValue
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Public Overloads Sub Init(ByVal pLegacyNumber As Integer, ByVal pDateReceived As String, ByVal pValue As String)
      CheckClassFields()
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add(mvClassFields(LegacyExpenseFields.LegacyNumber).Name, pLegacyNumber)
      vWhereFields.Add(mvClassFields(LegacyExpenseFields.DateReceived).Name, mvClassFields(LegacyExpenseFields.DateReceived).FieldType, pDateReceived)
      vWhereFields.Add(mvClassFields(LegacyExpenseFields.Value).Name, mvClassFields(LegacyExpenseFields.Value).FieldType, pValue)
      MyBase.InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Overrides Sub InitForUpdate(ByVal pParams As CDBParameters)
      Init()
      Init(pParams("LegacyNumber").LongValue, pParams("OldDateReceived").Value, pParams("OldAmount").Value)
    End Sub

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.ContainsKey("Amount") Then pParameterList.Add("Value", pParameterList("Amount").Value)
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.ContainsKey("Amount") Then pParameterList.Add("Value", pParameterList("Amount").Value)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      If pParameterList.ContainsKey("Amount") Then mvClassFields(LegacyExpenseFields.Value).Value = pParameterList("Amount").Value
      If pParameterList.ContainsKey("DateReceived") Then mvClassFields(LegacyExpenseFields.DateReceived).Value = pParameterList("DateReceived").Value
    End Sub

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return "LegacyNumber,DateReceived,Amount"
    End Function

    Public Overrides Function GetUpdateKeyFieldNames() As String
      Return "LegacyNumber,OldDateReceived,OldAmount"
    End Function

    Public Overrides Function GetUniqueKeyFieldNames() As String
      Return "LegacyNumber,DateReceived,Amount"
    End Function

    Public Overrides Function GetUniqueKeyParameters() As CDBParameters
      Dim vParams As New CDBParameters
      vParams.Add(mvClassFields(LegacyExpenseFields.LegacyNumber).ProperName, mvClassFields(LegacyExpenseFields.LegacyNumber).FieldType, mvClassFields(LegacyExpenseFields.LegacyNumber).Value)
      vParams.Add(mvClassFields(LegacyExpenseFields.DateReceived).ProperName, mvClassFields(LegacyExpenseFields.DateReceived).FieldType, mvClassFields(LegacyExpenseFields.DateReceived).Value)
      vParams.Add("Amount", mvClassFields(LegacyExpenseFields.Value).FieldType, mvClassFields(LegacyExpenseFields.Value).Value)
      Return vParams
    End Function

    Public Overrides Function GetUniqueKeyFields(ByVal pParams As CDBParameters) As CDBFields
      Dim vFields As New CDBFields
      Dim vClassField As ClassField
      vClassField = mvClassFields(LegacyExpenseFields.LegacyNumber)
      vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, pParams(vClassField.ProperName).Value))
      vClassField = mvClassFields(LegacyExpenseFields.DateReceived)
      vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, pParams(vClassField.ProperName).Value))
      vClassField = mvClassFields(LegacyExpenseFields.Value)
      vFields.Add(New CDBField(vClassField.Name, vClassField.FieldType, pParams("Amount").Value))
      Return vFields
    End Function

#End Region
  End Class
End Namespace
