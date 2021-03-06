Namespace Access

  Public Class BranchPostcode
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum BranchPostcodeFields
      AllFields = 0
      Branch
      OutwardPostcode
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("branch")
        .Add("outward_postcode")

        .Item(BranchPostcodeFields.OutwardPostcode).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "bp"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "branch_postcodes"
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
    Public ReadOnly Property Branch() As String
      Get
        Return mvClassFields(BranchPostcodeFields.Branch).Value
      End Get
    End Property
    Public ReadOnly Property OutwardPostcode() As String
      Get
        Return mvClassFields(BranchPostcodeFields.OutwardPostcode).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(BranchPostcodeFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(BranchPostcodeFields.AmendedBy).Value
      End Get
    End Property
#End Region

#Region "Public Methods"

    Public Sub MoveBranch(ByVal pNewBranch As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vSQL As String

      vUpdateFields.Add("branch", CDBField.FieldTypes.cftCharacter, pNewBranch)
      vUpdateFields.AddAmendedOnBy(mvEnv.User.Logname)
			vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, Branch)
			vWhereFields.Add("postcode", CDBField.FieldTypes.cftCharacter, OutwardPostcode, CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
			vWhereFields.Add("postcode#2", CDBField.FieldTypes.cftCharacter, OutwardPostcode & " " & "*", CDBField.FieldWhereOperators.fwoLike Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      mvEnv.Connection.UpdateRecords("addresses", vUpdateFields, vWhereFields, False)

      vSQL = "UPDATE branch_income SET branch_code = '" & pNewBranch & "'"
      vSQL = vSQL & " WHERE branch_code = '" & Branch & "' AND order_number IN"
			vSQL = vSQL & " (SELECT order_number FROM orders o,addresses a WHERE o.order_number = branch_income.order_number AND a.address_number = o.address_number AND (a.postcode = '" & OutwardPostcode & "' OR a.postcode " & mvEnv.Connection.DBLike(OutwardPostcode & " " & "*") & "))"
      mvEnv.Connection.ExecuteSQL(vSQL)

      vSQL = "UPDATE orders SET branch = '" & pNewBranch & "', amended_on " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftDate, TodaysDate) & ", amended_by = '" & mvEnv.User.Logname & "'"
      vSQL = vSQL & " WHERE branch = '" & Branch & "' AND address_number IN"
			vSQL = vSQL & " (SELECT address_number FROM addresses a WHERE a.address_number = orders.address_number AND (a.postcode = '" & OutwardPostcode & "' OR a.postcode " & mvEnv.Connection.DBLike(OutwardPostcode & " " & "*") & "))"
      mvEnv.Connection.ExecuteSQL(vSQL)

      vSQL = "UPDATE members SET branch = '" & pNewBranch & "', amended_on " & mvEnv.Connection.SQLLiteral("=", CDBField.FieldTypes.cftDate, TodaysDate) & ", amended_by = '" & mvEnv.User.Logname & "'"
      vSQL = vSQL & " WHERE branch = '" & Branch & "' AND address_number IN"
			vSQL = vSQL & " (SELECT address_number FROM addresses a WHERE a.address_number = members.address_number AND (a.postcode = '" & OutwardPostcode & "' OR a.postcode " & mvEnv.Connection.DBLike(OutwardPostcode & " " & "*") & "))"
      mvEnv.Connection.ExecuteSQL(vSQL)

      mvClassFields(BranchPostcodeFields.Branch).Value = pNewBranch
      Save()
    End Sub

#End Region

  End Class
End Namespace
