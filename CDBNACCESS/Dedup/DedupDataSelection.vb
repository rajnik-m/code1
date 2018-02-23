Imports System.Reflection
Imports System.Linq

Namespace Access.Deduplication

  Public Class DedupDataSelection
    Inherits DataSelection


    Public Shadows Enum DataSelectionTypes As Integer
      None = 7000
      Contacts
      Uniserv
    End Enum

    Public Property Parameters As CDBParameters
      Get
        Return mvParameters
      End Get
      Private Set(value As CDBParameters)
        mvParameters = value
      End Set
    End Property

    Public ReadOnly Property Environment As CDBEnvironment
      Get
        Return mvEnv
      End Get
    End Property


    Public Property DataFactory As IDedupDataGenerator

    Public Sub New(pEnv As CDBEnvironment, pParams As CDBParameters, pListType As DataSelectionListType, pDataFactory As IDedupDataGenerator)
      Me.DataFactory = pDataFactory
      Me.DataFactory.Parent = Me
      Init(pEnv, pParams, pListType)
    End Sub

    Public Sub New(pEnv As CDBEnvironment, pParams As CDBParameters, pListType As DataSelectionListType, pGeneratorType As DedupDataSelection.DataSelectionTypes)
      Me.DataFactory = DedupDataGeneratorFactory.GetDedupDataGenerator(pEnv, pGeneratorType)
      If Me.DataFactory Is Nothing Then
        Throw New TypeLoadException(String.Format("No Data Generator exists for type {0}", pGeneratorType.ToString()))
      End If
      Me.DataFactory.Parent = Me
      Init(pEnv, pParams, pListType)
    End Sub

    Protected Overridable Overloads Sub Init(ByVal pEnv As CDBEnvironment, pParams As CDBParameters, pListType As DataSelectionListType)
      mvEnv = pEnv
      mvParameters = pParams


      mvAvailableUsages = DataSelectionUsages.dsuWEBServices Or DataSelectionUsages.dsuSmartClient
      mvDisplayListItems = Nothing
      mvDataSelectionListType = pListType

      mvResultColumns = DataFactory.ResultColumns
      mvSelectColumns = DataFactory.SelectColumns
      mvHeadings = DataFactory.Headings
      mvRequiredItems = DataFactory.RequiredItems

      mvCode = DataFactory.Code

      Dim vGroupCode As String = String.Empty
      If mvParameters IsNot Nothing AndAlso mvParameters.Exists("ContactGroup") Then
        vGroupCode = mvParameters("ContactGroup").Value
      End If
      mvWidths = DataFactory.Widths

      If pListType = DataSelectionListType.dsltEditing Then
        mvResultColumns = mvResultColumns & ",DetailItems,NewColumn,NewColumn2,NewColumn3,Spacer"
      End If

      Select Case pListType
        Case DataSelectionListType.dsltUser
          ReadUserDisplayListItems(mvEnv.User.Department, mvEnv.User.Logname, vGroupCode, DataSelectionUsages.dsuSmartClient)
      End Select

    End Sub

    Private Sub ApplyOwnershipAccessSecurity(pDataTable As CDBDataTable, pAccessLevelColumn As String, Optional vAllowedCols As List(Of String) = Nothing, Optional ByVal pBrowseAccessLevelIdentifier As String = "B")
      If vAllowedCols Is Nothing Then vAllowedCols = New List(Of String)
      If pDataTable IsNot Nothing AndAlso pDataTable.Columns.ContainsKey(pAccessLevelColumn) Then
        For Each vRow As CDBDataRow In pDataTable.Rows
          If vRow.Item(pAccessLevelColumn) = pBrowseAccessLevelIdentifier Then
            For vIdx As Integer = 1 To pDataTable.Columns.Count
              If Not vAllowedCols.Contains(pDataTable.Columns(vIdx).Name) Then
                vRow.Item(vIdx) = Nothing
              End If
            Next
          End If
        Next
      End If
    End Sub

    Public Shadows Function DataTable(pRule As DedupRule) As CDBDataTable
      Dim vResult As New CDBDataTable()

      If mvResultColumns.Length > 0 Then vResult.AddColumnsFromList(mvResultColumns)

      If pRule.Clauses.Count > 0 AndAlso pRule.Clauses.All(Function(pClause) mvParameters.HasValue(pClause.Parameter)) Then
        Dim vSQL As SQLStatement = DataFactory.GenerateSQLStatement(pRule)
        If Me.Parameters.HasValue("NumberOfRows") Then
          vSQL.MaxRows = Me.Parameters("NumberOfRows").IntegerValue
        End If
        vResult.FillFromSQL(mvEnv, vSQL, String.Empty, "", True)
      End If

      Return vResult
    End Function

    Public Shared Function CreateJoins(pDedupClause As DedupClause) As List(Of AnsiJoin)
      Dim vResult As New List(Of AnsiJoin)
      If pDedupClause.Joins.Count > 0 Then
        For Each vDedupJoin As DedupClause.Join In pDedupClause.Joins
          Dim vJoin As AnsiJoin = vResult.FirstOrDefault(Function(pJoin As AnsiJoin) pJoin.TableName = vDedupJoin.TableAndAlias)
          If vJoin Is Nothing Then
            vJoin = New AnsiJoin(vDedupJoin.TableAndAlias, vDedupJoin.AnchorPart, vDedupJoin.JoinPart)
            vResult.Add(vJoin)
          Else
            vJoin.AddJoinFields(vDedupJoin.AnchorPart, vDedupJoin.JoinPart)
          End If
        Next
      End If
      Return vResult
    End Function
    Public Shared Function CreateWhere(pDedupClause As DedupClause, pValue As Object, pConnection As CDBConnection) As CDBField
      Dim vResult As CDBField = Nothing
      Dim vSeparator As String = If(String.IsNullOrWhiteSpace(pDedupClause.TableAlias), String.Empty, ".")
      vResult = New CDBField(String.Format("{0}{1}{2}", pDedupClause.TableAlias,
                                                        vSeparator,
                                                        If(pConnection.IsSpecialColumn(pDedupClause.Attribute), pConnection.DBSpecialCol(pDedupClause.Attribute), pDedupClause.Attribute)))
      Dim vMatchPattern As String = "{0}"
      Select Case pDedupClause.Match
        Case DedupMatch.Contains
          vResult.WhereOperator = CDBField.FieldWhereOperators.fwoLike
          vMatchPattern = "%{0}%"
        Case DedupMatch.IsLike
          vResult.WhereOperator = CDBField.FieldWhereOperators.fwoLike
          vMatchPattern = "{0}%"
        Case Else
          vResult.WhereOperator = CDBField.FieldWhereOperators.fwoEqual
      End Select
      vResult.Value = String.Format(vMatchPattern, pValue)

      Return vResult
    End Function

  End Class

End Namespace
