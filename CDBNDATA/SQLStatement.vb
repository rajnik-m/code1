Imports System.Linq
Namespace Data

#Region "SQLStatement class"

  Public Class SQLStatement
    Private mvConnection As CDBConnection
    Private mvFieldNames As String
    Private mvTableNames As String
    Private mvWhereFields As CDBFields
    Private mvOrderBy As String = ""
    Private mvGroupBy As String = ""
    Private mvForClause As String = ""
    Private mvDistinct As Boolean
    Private mvMaxRows As Integer
    Private mvSQL As String
    Private mvAnsiJoins As List(Of AnsiJoin)
    Private mvNullsFirst As Boolean
    Private mvUnions As List(Of UnionClause)
    Private mvTimeout As Integer
    Private mvNoLock As Boolean
    Private mvRecordSetOptions As CDBConnection.RecordSetOptions = CDBConnection.RecordSetOptions.None
    Private mvSetDecimalPlaces As Boolean
    Private mvUseAnsiSQL As Boolean         'Use Ansi SQL even if running on Oracle

    Private Enum UnionClauseType
      Union
      UnionAll
    End Enum
    Private Class UnionClause
      Public Sub New(pStatement As SQLStatement, pUnionType As UnionClauseType)
        Me.Statement = pStatement
        Me.UnionType = pUnionType
      End Sub
      Public Property Statement As SQLStatement
      Public Property UnionType As UnionClauseType
      Public ReadOnly Property UnionKeyword As String
        Get
          Return If(UnionType = UnionClauseType.UnionAll, " UNION ALL ", " UNION ")
        End Get
      End Property
    End Class

    ''' <summary>
    ''' DO NOT USE THIS CONSTRUCTOR
    ''' </summary>
    ''' <param name="pConn"></param>
    ''' <param name="pSQL"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pConn As CDBConnection, ByVal pSQL As String)
      mvConnection = pConn
      mvSQL = pSQL
    End Sub
    ''' <summary>
    ''' DO NOT USE THIS CONSTRUCTOR
    ''' </summary>
    ''' <param name="pConn"></param>
    ''' <param name="pSQL"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pConn As CDBConnection, ByVal pSQL As String, ByVal pWhereFields As CDBFields)
      mvConnection = pConn
      mvSQL = pSQL
      mvWhereFields = pWhereFields
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableNames As String)
      Init(pConn, pFieldNames, pTableNames, Nothing, "", Nothing, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableNames As String, ByVal pWhereField As CDBField)
      Init(pConn, pFieldNames, pTableNames, New CDBFields(pWhereField), "", Nothing, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableNames As String, ByVal pWhereField As CDBField, ByVal pOrderBy As String)
      Init(pConn, pFieldNames, pTableNames, New CDBFields(pWhereField), pOrderBy, Nothing, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableNames As String, ByVal pWhereFields As CDBFields)
      Init(pConn, pFieldNames, pTableNames, pWhereFields, "", Nothing, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableNames As String, ByVal pWhereFields As CDBFields, ByVal pOrderBy As String)
      Init(pConn, pFieldNames, pTableNames, pWhereFields, pOrderBy, Nothing, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableName As String, ByVal pWhereFields As CDBFields, ByVal pOrderBy As String, ByVal pAnsiJoins As List(Of AnsiJoin))
      Init(pConn, pFieldNames, pTableName, pWhereFields, pOrderBy, pAnsiJoins, False)
    End Sub
    Public Sub New(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableName As String, ByVal pWhereFields As CDBFields, ByVal pOrderBy As String, ByVal pAnsiJoins As List(Of AnsiJoin), ByVal pNullsFirst As Boolean)
      Init(pConn, pFieldNames, pTableName, pWhereFields, pOrderBy, pAnsiJoins, pNullsFirst)
    End Sub

    Public Sub Init(ByVal pConn As CDBConnection, ByVal pFieldNames As String, ByVal pTableName As String, ByVal pWhereFields As CDBFields, ByVal pOrderBy As String, ByVal pAnsiJoins As List(Of AnsiJoin), ByVal pNullsFirst As Boolean)
      mvConnection = pConn
      mvFieldNames = pFieldNames
      mvTableNames = pTableName
      mvWhereFields = pWhereFields
      mvOrderBy = pOrderBy
      mvAnsiJoins = pAnsiJoins
      mvNullsFirst = pNullsFirst
    End Sub

    Public Sub SetOrderByClientDeptLogname(ByVal pOtherField As String)
      If mvConnection.NullsSortAtEnd Then
        mvOrderBy = "logname,department,client," & pOtherField
      Else
        mvOrderBy = "logname DESC,department DESC,client DESC," & pOtherField
      End If
    End Sub

    Public Function GetDataTable() As DataTable
      Return mvConnection.GetDataTable(Me)
    End Function

    Public Function GetDataSet() As DataSet
      Return mvConnection.GetDataSet(Me)
    End Function

    Public Function GetRecordSet() As CDBRecordSet
      Dim vRecordSet As CDBRecordSet = mvConnection.GetRecordSet(Me, mvTimeout, mvRecordSetOptions)
      vRecordSet.SetDecimalPlaces = mvSetDecimalPlaces
      Return vRecordSet
    End Function

    Public Function GetRecordSet(ByVal pOptions As CDBConnection.RecordSetOptions) As CDBRecordSet
      Dim vRecordSet As CDBRecordSet = mvConnection.GetRecordSet(Me, mvTimeout, pOptions)
      vRecordSet.SetDecimalPlaces = mvSetDecimalPlaces
      Return vRecordSet
    End Function

    Public Function GetIntegerValue() As Integer
      Return IntegerValue(GetValue())
    End Function

    Public Function GetValue() As String
      Dim vValue As String = ""
      Dim vRecordSet As CDBRecordSet = mvConnection.GetRecordSet(Me, mvTimeout)
      If vRecordSet.Fetch() = True Then
        vValue = vRecordSet.Fields(1).Value
      End If
      vRecordSet.CloseRecordSet()
      Return vValue
    End Function

    Public Function GetValues() As List(Of String)
      Dim vValues As New List(Of String)
      Dim vRecordSet As CDBRecordSet = mvConnection.GetRecordSet(Me, mvTimeout)
      While vRecordSet.Fetch
        vValues.Add(vRecordSet.Fields(1).Value)
      End While
      vRecordSet.CloseRecordSet()
      Return vValues
    End Function

    Private Function GetCountStatement() As String
      Dim vSQL As New StringBuilder(mvConnection.GetSelectSQLCSC)
      vSQL.Append("Count(*) AS record_count FROM ")
      vSQL.Append(mvTableNames)
      If mvAnsiJoins IsNot Nothing AndAlso mvAnsiJoins.Count > 0 Then
        AddAnsiJoins(vSQL)
      End If
      If mvWhereFields IsNot Nothing AndAlso mvWhereFields.Count > 0 Then
        vSQL.Append(" WHERE ")
        vSQL.Append(mvConnection.WhereClause(mvWhereFields))
      End If
      Return vSQL.ToString
    End Function

    Private Function GetSQLStatement() As String
      Dim vSQL As New StringBuilder
      If mvSQL IsNot Nothing Then
        vSQL.Append(mvSQL)
      Else
        vSQL.Append(mvConnection.GetSelectSQLCSC)
        If mvDistinct Then vSQL.Append("DISTINCT ")
        If mvMaxRows > 0 AndAlso mvConnection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseTopN Then
          vSQL.Append("TOP ")
          vSQL.Append(mvMaxRows)
          vSQL.Append(" ")
        End If
        vSQL.Append(mvFieldNames)
        vSQL.Append(" FROM ")
        If mvTableNames.Length = 0 Then vSQL.Append("( ")
        vSQL.Append(mvTableNames)
        If mvNoLock Then vSQL.Append(" WITH(NOLOCK) ")
        If mvAnsiJoins IsNot Nothing AndAlso mvAnsiJoins.Count > 0 Then
          AddAnsiJoins(vSQL)
        End If
      End If
      If mvWhereFields IsNot Nothing AndAlso mvWhereFields.Count > 0 Then
        vSQL.Append(" WHERE ")
        vSQL.Append(mvConnection.WhereClause(mvWhereFields))
      End If
      If mvMaxRows > 0 AndAlso mvConnection.RowRestrictionType = CDBConnection.RowRestrictionTypes.UseRownum Then
        If mvWhereFields IsNot Nothing AndAlso mvWhereFields.Count > 0 Then
          vSQL.Append(" AND ")
        Else
          vSQL.Append(" WHERE ")
        End If
        vSQL.Append(" rownum < ")
        vSQL.Append(mvMaxRows + 1)
      End If

      Dim vUnionSQL As String = GenerateUnionSQL()
      vSQL.AppendLine()
      vSQL.AppendLine(vUnionSQL)

      If String.IsNullOrEmpty(mvTableNames) AndAlso String.IsNullOrEmpty(mvSQL) Then vSQL.Append(") result")
      If mvGroupBy.Length > 0 Then
        vSQL.Append(" GROUP BY ")
        vSQL.Append(mvGroupBy)
      End If
      If mvOrderBy.Length > 0 Then
        vSQL.Append(" ORDER BY ")
        Dim vOrderBy As New StringBuilder
        If mvNullsFirst AndAlso mvConnection.NullsSortAtEnd Then
          Dim vOrderItems() As String = mvOrderBy.Split(","c)
          For Each vItem As String In vOrderItems
            If vOrderBy.Length > 0 Then vOrderBy.Append(",")
            vOrderBy.Append(vItem)
            If Not vItem.Contains(" DESC") Then vOrderBy.Append(" Nulls First")
          Next
          vSQL.Append(vOrderBy.ToString)
        Else
          vSQL.Append(mvOrderBy)
        End If
      End If
      If mvForClause.Length > 0 Then
        vSQL.Append(" ")
        vSQL.Append(mvForClause)
      End If
      Return vSQL.ToString
    End Function

    Private Property UnionClauses As List(Of UnionClause)
      Get
        If mvUnions Is Nothing Then
          Me.UnionClauses = New List(Of UnionClause)
        End If
        Return mvUnions
      End Get
      Set(value As List(Of UnionClause))
        mvUnions = value
      End Set
    End Property

    Private Function GenerateUnionSQL() As String
      Dim vSQL As New StringBuilder()
      Dim vAddUnion As Boolean
      For Each vUnionEntry As UnionClause In Me.UnionClauses
        If mvTableNames.Length > 0 OrElse vAddUnion Then
          vSQL.Append(vbCrLf)
          vSQL.Append(vUnionEntry.UnionKeyword)
        End If
        vSQL.Append(vbCrLf)
        vSQL.Append(vUnionEntry.Statement.SQL)
        vAddUnion = True
      Next
      Return vSQL.ToString()
    End Function


    ''' <summary>
    ''' DO NOT USE THIS IN CODE.  Returns the SQL property with carriage returns to make it more readable when pasting into a SQL Query window
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property DEBUG_SQL() As String
      Get
        Dim vSQL As String = GetSQLStatement()
        vSQL = vSQL.Replace(" FROM ", Environment.NewLine + " FROM ")
        vSQL = vSQL.Replace(" INNER JOIN ", Environment.NewLine + vbTab + "  INNER JOIN ")
        vSQL = vSQL.Replace(" LEFT OUTER JOIN ", Environment.NewLine + vbTab + "   LEFT OUTER JOIN ")
        vSQL = vSQL.Replace(" RIGHT OUTER JOIN ", Environment.NewLine + vbTab + "  RIGHT OUTER JOIN ")
        vSQL = vSQL.Replace(" ON ", Environment.NewLine + vbTab + vbTab + " ON ")
        vSQL = vSQL.Replace(" WHERE ", Environment.NewLine + " WHERE ")
        vSQL = vSQL.Replace(" AND ", Environment.NewLine + vbTab + "   AND ")
        vSQL = vSQL.Replace(" GROUP BY ", Environment.NewLine + " GROUP BY ")
        vSQL = vSQL.Replace(" ORDER BY ", Environment.NewLine + " ORDER BY ")
        vSQL = vSQL.Replace(" UNION ", Environment.NewLine + Environment.NewLine + " UNION " + Environment.NewLine + Environment.NewLine)
        vSQL = vSQL.Replace(" UNION " + Environment.NewLine + Environment.NewLine + "ALL ", " UNION ALL " + Environment.NewLine + Environment.NewLine)
        vSQL = vSQL.Replace(" (SELECT ", Environment.NewLine + vbTab + vbTab + " (SELECT ")
        vSQL = vSQL.Replace(" ( SELECT ", Environment.NewLine + vbTab + vbTab + " ( SELECT ")
        vSQL = "--For debug use only" + Environment.NewLine + vSQL
        Return vSQL
      End Get
    End Property

    Private Sub AddAnsiJoins(ByVal pSQL As StringBuilder)
      For Each vAnsiJoin As AnsiJoin In mvAnsiJoins
        With pSQL
          Select Case vAnsiJoin.AnsiJoinType
            Case AnsiJoin.AnsiJoinTypes.InnerJoin
              .Append(" INNER JOIN ")
            Case AnsiJoin.AnsiJoinTypes.LeftOuterJoin
              .Append(" LEFT OUTER JOIN ")
            Case AnsiJoin.AnsiJoinTypes.RightOuterJoin
              .Append(" RIGHT OUTER JOIN ")
          End Select
          .Append(vAnsiJoin.TableName)
          If mvNoLock Then .Append(" WITH (NOLOCK)")
          .Append(" ON ")
          Dim vAnd As Boolean = False
          For Each vJoin As JoinFields In vAnsiJoin.Joins
            If vAnd Then .Append(" AND ")
            .Append(vJoin.Attribute1)
            If Not String.IsNullOrWhiteSpace(vJoin.Attribute2) Then
              .Append(" = ")
              .Append(vJoin.Attribute2)
            Else
              .Append(" IS NULL ")
            End If
            vAnd = True
          Next
        End With
      Next
    End Sub

    Public Sub AddUnion(ByVal pSQL As SQLStatement)
      Dim vEntry As New UnionClause(pSQL, UnionClauseType.Union)
      Me.UnionClauses.Add(vEntry)
    End Sub
    Public Sub AddUnionAll(ByVal pSQL As SQLStatement)
      Dim vEntry As New UnionClause(pSQL, UnionClauseType.UnionAll)
      Me.UnionClauses.Add(vEntry)
    End Sub

    Public ReadOnly Property CountSQL() As String
      Get
        Return GetCountStatement()
      End Get
    End Property

    Public ReadOnly Property SQL() As String
      Get
        If mvAnsiJoins IsNot Nothing AndAlso mvAnsiJoins.Count > 0 AndAlso mvUseAnsiSQL = False Then
          Return mvConnection.ProcessAnsiJoins(GetSQLStatement)
        Else
          Return GetSQLStatement()
        End If
      End Get
    End Property

    Public Property FieldNames() As String
      Get
        Return mvFieldNames
      End Get
      Set(ByVal pValue As String)
        mvFieldNames = pValue
      End Set
    End Property

    Public Property OrderBy() As String
      Get
        Return mvOrderBy
      End Get
      Set(ByVal pValue As String)
        mvOrderBy = pValue
      End Set
    End Property

    Public Property UseAnsiSQL() As Boolean
      Get
        Return mvUseAnsiSQL
      End Get
      Set(ByVal pValue As Boolean)
        mvUseAnsiSQL = pValue
      End Set
    End Property

    Public Property GroupBy() As String
      Get
        Return mvGroupBy
      End Get
      Set(ByVal pValue As String)
        mvGroupBy = pValue
      End Set
    End Property

    Public Property ForClause() As String
      Get
        Return mvForClause
      End Get
      Set(ByVal pValue As String)
        mvForClause = pValue
      End Set
    End Property

    Public Property Distinct() As Boolean
      Get
        Return mvDistinct
      End Get
      Set(ByVal value As Boolean)
        mvDistinct = value
      End Set
    End Property

    Public Property NoLock() As Boolean
      Get
        Return mvNoLock
      End Get
      Set(ByVal value As Boolean)
        If mvConnection.SupportsNoLock Then
          mvNoLock = value
        Else
          mvNoLock = False
        End If
      End Set
    End Property

    Public Property MaxRows() As Integer
      Get
        Return mvMaxRows
      End Get
      Set(ByVal value As Integer)
        mvMaxRows = value
      End Set
    End Property

    Public Property Timeout() As Integer
      Get
        Return mvTimeout
      End Get
      Set(ByVal value As Integer)
        mvTimeout = value
      End Set
    End Property

    Public Property RecordSetOptions As CDBConnection.RecordSetOptions
      Get
        Return mvRecordSetOptions
      End Get
      Set(ByVal pValue As CDBConnection.RecordSetOptions)
        mvRecordSetOptions = pValue
      End Set
    End Property

    Public Property SetDecimalPlaces As Boolean
      Get
        Return mvSetDecimalPlaces
      End Get
      Set(ByVal pValue As Boolean)
        mvSetDecimalPlaces = pValue
      End Set
    End Property

    Public ReadOnly Property WhereFields As CDBFields
      Get
        Return mvWhereFields
      End Get
    End Property



    Public ReadOnly Property JoinFields As List(Of AnsiJoin)
      Get
        Return mvAnsiJoins
      End Get
    End Property

    ''' <summary>
    ''' Builds a Where clause to check if two periods of time overlap.  Two periods of time are said to overlap if one period's starting point falls within the other period
    ''' or the other period's starting point falls within the first period.
    ''' Note that null or empty values will be considered infinite, i.e. replaced by their minimum value for the period's starting point or maximum value for ending point
    ''' </summary>
    ''' <param name="pStartDateField">A CDBField or ClassField that represents the start point.  The ClassField and CDBField class both have widening operators that turn one into the other</param>
    ''' <param name="pEndDateField">A CDBField or ClassField that represents the end point</param>
    ''' <returns></returns>
    Public Shared Function BuildOverlappingWhere(pConn As CDBConnection, pStartDateField As CDBField, pEndDateField As CDBField, Optional pWhereFields As CDBFields = Nothing) As CDBFields

      If pWhereFields Is Nothing Then
        pWhereFields = New CDBFields()
      End If

      Dim NonEmptyStringSearch As Func(Of String, Boolean) = Function(vString) Not String.IsNullOrWhiteSpace(vString) 'Create a generic function that returns True if the string isn't null.  We'll use it in array searches to return the first non-empty string

      'NB this has to work the same in both ORACLE and SQL Server.  When using COALESCE, in ORACLE it is not possible to use dates as strings, so you have to convert everything to a date.  SQL Server handles it fine.
      Dim vMinDate As String = pConn.DBToDate(String.Format("'{0}'", DateTime.MinValue.ToString(CAREDateFormat)))
      Dim vMaxDate As String = pConn.DBToDate(String.Format("'{0}'", DateTime.MaxValue.ToString(CAREDateFormat)))
      Dim vMyStartDate As String = {pStartDateField.Value, DateTime.MinValue.ToString(CAREDateFormat)}.FirstOrDefault(NonEmptyStringSearch)
      Dim vMyQuotedStartDate As String = String.Format(pConn.DBToDate("'{0}'"), vMyStartDate)
      Dim vMyStartDateColumn As String = String.Format("COALESCE({0}, {1})", pConn.DBDateTimeAttribToDate(pStartDateField.Name), vMinDate) 'NB DBToDate is needed to convert DateTime columns to Date types, as the Date type can handle 1st Jan 0001 but DateTime can't
      Dim vTheirStartDateColumn As String = String.Format("COALESCE({0}, {1})", pConn.DBDateTimeAttribToDate(pStartDateField.Name), vMinDate)
      Dim vTheirEndDateColumn As String = String.Format("COALESCE({0}, {1})", pConn.DBDateTimeAttribToDate(pEndDateField.Name), vMaxDate)

      'clause 1: where another record between my start date
      pWhereFields.Add(vMyStartDateColumn, pStartDateField.FieldType, vMyStartDate,
                          CDBField.FieldWhereOperators.fwoBetweenFrom Or CDBField.FieldWhereOperators.fwoOpenBracket)
      'clause 1: ...and my end date
      pWhereFields.Add(String.Format("{0}_2", vMyStartDateColumn), pStartDateField.FieldType,
                       {pEndDateField.Value, DateTime.MaxValue.ToString(CAREDateFormat)}.FirstOrDefault(NonEmptyStringSearch),
                          CDBField.FieldWhereOperators.fwoBetweenTo)

      'clause 2: or my start date between another record's start date
      pWhereFields.Add(vMyQuotedStartDate, CDBField.FieldTypes.cftUnknown, vTheirStartDateColumn,
                         CDBField.FieldWhereOperators.fwoBetweenFrom Or CDBField.FieldWhereOperators.fwoOR)
      'clause 2: ...and end date
      pWhereFields.Add(String.Format("{0}_2", vMyQuotedStartDate), CDBField.FieldTypes.cftUnknown, vTheirEndDateColumn,
                         CDBField.FieldWhereOperators.fwoBetweenTo Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Return pWhereFields
    End Function

    Public Sub BuildOverlappingWhere(pStartDateField As CDBField, pEndDateField As CDBField)
      SQLStatement.BuildOverlappingWhere(Me.Connection, pStartDateField, pEndDateField, Me.WhereFields)
    End Sub

    Public ReadOnly Property Connection As CDBConnection
      Get
        Return mvConnection
      End Get
    End Property

  End Class

#End Region

#Region "AnsiJoins class"
  Public Class AnsiJoins
    Inherits List(Of AnsiJoin)

    Public Sub New()
      MyBase.New()
    End Sub

    ''' <summary>
    ''' Create a new collection of joins from the an array of joins.
    ''' </summary>
    ''' <param name="pJoins">An array containing the <see cref="AnsiJoin"/> objects that should form the
    ''' initial content of the collection.</param>
    ''' <remarks></remarks>
    Public Sub New(pJoins As AnsiJoin())
      MyBase.New()
      For Each vJoin As AnsiJoin In pJoins
        Me.Add(vJoin)
      Next vJoin
    End Sub

    Public Overloads Sub Add(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2))
    End Sub

    Public Overloads Sub Add(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, pAttr3, pAttr4))
    End Sub

    Public Overloads Sub Add(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String, ByVal pAttr5 As String, ByVal pAttr6 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, pAttr3, pAttr4, pAttr5, pAttr6, AnsiJoin.AnsiJoinTypes.InnerJoin))
    End Sub

    Public Overloads Sub Add(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAnsiJoinType As AnsiJoin.AnsiJoinTypes)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, pAnsiJoinType))
    End Sub

    Public Sub AddLeftOuterJoin(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, AnsiJoin.AnsiJoinTypes.LeftOuterJoin))
    End Sub

    Public Sub AddLeftOuterJoin(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, pAttr3, pAttr4, AnsiJoin.AnsiJoinTypes.LeftOuterJoin))
    End Sub

    Public Sub AddLeftOuterJoin(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String, ByVal pAttr5 As String, ByVal pAttr6 As String)
      Add(New AnsiJoin(pTableName, pAttr1, pAttr2, pAttr3, pAttr4, pAttr5, pAttr6, AnsiJoin.AnsiJoinTypes.LeftOuterJoin))
    End Sub

    Public Function ContainsJoinToTable(ByVal pTableName As String) As Boolean
      For Each vJoin As AnsiJoin In Me
        If vJoin.TableName = pTableName Then Return True
      Next
    End Function

    Public Function ContainsAnyJoinToTable(ByVal pTableName As String) As Boolean
      For Each vJoin As AnsiJoin In Me
        If vJoin.TableName.Contains(pTableName) Then Return True
      Next
    End Function

    Public Sub RemoveJoin(ByVal pTableName As String)
      For Each vJoin As AnsiJoin In Me
        If vJoin.TableName = pTableName Then
          Me.Remove(vJoin)
          Exit For
        End If
      Next
    End Sub

  End Class
#End Region

#Region "AnsiJoin class"

  Public Class AnsiJoin
    Public Enum AnsiJoinTypes
      InnerJoin
      LeftOuterJoin
      RightOuterJoin
    End Enum

    Private mvJoinType As AnsiJoinTypes
    Private mvTableName As String
    Private mvJoinFields As New List(Of JoinFields)

    Public Sub New(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String)
      mvTableName = pTableName
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
      mvJoinType = AnsiJoinTypes.InnerJoin
    End Sub

    Public Sub New(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAnsiJoinType As AnsiJoinTypes)
      mvTableName = pTableName
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
      mvJoinType = pAnsiJoinType
    End Sub

    Public Sub New(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String)
      mvTableName = pTableName
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
      mvJoinFields.Add(New JoinFields(pAttr3, pAttr4))
      mvJoinType = AnsiJoinTypes.InnerJoin
    End Sub

    Public Sub New(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String, ByVal pAnsiJoinType As AnsiJoinTypes)
      mvTableName = pTableName
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
      mvJoinFields.Add(New JoinFields(pAttr3, pAttr4))
      mvJoinType = pAnsiJoinType
    End Sub

    Public Sub New(ByVal pTableName As String, ByVal pAttr1 As String, ByVal pAttr2 As String, ByVal pAttr3 As String, ByVal pAttr4 As String, ByVal pAttr5 As String, ByVal pAttr6 As String, ByVal pAnsiJoinType As AnsiJoinTypes)
      mvTableName = pTableName
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
      mvJoinFields.Add(New JoinFields(pAttr3, pAttr4))
      mvJoinFields.Add(New JoinFields(pAttr5, pAttr6))
      mvJoinType = pAnsiJoinType
    End Sub

    Public Sub AddJoinFields(ByVal pAttr1 As String, ByVal pAttr2 As String)
      mvJoinFields.Add(New JoinFields(pAttr1, pAttr2))
    End Sub

    Friend ReadOnly Property AnsiJoinType() As AnsiJoinTypes
      Get
        Return mvJoinType
      End Get
    End Property

    Public ReadOnly Property TableName() As String
      Get
        Return mvTableName
      End Get
    End Property
    Public ReadOnly Property Joins() As List(Of JoinFields)
      Get
        Return mvJoinFields
      End Get
    End Property
  End Class
#End Region

#Region "JoinFields Class"

  Public Class JoinFields
    Private mvAttr1 As String
    Private mvAttr2 As String

    Public Sub New(ByVal pAttr1 As String, ByVal pAttr2 As String)
      mvAttr1 = pAttr1
      mvAttr2 = pAttr2
    End Sub
    Public ReadOnly Property Attribute1() As String
      Get
        Return mvAttr1
      End Get
    End Property
    Public ReadOnly Property Attribute2() As String
      Get
        Return mvAttr2
      End Get
    End Property
  End Class

#End Region

End Namespace
