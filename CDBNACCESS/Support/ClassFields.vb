Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Linq

Namespace Access

  Partial Public Class ClassFields
    Inherits CollectionList(Of ClassField)

    Private mvTableName As String
    Private mvTableAlias As String
    Private mvUniqueFields As CollectionList(Of ClassField)
    Private mvControlNumberIndex As Integer
    Private mvControlNumberType As String
    Private mvSetPrefixRequired As Boolean
    Private mvSaveAmendedOnChanges As Boolean
    Private mvTableMaintenance As Boolean

    Public Sub New()
      MyBase.New(1)
    End Sub

    Public Property DatabaseTableName() As String
      Get
        Return mvTableName
      End Get
      Set(ByVal pValue As String)
        mvTableName = pValue
      End Set
    End Property

    Public Property TableAlias() As String
      Get
        Return mvTableAlias
      End Get
      Set(ByVal pValue As String)
        mvTableAlias = pValue
      End Set
    End Property

    Public Property SetPrefixRequired() As Boolean
      Get
        Return mvSetPrefixRequired
      End Get
      Set(ByVal pValue As Boolean)
        mvSetPrefixRequired = pValue
      End Set
    End Property

    Public ReadOnly Property TableNameAndAlias() As String
      Get
        Dim vSB As New StringBuilder
        With vSB
          .Append(mvTableName)
          .Append(" ")
          .Append(mvTableAlias)
          Return .ToString
        End With
      End Get
    End Property

    Public Overloads Sub Add(ByVal pClassField As ClassField)
      MyBase.Add(pClassField.Name, pClassField)
    End Sub

    Public Overloads Function Add(ByVal pFieldName As String) As ClassField
      Dim vClassField As New ClassField(pFieldName, CDBField.FieldTypes.cftCharacter)
      MyBase.Add(vClassField.Name, vClassField)
      If mvSetPrefixRequired Then vClassField.PrefixRequired = True
      Return vClassField
    End Function

    Public Overloads Function Add(ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes) As ClassField
      Dim vClassField As New ClassField(pFieldName, pFieldType)
      MyBase.Add(vClassField.Name, vClassField)
      If mvSetPrefixRequired Then vClassField.PrefixRequired = True
      Return vClassField
    End Function

    Public Sub SetItem(ByVal pFieldIndex As Integer, ByVal pFields As CDBFields)
      Dim vClassField As ClassField = Item(pFieldIndex)
      If pFields(vClassField.Name).FieldType = CDBField.FieldTypes.cftBinary Then
        vClassField.ByteValue = pFields(vClassField.Name).ByteValue
        'pFields(vClassField.Name).Value contains the byte converted to string but is not used
      End If
      vClassField.SetValue = pFields(vClassField.Name).Value
    End Sub

    Public Sub ClearItems()
      For Each vClassField As ClassField In Me
        vClassField.SetValue = ""
      Next
    End Sub

    Public Function FieldNames(ByVal pEnv As CDBEnvironment, ByVal pPrefix As String, ByVal pExcludeFields As List(Of ClassField)) As String
      Dim vClassFields As New ClassFields
      For Each vClassField As ClassField In Me
        If Not pExcludeFields.Contains(vClassField) Then
          vClassFields.Add(vClassField)
        End If
      Next
      Return FieldNames(pEnv, vClassFields, pPrefix)
    End Function

    Public Function FieldNames(ByVal pEnv As CDBEnvironment, ByVal pPrefix As String) As String
      Return FieldNames(pEnv, Me, pPrefix)
    End Function

    Private Function FieldNames(ByVal pEnv As CDBEnvironment, ByVal pClassFields As ClassFields, ByVal pPrefix As String) As String
      Dim vFields As New StringBuilder
      Dim vFirstField As Boolean = True

      For Each vClassField As ClassField In pClassFields
        With vClassField
          If .InDatabase AndAlso .FieldType <> CDBField.FieldTypes.cftBulk Then
            If Not vFirstField Then vFields.Append(",")
            If .SpecialColumn Then
              vFields.Append(pEnv.Connection.DBSpecialCol(pPrefix, .Name))
            Else
              If .PrefixRequired AndAlso pPrefix.Length > 0 Then
                vFields.Append(pPrefix)
                vFields.Append(".")
              End If
              vFields.Append(vClassField.Name)
            End If
            vFirstField = False
          End If
        End With
      Next
      Return vFields.ToString
    End Function

    Public Function FieldsChanged() As Boolean
      For Each vClassField As ClassField In Me
        With vClassField
          If .ValueChanged AndAlso .InDatabase Then
            If (.Name = "amended_on" OrElse .Name = "amended_by") AndAlso mvSaveAmendedOnChanges = False Then
              'ignore
            Else
              Return True
            End If
          End If
        End With
      Next
      Return False
    End Function

    Private Sub GetPrimaryKeys(ByRef pSelect1 As Integer, ByRef pSelect2 As Integer)
      For Each vClassField As ClassField In Me
        If vClassField.PrimaryKey AndAlso (vClassField.FieldType = CDBField.FieldTypes.cftInteger OrElse vClassField.FieldType = CDBField.FieldTypes.cftLong) Then
          If pSelect1 = 0 Then
            pSelect1 = vClassField.LongValue
          Else
            pSelect2 = vClassField.LongValue
            Exit For
          End If
        End If
      Next
    End Sub

    Public Function GetUniquePrimaryKey(Optional vSuppressException As Boolean = False) As ClassField
      Dim vPrimaryKey As ClassField = Nothing
      For Each vClassField As ClassField In Me
        If vClassField.PrimaryKey Then
          If vPrimaryKey Is Nothing Then
            vPrimaryKey = vClassField
          Else
            If Not vSuppressException Then
              RaiseError(DataAccessErrors.daeNoUniquePrimaryKey, mvTableName)
            End If
          End If
        End If
      Next
      If vPrimaryKey Is Nothing AndAlso vSuppressException = False Then
        RaiseError(DataAccessErrors.daeNoUniquePrimaryKey, mvTableName)
      End If
      Return vPrimaryKey
    End Function

    Public Sub SetControlNumber(ByVal pEnv As CDBEnvironment)
      If mvControlNumberIndex > 0 Then
        If Item(mvControlNumberIndex).IntegerValue = 0 Then Item(mvControlNumberIndex).IntegerValue = pEnv.GetCachedControlNumber(Me.ControlNumberType)
      End If
    End Sub

    Public Sub SetControlNumberField(ByVal pIndex As Integer, ByVal pControlNumberType As String)
      mvControlNumberIndex = pIndex
      Me.ControlNumberType = pControlNumberType
    End Sub

    Protected Friend Property ControlNumberType As String
      Get
        Return mvControlNumberType
      End Get
      Private Set(value As String)
        mvControlNumberType = value
      End Set
    End Property

    Public Sub SetSaved()
      For Each vClassField As ClassField In Me
        With vClassField
          If .ValueChanged Then .SetValue = .Value
        End With
      Next
    End Sub

    Public Sub SetUniqueField(ByVal pIndexKey As Integer)
      'unique field is used to identify the list of attributes that should be unique for this table/class
      If mvUniqueFields Is Nothing Then mvUniqueFields = New CollectionList(Of ClassField)
      If Not mvUniqueFields.ContainsKey(Item(pIndexKey).Name) Then
        mvUniqueFields.Add(Item(pIndexKey).Name, Me.Item(pIndexKey))
        'unique fields are added to a collection so that we dont have to traverse the whole list when trying to find them
      End If
      Item(pIndexKey).UniqueField = True
    End Sub

    Public Sub SetOptionalItem(ByVal pIndexKey As Integer, ByVal pFields As CDBFields)
      Dim vField As ClassField = Item(pIndexKey)
      If vField.InDatabase Then
        vField.SetValue = pFields(vField.Name).Value
      End If
    End Sub

    Friend Sub CheckRecordExists(ByVal pEnv As CDBEnvironment)
      Dim vWhereFields As New CDBFields
      Dim vAttrs As New StringBuilder
      Dim vNeedsSeparator As Boolean

      If mvUniqueFields IsNot Nothing Then
        'if we have some unique fields set
        For Each vClassField As ClassField In mvUniqueFields
          vWhereFields.Add(vClassField.Name, vClassField.FieldType, vClassField.Value)
          If vNeedsSeparator Then vAttrs.Append(",")
          vAttrs.Append(vClassField.Name)
          vNeedsSeparator = True
        Next
        If pEnv.Connection.GetCount(mvTableName, vWhereFields) > 0 Then
          RaiseError(DataAccessErrors.daeRecordExists, vAttrs.ToString)
        End If
      End If
    End Sub

    Public Function UpdateFields() As CDBFields
      Dim vUpdateFields As New CDBFields

      For Each vClassField As ClassField In Me
        With vClassField
          If .ValueChanged AndAlso vClassField.InDatabase Then
            If CBool(vClassField.FieldType = CDBField.FieldTypes.cftBinary) Then
              vUpdateFields.Add(.Name, CDBField.FieldTypes.cftBinary, .DBParam.ParameterName).DBParam = .DBParam
            Else
              vUpdateFields.Add(.Name, .FieldType, .Value)
            End If
            If .SpecialColumn Then vUpdateFields(vUpdateFields.Count).SpecialColumn = True
          End If
        End With
      Next
      Return vUpdateFields
    End Function

    Public Function WhereFields() As CDBFields
      Dim vWhereFields As New CDBFields

      For Each vClassField As ClassField In Me
        With vClassField
          If .PrimaryKey Then
            vWhereFields.Add(.Name, .FieldType, .SetValue)
            If .SpecialColumn Then vWhereFields(vWhereFields.Count).SpecialColumn = True
          End If
        End With
      Next
      If vWhereFields.Count = 0 Then
        For Each vClassField As ClassField In Me
          With vClassField
            If .InDatabase = True AndAlso .Name <> "amended_on" AndAlso .Name <> "amended_by" AndAlso .FieldType <> CDBField.FieldTypes.cftBulk AndAlso .FieldType <> CDBField.FieldTypes.cftMemo Then
              vWhereFields.Add(.Name, .FieldType, .SetValue)
              If .SpecialColumn Then vWhereFields(vWhereFields.Count).SpecialColumn = True
            End If
          End With
        Next
      End If
      Return vWhereFields
    End Function

    'Public Sub Delete(ByVal pConn As CDBConnection)
    '  Delete(pConn, Nothing, "", False, 0)
    'End Sub
    'Public Sub Delete(ByVal pConn As CDBConnection, ByVal pEnv As CDBEnvironment)
    '  Delete(pConn, pEnv, "", False, 0)
    'End Sub
    'Public Sub Delete(ByVal pConn As CDBConnection, ByVal pEnv As CDBEnvironment, ByVal pAmendedBy As String)
    '  Delete(pConn, pEnv, pAmendedBy, False, 0)
    'End Sub
    Public Sub Delete(ByVal pConn As CDBConnection, ByVal pEnv As CDBEnvironment, ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vSelect1 As Integer
      Dim vSelect2 As Integer

      If pAudit Then 'Assume numeric primary keys
        GetPrimaryKeys(vSelect1, vSelect2)
        pEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audDelete, mvTableName, vSelect1, vSelect2, pAmendedBy, Me, pJournalNumber)
      End If
      pConn.DeleteRecords(mvTableName, WhereFields)
    End Sub

    'Public Sub Save(ByVal pEnv As CDBEnvironment, ByRef pExisting As Boolean, ByVal pAmendedBy As String)
    '  Save(pEnv, pExisting, pAmendedBy, False, 0)
    'End Sub
    Public Sub Save(ByVal pEnv As CDBEnvironment, ByRef pExisting As Boolean, ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vSelect1 As Integer
      Dim vSelect2 As Integer

      If pExisting Then
        If FieldsChanged() Then
          If pAmendedBy.Length > 0 AndAlso Me.ContainsKey("amended_by") Then
            Item("amended_on").Value = TodaysDate()
            Item("amended_by").Value = pAmendedBy
          End If
          If pAudit Then 'Assume numeric primary keys
            GetPrimaryKeys(vSelect1, vSelect2)
            pEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audUpdate, mvTableName, vSelect1, vSelect2, pAmendedBy, Me, pJournalNumber)
          End If
          pEnv.Connection.UpdateRecords(mvTableName, UpdateFields, WhereFields)
        End If
      Else
        CheckRecordExists(pEnv)
        If pAudit Then 'Assume numeric primary keys
          GetPrimaryKeys(vSelect1, vSelect2)
          pEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audInsert, mvTableName, vSelect1, vSelect2, pAmendedBy, Me, pJournalNumber)
        End If
        pEnv.Connection.InsertRecord(mvTableName, UpdateFields)
        pExisting = True
      End If
      SetSaved()
    End Sub
    ''' <summary>
    ''' SAve with amendment history
    ''' </summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pExisting"></param>
    ''' <param name="pAmendedBy"></param>
    ''' <param name="pAudit"></param>
    ''' <param name="pJournalNumber"></param>
    ''' <param name="pForceAmendmentHistory"></param>
    ''' <remarks></remarks>
    Public Sub Save(ByVal pEnv As CDBEnvironment, ByRef pExisting As Boolean, ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer, ByVal pForceAmendmentHistory As Boolean)
      Dim vSelect1 As Integer
      Dim vSelect2 As Integer

      If pExisting Then
        If FieldsChanged() Then
          If pAmendedBy.Length > 0 AndAlso Me.ContainsKey("amended_by") Then
            Item("amended_on").Value = TodaysDate()
            Item("amended_by").Value = pAmendedBy
          End If
          If pAudit Then 'Assume numeric primary keys
            GetPrimaryKeys(vSelect1, vSelect2)
            pEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audUpdate, mvTableName, vSelect1, vSelect2, pAmendedBy, Me, pJournalNumber, pForceAmendmentHistory)
          End If
          pEnv.Connection.UpdateRecords(mvTableName, UpdateFields, WhereFields)
        End If
      Else
        CheckRecordExists(pEnv)
        If pAudit Then 'Assume numeric primary keys
          GetPrimaryKeys(vSelect1, vSelect2)
          pEnv.AddAmendmentHistory(CDBEnvironment.AuditTypes.audInsert, mvTableName, vSelect1, vSelect2, pAmendedBy, Me, pJournalNumber, pForceAmendmentHistory)
        End If
        pEnv.Connection.InsertRecord(mvTableName, UpdateFields)
        pExisting = True
      End If
      SetSaved()
    End Sub

    Public Function DataTableColumnNames() As String
      Dim vArray As New ArrayListEx
      For Each vClassField As ClassField In Me
        vArray.Add(ProperName(vClassField.Name))
      Next
      Return vArray.CSList
    End Function

    Public Sub AddToDataTable(ByVal pTable As CDBDataTable)
      With pTable
        Dim vRow As CDBDataRow = .AddRow
        For vIndex As Integer = 1 To Me.Count
          vRow.Item(vIndex) = Item(vIndex).Value
        Next
      End With
    End Sub

    Public Function DataTable() As CDBDataTable
      Dim vTable As New CDBDataTable
      With vTable
        For Each vClassField As ClassField In Me
          .AddColumn(ProperName(vClassField.Name), vClassField.FieldType)
        Next
        Dim vRow As CDBDataRow = .AddRow
        For vIndex As Integer = 1 To Me.Count
          vRow.Item(vIndex) = Item(vIndex).Value
        Next
      End With
      Return vTable
    End Function

    Public WriteOnly Property SaveAmendedOnChanges() As Boolean
      Set(ByVal pValue As Boolean)
        mvSaveAmendedOnChanges = pValue
      End Set
    End Property

    Public Property TableMaintenance As Boolean
      Get
        Return mvTableMaintenance
      End Get
      Set(pValue As Boolean)
        mvTableMaintenance = pValue
      End Set
    End Property

    Public Function OrderByClause(pOrderByItems As Dictionary(Of Integer, OrderByDirection)) As String
      Dim vEntries As Dictionary(Of ClassField, OrderByDirection) = pOrderByItems.ToDictionary(Of ClassField, OrderByDirection)(Function(vEntry) Me(vEntry.Key), Function(vEntry) vEntry.Value)
      Return ClassFields.OrderByClause(vEntries)
    End Function

    ''' <summary>
    ''' Builds a string that can be used as an Order By clause in a SQL Statement from a collection of ClassFields and OrderByDirection pairs
    ''' </summary>
    ''' <param name="pOrderByItems">The dictionary of ClassField and OrderByDirection items.</param>
    ''' <returns>an Order By string made up of the database name of each class field and the order by direction.</returns>
    Public Shared Function OrderByClause(pOrderByItems As Dictionary(Of ClassField, OrderByDirection)) As String
      Dim vResult As String = String.Empty
      Dim vOrderByBuilder As New StringBuilder()
      If pOrderByItems IsNot Nothing Then
        For Each vEntry As KeyValuePair(Of ClassField, OrderByDirection) In pOrderByItems
          Dim vOrderByClause As String = String.Format(", {0} {1}", vEntry.Key.Name, If(vEntry.Value = OrderByDirection.Ascending, "ASC", "DESC"))
          vOrderByBuilder.Append(vOrderByClause)
        Next
        vOrderByBuilder.Remove(0, 1) 'remove the first comma
        vResult = vOrderByBuilder.ToString()
      End If
      Return vResult
    End Function

  End Class
End Namespace