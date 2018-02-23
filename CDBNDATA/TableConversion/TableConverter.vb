Imports System.Data.Common
Imports CARE.Data
Imports System.IO

Public MustInherit Class TableConverter

  Private changedColumnMapField As New Dictionary(Of String, ColumnDescriptor)(StringComparer.InvariantCultureIgnoreCase)
  Private tableNameField As String = String.Empty
  Private columnListField As New List(Of ColumnDescriptor)
  Private connectionField As CDBConnection = Nothing
  Private rowsConvertedField As Integer = 0
  Private indexesField As CDBIndexes = Nothing
  Private conversionRequiredField As Boolean = False
  Private logField As StreamWriter = Nothing

  Protected Sub New(tableName As String, changedColumns As List(Of ColumnDescriptor), connection As CDBConnection, logfile As StreamWriter)
    Me.TableName = tableName
    Me.Connection = connection
    Me.Log = logfile
    Me.Log.WriteLine(String.Format("Attempting to convert table {0}...", tableName))
    Indexes = New CDBIndexes()
    Indexes.Init(connection, tableName)
    For Each changedColumn As ColumnDescriptor In changedColumns
      Me.ChangedColumnMap.Add(changedColumn.ColumnName, changedColumn)
    Next changedColumn
    For Each existingColumn As ColumnDescriptor In Me.GetExistingTableSchema
      If Me.ChangedColumnMap.ContainsKey(existingColumn.ColumnName) Then
        If Me.ChangedColumnMap(existingColumn.ColumnName) Is Nothing Then
          Throw New InvalidOperationException(String.Format("Column {0} appears to occur more than once in this table", existingColumn.ColumnName))
        End If
        Me.IsConversionRequired = Me.IsConversionRequired OrElse Not (Me.ChangedColumnMap(existingColumn.ColumnName).DataType = existingColumn.DataType)
        Me.NewSchemaWritable.Add(New ColumnDescriptor(existingColumn.ColumnName,
                                                      Me.ChangedColumnMap(existingColumn.ColumnName).DataType,
                                                      existingColumn.IsNullable))
        Me.ChangedColumnMap(existingColumn.ColumnName) = Nothing
      Else
        Me.NewSchemaWritable.Add(existingColumn)
      End If
    Next existingColumn
  End Sub

  Protected Property TableName As String
    Get
      Return tableNameField
    End Get
    Private Set(value As String)
      tableNameField = value
    End Set
  End Property

  Protected ReadOnly Property NewSchema As IList(Of ColumnDescriptor)
    Get
      Return Me.columnListField.AsReadOnly
    End Get
  End Property

  Private ReadOnly Property NewSchemaWritable As IList(Of ColumnDescriptor)
    Get
      Return Me.columnListField
    End Get
  End Property

  Private ReadOnly Property ChangedColumnMap As Dictionary(Of String, ColumnDescriptor)
    Get
      Return changedColumnMapField
    End Get
  End Property

  Protected Property Connection As CDBConnection
    Get
      Return Me.connectionField
    End Get
    Private Set(value As CDBConnection)
      Me.connectionField = value
    End Set
  End Property

  Public Property RowsConverted As Integer
    Get
      Return rowsConvertedField
    End Get
    Protected Set(value As Integer)
      rowsConvertedField = value
    End Set
  End Property

  Protected Property Indexes As CDBIndexes
    Get
      Return indexesField
    End Get
    Private Set(value As CDBIndexes)
      indexesField = value
    End Set
  End Property
  Protected Property Log As StreamWriter
    Get
      Return logField
    End Get
    Private Set(value As StreamWriter)
      logField = value
    End Set
  End Property

  Protected MustOverride Function GetExistingTableSchema() As IList(Of ColumnDescriptor)

  Public Sub ConvertTable()
    If Me.IsConversionRequired Then
      Indexes.DropAll(Me.Connection)
      DoConvertTable()
      Indexes.ReCreate(Me.Connection)
      Me.Log.WriteLine(String.Format("...schema changed " & If(Me.RowsConverted > 0, "and {0} data rows converted ", String.Empty) & "successfully" & vbCrLf, Me.RowsConverted))
    Else
      Me.Log.WriteLine("...no conversion required.")
    End If
  End Sub

  Protected MustOverride Sub DoConvertTable()

  Public Class ColumnDescriptor

    Private columnNameField As String
    Private dataTypeField As String
    Private isNullableField As Boolean

    Public Sub New(columName As String,
                   dataType As String)
      Me.New(columName,
             dataType,
             True)
    End Sub

    Public Sub New(columName As String,
                   dataType As String,
                   isNullable As Boolean)
      Me.ColumnName = columName
      Me.DataType = dataType
      Me.IsNullable = isNullable
    End Sub

    Public Property ColumnName As String
      Get
        Return Me.columnNameField
      End Get
      Private Set(value As String)
        Me.columnNameField = value
      End Set
    End Property

    Public Property DataType As String
      Get
        Return Me.dataTypeField
      End Get
      Private Set(value As String)
        Me.dataTypeField = value
      End Set
    End Property

    Public Property IsNullable As Boolean
      Get
        Return Me.isNullableField
      End Get
      Private Set(value As Boolean)
        Me.isNullableField = value
      End Set
    End Property
  End Class

  Private Property IsConversionRequired As Boolean
    Get
      Return Me.conversionRequiredField
    End Get
    Set(value As Boolean)
      Me.conversionRequiredField = value
    End Set
  End Property
End Class
