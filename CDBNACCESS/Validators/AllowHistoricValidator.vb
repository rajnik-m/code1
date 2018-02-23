Public Class AllowHistoricValidator : Implements IValidator

  Private mvEnvironment As CDBEnvironment
  Private mvTable As String
  Private mvAllowHistoricOnly As Boolean
  Private mvIsHistoricColumnName As String
  Private mvFieldValues As CDBFields

  Public Property Table As String
    Get
      Return mvTable
    End Get
    Private Set(value As String)
      mvTable = value
    End Set
  End Property
  Public Property IsHistoricColumnName As String
    Get
      Return mvIsHistoricColumnName
    End Get
    Private Set(value As String)
      mvIsHistoricColumnName = value
    End Set
  End Property
  Public Property AllowHistoric As Boolean
    Get
      Return mvAllowHistoricOnly
    End Get
    Private Set(value As Boolean)
      mvAllowHistoricOnly = value
    End Set
  End Property
  Public Property FieldValues As CDBFields
    Get
      Return mvFieldValues
    End Get
    Private Set(value As CDBFields)
      mvFieldValues = value
    End Set
  End Property
  Private Property Environment As CDBEnvironment
    Get
      Return mvEnvironment
    End Get
    Set(value As CDBEnvironment)
      mvEnvironment = value
    End Set
  End Property


  Public Sub New(pEnv As CDBEnvironment, pTable As String, pIdentifyingValues As CDBFields)
    Me.New(pEnv, pTable, pIdentifyingValues, "is_historic", False)
  End Sub
  Public Sub New(pEnv As CDBEnvironment, pTable As String, pIdentifyingValues As CDBFields, pIsHistoricColumnName As String)
    Me.New(pEnv, pTable, pIdentifyingValues, pIsHistoricColumnName, False)
  End Sub
  Public Sub New(pEnv As CDBEnvironment, pTable As String, pIdentifyingValues As CDBFields, pIsHistoricColumnName As String, pAllowHistoricOnly As Boolean)
    Me.Environment = pEnv
    Me.Table = pTable
    Me.FieldValues = pIdentifyingValues
    Me.IsHistoricColumnName = pIsHistoricColumnName
    Me.AllowHistoric = pAllowHistoricOnly
  End Sub

  Public Function IsRecordHistoric() As Boolean
    Dim vRtn As Boolean = False
    If Not String.IsNullOrWhiteSpace(Me.Table) Then
      Dim vWhereFields As New CDBFields()
      For Each vField As CDBField In Me.FieldValues
        vWhereFields.Add(vField)
      Next
      vWhereFields.Add(Me.IsHistoricColumnName, CDBField.FieldTypes.cftCharacter, "Y")
      If Me.Environment.Connection.GetCount(Me.Table, vWhereFields) > 0 Then
        vRtn = True
      End If
    End If
    Return vRtn
  End Function

  Public Function Validate() As Boolean Implements IValidator.Validate
    Return IsRecordHistoric() = AllowHistoric
  End Function
End Class
