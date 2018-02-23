Imports System.Xml.Linq

Public Class DataDependencyValidator
  Implements IValidator

  Dim mvTable As String
  Dim mvWhereClause As CDBFields
  Dim mvValidCount As Integer
  Dim mvCompareDelegate As ValidityComparer
  Dim mvDescription As String

  Private Property Description As String

  Public Delegate Function ValidityComparer(pActualCount As Integer, pValidCount As Integer) As Boolean

  Public Sub New(pEnv As CDBEnvironment, pTable As String, pDescription As String, pWhereClause As CDBFields)
    Me.Environment = pEnv
    Me.Table = pTable
    Me.Description = pDescription
    Me.WhereClause = pWhereClause
    Me.ValidCount = 0
    Me.Comparer = New ValidityComparer(Function(pActualCount As Integer, pValidCount As Integer) pActualCount = pValidCount)
  End Sub

  Public Property Table As String
    Get
      Return mvTable
    End Get
    Private Set(value As String)
      mvTable = value
    End Set
  End Property

  Public Property ValidCount As Integer
    Get
      Return mvValidCount
    End Get
    Set(value As Integer)
      mvValidCount = value
    End Set
  End Property

  Public Property Comparer As ValidityComparer
    Get
      Return mvCompareDelegate
    End Get
    Set(value As ValidityComparer)
      mvCompareDelegate = value
    End Set
  End Property

  Public Property WhereClause As CDBFields
    Get
      Return mvWhereClause
    End Get
    Private Set(value As CDBFields)
      mvWhereClause = value
    End Set
  End Property

  Private Property Environment As CDBEnvironment

  Public Function Validate() As Boolean Implements IValidator.Validate

    Dim vRtn As Boolean = Comparer(Environment.Connection.GetCount(Me.Table, Me.WhereClause), ValidCount)

    Return vRtn

  End Function

  Public Overrides Function ToString() As String
    Dim vRtn As String = Me.Description
    If String.IsNullOrWhiteSpace(vRtn) Then
      vRtn = StrConv(Me.Table.Replace("_", " "), VbStrConv.ProperCase)
    End If
    Return vRtn
  End Function

  Shared Function FromXElement(pEnv As CDBEnvironment, pDataStore As XElement, pTableNameAttribute As String, pDescriptionAttribute As String, pAttributeValuePairs As Dictionary(Of String, String)) As List(Of DataDependencyValidator)
    Dim vRtn As New List(Of DataDependencyValidator)
    For Each vEntry In pDataStore.Elements
      Dim vWhere As New CDBFields
      If vEntry.Attribute(pTableNameAttribute) IsNot Nothing AndAlso vEntry.Attribute(pDescriptionAttribute) IsNot Nothing Then
        For Each vKVP In pAttributeValuePairs
          If vEntry.Attribute(vKVP.Key) IsNot Nothing Then
            vWhere.Add(vEntry.Attribute(vKVP.Key).Value, vKVP.Value)
          Else
            'An attribute Value has been specified that doesn't exist in this validator.  Don't include the validator as it will only partially validate.  E.g. When making an Activity Value historic, it should not validate against Activity Group Details as that table doesn't include both Activity Category and Activity Value
            vWhere = New CDBFields
            Exit For
          End If
        Next
        If vWhere.Count > 0 Then
          vRtn.Add(New DataDependencyValidator(pEnv, vEntry.Attribute(pTableNameAttribute).Value, vEntry.Attribute(pDescriptionAttribute).Value, vWhere))
        End If
      End If
    Next
    Return vRtn
  End Function

End Class
