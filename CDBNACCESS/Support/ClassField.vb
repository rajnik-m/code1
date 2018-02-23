Imports System.Data.Common
Imports Advanced.LanguageExtensions

Namespace Access

  Public Class ClassField

    Private mvName As String
    Private mvValue As String = ""
    Private mvSetValue As String = ""
    Private mvType As CDBField.FieldTypes
    Private mvParameterName As String
    Private mvInDatabase As Boolean
    Private mvSpecialColumn As Boolean
    Private mvPrimaryKey As Boolean
    Private mvPrefixRequired As Boolean
    Private mvUniqueField As Boolean
    Private mvForceAmendmentHistory As Boolean
    Private mvCaption As String
    Private mvNonUpdatable As Boolean
    Private mvByteValue As Byte()

    Public Sub New(ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes)
      mvName = pFieldName
      mvType = pFieldType
      mvInDatabase = True
    End Sub

    Public ReadOnly Property FieldType() As CDBField.FieldTypes
      Get
        Return mvType
      End Get
    End Property

    Public Property InDatabase() As Boolean
      Get
        Return mvInDatabase
      End Get
      Set(ByVal pValue As Boolean)
        mvInDatabase = pValue
      End Set
    End Property

    Public Property SetValue() As String
      Get
        Return mvSetValue
      End Get
      Set(ByVal pValue As String)
        mvSetValue = pValue
        mvValue = pValue
      End Set
    End Property

    Friend WriteOnly Property SetValueOnly() As String
      Set(ByVal pValue As String)
        mvSetValue = pValue
      End Set
    End Property

    Public Property Bool() As Boolean
      Get
        Return (mvValue = "Y")
      End Get
      Set(ByVal pValue As Boolean)
        If pValue Then
          mvValue = "Y"
        Else
          mvValue = "N"
        End If
      End Set
    End Property

    Public Property DoubleValue() As Double
      Get
        If mvValue = "" Then
          Return 0
        Else
          Return CDbl(mvValue)
        End If
      End Get
      Set(ByVal pValue As Double)
        mvValue = CStr(pValue)
      End Set
    End Property

    Public ReadOnly Property MultiLineValue() As String
      Get
        Return MultiLine(Value)
      End Get
    End Property

    Public Property IntegerValue() As Integer
      Get
        If mvValue = "" Then
          Return 0
        Else
          Return CInt(mvValue)
        End If
      End Get
      Set(ByVal pValue As Integer)
        mvValue = CStr(pValue)
      End Set
    End Property

    Public Property ByteValue() As Byte()
      Get
        Return mvByteValue
      End Get
      Set(ByVal pValue As Byte())
        mvByteValue = pValue
      End Set
    End Property

    Public Property LongValue() As Integer
      Get
        If mvValue = "" Then
          Return 0
        Else
          Return CInt(mvValue)
        End If
      End Get
      Set(ByVal pValue As Integer)
        mvValue = CStr(pValue)
      End Set
    End Property

    Public ReadOnly Property HasValue As Boolean
      Get
        Return Me.Value.HasValue
      End Get
    End Property

    Public ReadOnly Property IsNullOrWhitespace As Boolean
      Get
        Return Me.Value.IsNullOrWhitespace
      End Get
    End Property

    Public Property Value() As String
      Get
        Return mvValue
      End Get
      Set(ByVal pValue As String)
        mvValue = pValue
      End Set
    End Property

    Public Property PrimaryKey() As Boolean
      Get
        Return mvPrimaryKey
      End Get
      Set(ByVal pValue As Boolean)
        mvPrimaryKey = pValue
        If pValue = True Then mvPrefixRequired = True 'Put the prefix on the primary key
      End Set
    End Property

    Public Sub SetPrimaryKeyOnly()
      'User by the VB6 code so as not to set the prefix required
      mvPrimaryKey = True
    End Sub


    Public Property SpecialColumn() As Boolean
      Get
        Return mvSpecialColumn
      End Get
      Set(ByVal pValue As Boolean)
        mvSpecialColumn = pValue
      End Set
    End Property

    Public Property PrefixRequired() As Boolean
      Get
        Return mvPrefixRequired
      End Get
      Set(ByVal pValue As Boolean)
        mvPrefixRequired = pValue
      End Set
    End Property

    Public ReadOnly Property Name() As String
      Get
        Return mvName
      End Get
    End Property

    Friend Sub SetName(ByVal pName As String)
      mvName = pName
    End Sub

    Public Property ParameterName() As String
      Get
        If mvParameterName IsNot Nothing Then
          Return mvParameterName
        Else
          Return ProperName
        End If
      End Get
      Set(ByVal pValue As String)
        mvParameterName = pValue
      End Set
    End Property

    Public Property Caption() As String
      Get
        If mvCaption IsNot Nothing Then
          Return mvCaption
        Else
          Return StrConv(mvName.Replace("_", " "), VbStrConv.ProperCase)
        End If
      End Get
      Set(ByVal pValue As String)
        mvCaption = pValue
      End Set
    End Property

    Public ReadOnly Property ProperName() As String
      Get
        Return StrConv(mvName.Replace("_", " "), VbStrConv.ProperCase).Replace(" ", "")
      End Get
    End Property

    Public ReadOnly Property ValueChanged() As Boolean
      Get
        If mvType = CDBField.FieldTypes.cftNumeric Then
          If Val(mvValue) <> Val(mvSetValue) Then
            Return True
          ElseIf Val(mvValue) = 0 Then
            If ((mvValue = "") AndAlso (mvSetValue <> "")) OrElse ((mvSetValue = "") AndAlso (mvValue <> "")) Then Return True
          End If
        Else
          If mvValue <> mvSetValue Then Return True
        End If
      End Get
    End Property

    Friend Property UniqueField() As Boolean
      Get
        Return mvUniqueField
      End Get
      Set(ByVal pValue As Boolean)
        mvUniqueField = pValue
      End Set
    End Property

    Public Property ForceAmendmentHistory() As Boolean
      Get
        Return mvForceAmendmentHistory
      End Get
      Set(ByVal pValue As Boolean)
        mvForceAmendmentHistory = pValue
      End Set
    End Property

    Public ReadOnly Property FormattedValue() As String
      Get
        If mvType = CDBField.FieldTypes.cftNumeric Then
          If mvValue.IndexOf(".") >= 0 Then
            Return mvValue
          Else
            Return Format(Value, "Fixed")
          End If
        Else
          Return mvValue
        End If
      End Get
    End Property

    Public Property NonUpdatable() As Boolean
      Get
        Return mvNonUpdatable
      End Get
      Set(ByVal pValue As Boolean)
        mvNonUpdatable = pValue
      End Set
    End Property

    Public Property DBParam() As DbParameter

    Public Shared Widening Operator CType(ByVal pClassField As ClassField) As CDBField
      Dim vResult As New CDBField(pClassField.Name, pClassField.FieldType, pClassField.Value)
      Return vResult
    End Operator

  End Class
End Namespace