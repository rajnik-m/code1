Imports System.Data.Common

Namespace Data

  Public Class CDBField
    Public Enum FieldTypes As Integer
      cftCharacter = 1
      cftMemo = 2
      cftInteger = 3
      cftLong = 4
      cftNumeric = 5
      cftDate = 6
      cftTime = 7
      cftBulk = 8
      cftFile = 9
      cftIdentity = 10        'Only used by the IDEM code at present as this is SQL server specific
      cftBit = 11             'Only used by the IDEM code at present as this is SQL server specific
      cftUnicode = 12         'Unicode
      cftBinary
      cftGUID
      cftUnknown
    End Enum

    Public Enum FieldWhereOperators
      fwoEqual = 0
      fwoNullOrGreaterThan
      fwoNullOrGreaterThanEqual
      fwoNullOrLessThan
      fwoNullOrLessThanEqual
      fwoNullOrNotEqual
      fwoNullOrEqual
      fwoGreaterThan
      fwoGreaterThanEqual
      fwoLessThan
      fwoLessThanEqual
      fwoNotEqual
      fwoNotLike
      fwoIn
      fwoNotIn
      fwoBetweenFrom
      fwoBetweenTo
      fwoLike
      fwoInOrEqual
      fwoLikeOrEqual
      fwoExist
      'All previous values are sequential 0-63
      'Values after here are bit values
      fwoOperatorOnly = 63          'Use to AND with operator to get operator values only (3FH)
      fwoOR = 64                    'Bit value which forces OR with previous field i.e  FieldA = x OR FieldB = y
      fwoOpenBracket = 128          'Bit value which adds an opening bracket prior to the field
      fwoCloseBracket = 256         'Bit value which adds an closing bracket after the field
      fwoCloseBracketTwice = 512    'Bit value which adds two closing brackets after the field
      fwoOpenBracketTwice = 1024    'Bit value which adds two opening brackets before the field
      fwoNOT = 2048                 'Bit value which adds NOT before the field
    End Enum

    Private mvFieldName As String
    Private mvFieldType As CDBField.FieldTypes
    Private mvValue As String = ""
    Private mvSpecialColumn As Boolean
    Private mvWhereOperator As FieldWhereOperators = FieldWhereOperators.fwoEqual
    Private mvDecimalPlaces As Integer
    Private mvMandatory As Boolean            'Used by create table to ensure attribute created as mandatory
    Private mvByteValue As Byte()

    Public Sub New(ByVal pFieldName As String)
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftCharacter
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValue As Integer)
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftInteger
      mvValue = pValue.ToString
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValue As String)
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftCharacter
      mvValue = pValue.ToString
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValue As String, ByVal pWhereOperator As CDBField.FieldWhereOperators)
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftCharacter
      mvValue = pValue
      mvWhereOperator = pWhereOperator
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValue As Integer, ByVal pWhereOperator As CDBField.FieldWhereOperators)
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftInteger
      mvValue = pValue.ToString
      mvWhereOperator = pWhereOperator
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValues As System.Collections.Generic.IEnumerable(Of Integer))
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftInteger
      mvValue = String.Empty
      For Each vNumber As Integer In pValues
        mvValue &= If(mvValue.Length > 0, ",", String.Empty) & vNumber.ToString
      Next vNumber
      mvWhereOperator = FieldWhereOperators.fwoIn
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pValues As System.Collections.Generic.IEnumerable(Of String))
      mvFieldName = pFieldName
      mvFieldType = FieldTypes.cftCharacter
      mvValue = String.Empty
      For Each vString As Integer In pValues
        mvValue &= If(mvValue.Length > 0, ",", String.Empty) & vString
      Next vString
      mvWhereOperator = FieldWhereOperators.fwoIn
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes)
      mvFieldName = pFieldName
      mvFieldType = pFieldType
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes, ByVal pValue As String)
      mvFieldName = pFieldName
      mvFieldType = pFieldType
      mvValue = pValue
    End Sub

    Public Sub New(ByVal pFieldName As String, ByVal pFieldType As CDBField.FieldTypes, ByVal pValue As String, ByVal pWhereOperator As FieldWhereOperators)
      mvFieldName = pFieldName
      mvFieldType = pFieldType
      mvValue = pValue
      mvWhereOperator = pWhereOperator
    End Sub

    Public Property FieldType() As CDBField.FieldTypes
      Get
        Return mvFieldType
      End Get
      Set(ByVal value As CDBField.FieldTypes)   'Should only be used by VB6 converted code
        mvFieldType = value
      End Set
    End Property

    Friend Sub SetFieldType(ByVal pType As CDBField.FieldTypes)
      mvFieldType = pType
    End Sub

    Public Property Name() As String
      Get
        Return mvFieldName
      End Get
      Set(ByVal value As String)
        mvFieldName = value
      End Set
    End Property

    Public Property Value() As String
      Get
        Return mvValue
      End Get
      Set(ByVal pValue As String)
        mvValue = pValue
      End Set
    End Property

    Public ReadOnly Property LongValue() As Integer
      Get
        Return IntegerValue
      End Get
    End Property

    Public ReadOnly Property DoubleValue() As Double
      Get
        Dim vDouble As Double
        If Double.TryParse(mvValue, vDouble) Then
          Return vDouble
        Else
          Return 0
        End If
      End Get
    End Property

    Public Property Bool() As Boolean
      Get
        Return mvValue = "Y"
      End Get
      Set(ByVal value As Boolean)
        If value Then
          mvValue = "Y"
        Else
          mvValue = "N"
        End If
      End Set
    End Property

    Public ReadOnly Property IntegerValue() As Integer
      Get
        If mvValue.Length > 0 Then
          Return CInt(mvValue)
        Else
          Return 0
        End If
      End Get
    End Property

    Public ReadOnly Property FixedValue() As String
      Get
        Dim vPos As Integer = mvValue.IndexOf(".")
        If vPos >= 0 And (mvValue.Length - vPos) > 7 Then
          Return FixTwoPlaces(mvValue).ToString("F")
        Else
          Return mvValue
        End If
      End Get
    End Property

    Public ReadOnly Property MultiLine() As String
      Get
        If mvValue.IndexOf(vbCr) < 0 Then           'If no returns
          Return mvValue.Replace(vbLf, vbCrLf)      'Then replace linefeeds with return & linefeed
        Else
          Return mvValue
        End If
      End Get
    End Property

    Public Property WhereOperator() As FieldWhereOperators
      Get
        Return mvWhereOperator
      End Get
      Set(ByVal pValue As FieldWhereOperators)
        mvWhereOperator = pValue
      End Set
    End Property

    Public Property Mandatory() As Boolean
      Get
        Return mvMandatory
      End Get
      Set(ByVal pValue As Boolean)
        mvMandatory = pValue
      End Set
    End Property

    Public Property SpecialColumn() As Boolean
      Get
        Return mvSpecialColumn
      End Get
      Set(ByVal pValue As Boolean)
        mvSpecialColumn = pValue
      End Set
    End Property

    Public Property DecimalPlaces() As Integer
      Get
        Return mvDecimalPlaces
      End Get
      Set(ByVal pValue As Integer)
        mvDecimalPlaces = pValue
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

    Public Property DBParam() As DbParameter


    Public Shared Function GetFieldType(ByVal pTypeCode As String) As FieldTypes
      Select Case pTypeCode
        Case "T"
          Return FieldTypes.cftTime
        Case "D"
          Return FieldTypes.cftDate
        Case "N"
          Return FieldTypes.cftNumeric
        Case "L"
          Return FieldTypes.cftLong
        Case "I"
          Return FieldTypes.cftInteger
        Case "M"
          Return FieldTypes.cftMemo
        Case "B"
          Return FieldTypes.cftBulk
        Case "Y"
          Return FieldTypes.cftBit
        Case "U"
          Return FieldTypes.cftUnicode
        Case "A"
          Return FieldTypes.cftBinary
        Case Else
          Return FieldTypes.cftCharacter
      End Select
    End Function

    Public Shared Function GetFieldTypeCode(ByVal pType As FieldTypes) As String
      Select Case pType
        Case FieldTypes.cftTime
          Return "T"
        Case FieldTypes.cftDate
          Return "D"
        Case FieldTypes.cftNumeric
          Return "N"
        Case FieldTypes.cftLong
          Return "L"
        Case FieldTypes.cftInteger
          Return "I"
        Case FieldTypes.cftMemo
          Return "M"
        Case FieldTypes.cftBulk
          Return "B"
        Case FieldTypes.cftBit
          Return "Y"
        Case FieldTypes.cftUnicode
          Return "U"
        Case FieldTypes.cftBinary
          Return "A"
        Case Else
          Return "C"
      End Select
    End Function

    Public Overrides Function ToString() As String
      Return String.Format("{0} Name={1} Value={2}", MyBase.ToString(), Me.Name, Me.Value)
    End Function

  End Class


End Namespace