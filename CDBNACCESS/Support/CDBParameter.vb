Namespace Access

  Partial Public Class CDBParameter

    Public Name As String
    Public Value As String
    Public DataType As CDBField.FieldTypes
    Public Mandatory As Boolean

    Public Overrides Function ToString() As String
      ToString = Name + " = """ + Value + """"
    End Function

    Public Sub New(ByVal pName As String)
      Name = pName
      DataType = CDBField.FieldTypes.cftCharacter
      Value = ""
    End Sub

    Public Sub New(ByVal pName As String, ByVal pType As CDBField.FieldTypes)
      Name = pName
      DataType = pType
      Value = ""
    End Sub

    Public Sub New(ByVal pName As String, ByVal pValue As String)
      Name = pName
      DataType = CDBField.FieldTypes.cftCharacter
      Value = pValue.ToString
    End Sub

    Public Sub New(ByVal pName As String, ByVal pValue As Integer)
      Name = pName
      DataType = CDBField.FieldTypes.cftInteger
      Value = pValue.ToString
    End Sub

    Public Sub New(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As String)
      Name = pName
      DataType = pType
      Value = pValue
    End Sub

    Public ReadOnly Property Bool() As Boolean
      Get
        Return Value = "Y"
      End Get
    End Property

    Public ReadOnly Property IntegerValue() As Integer
      Get
        If Value Is Nothing OrElse Value = "" Then
          Return 0
        Else
          Return CInt(Value)
        End If
      End Get
    End Property

    Public ReadOnly Property LongValue() As Integer
      Get
        If Value Is Nothing OrElse Value = "" Then
          Return 0
        Else
          Return CInt(Value)
        End If
      End Get
    End Property

    Public ReadOnly Property DoubleValue() As Double
      Get
        If Value Is Nothing OrElse Value = "" Then
          Return 0
        Else
          Return CDbl(Value)
        End If
      End Get
    End Property

    Public ReadOnly Property DataTypeCode() As String
      Get
        Return CDBField.GetFieldTypeCode(DataType)
      End Get
    End Property

  End Class

End Namespace