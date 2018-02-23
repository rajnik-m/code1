Namespace Access
  Public Class DataValue

    'local variable(s) to hold property value(s)
    Private mvKey As String
    Private mvAttr As String
    Private mvDataType As String
    Private mvAttrValue As String
    Private mvKeyValue As Boolean

    Public Property Attr() As String
      Get
        Attr = mvAttr
      End Get
      Set(ByVal Value As String)
        mvAttr = Value
      End Set
    End Property
    Public Property KeyValue() As Boolean
      Get
        KeyValue = mvKeyValue
      End Get
      Set(ByVal Value As Boolean)
        mvKeyValue = Value
      End Set
    End Property
    Public Property AttrValue() As String
      Get
        AttrValue = mvAttrValue
      End Get
      Set(ByVal Value As String)
        mvAttrValue = Value
      End Set
    End Property
    Public Property DataType() As String
      Get
        DataType = mvDataType
      End Get
      Set(ByVal Value As String)
        mvDataType = Value
      End Set
    End Property
    Public Property Key() As String
      Get
        Key = mvKey
      End Get
      Set(ByVal Value As String)
        mvKey = Value
      End Set
    End Property

    Public ReadOnly Property FieldType() As CDBField.FieldTypes
      Get
        Select Case mvDataType
          Case "integer", "smallint", "int", ""
            FieldType = CDBField.FieldTypes.cftInteger
          Case "longinteger"
            FieldType = CDBField.FieldTypes.cftLong
          Case "date", "datetime"
            FieldType = CDBField.FieldTypes.cftDate
          Case "time"
            FieldType = CDBField.FieldTypes.cftTime
          Case "decimal", "number", "double"
            FieldType = CDBField.FieldTypes.cftCharacter
          Case "text", "nlstext", "memo", "long", "longtext", "varchar(max)", "clob"
            FieldType = CDBField.FieldTypes.cftMemo
          Case "char", "character", "varchar", "varchar2", "nlschar", "nlscharacter"
            FieldType = CDBField.FieldTypes.cftCharacter
          Case Else
            Throw New InvalidOperationException("Invalid data type encountered.")
        End Select
      End Get
    End Property
  End Class
End Namespace
