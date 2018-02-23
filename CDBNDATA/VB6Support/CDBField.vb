Namespace Data

  Partial Public Class CDBField
    Public ReadOnly Property ValueAsFieldType() As CDBField.FieldTypes
      Get
        Select Case Value
          Case "T"
            ValueAsFieldType = CDBField.FieldTypes.cftTime
          Case "D"
            ValueAsFieldType = CDBField.FieldTypes.cftDate
          Case "N"
            ValueAsFieldType = CDBField.FieldTypes.cftNumeric
          Case "L"
            ValueAsFieldType = CDBField.FieldTypes.cftLong
          Case "I"
            ValueAsFieldType = CDBField.FieldTypes.cftInteger
          Case "M"
            ValueAsFieldType = CDBField.FieldTypes.cftMemo
          Case "U"
            ValueAsFieldType = CDBField.FieldTypes.cftUnicode
          Case Else
            ValueAsFieldType = CDBField.FieldTypes.cftCharacter
        End Select
      End Get
    End Property
  End Class
End Namespace
