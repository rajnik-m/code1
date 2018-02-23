Namespace Access

  Partial Public Class CDBParameter

    Public Sub let_Value(ByVal pValue As String)
      Value = pValue
    End Sub

    Public Sub SetData(ByVal pName As String, ByVal pValue As String, ByVal pDataType As CDBField.FieldTypes)
      Name = pName
      Value = pValue
      DataType = pDataType
    End Sub

    Public Function CapitalisedValue(Optional ByVal pNoCapitalise As Boolean = False) As String
      If pNoCapitalise Then
        Return Value
      Else
        Return CapitaliseWords(Value)
      End If
    End Function
  End Class

End Namespace
