Namespace Data

  Partial Public Class CDBFields

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As Double) As CDBField
      Dim vField As CDBField = New CDBField(pName, pType, pValue.ToString)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Sub Clone(ByVal pFields As CDBFields)
      'Clone the fields from another collection into this collection
      'If this collection already has some fields then it will assume the names match
      'In this case just the field values will be copied
      If Me.Count > 0 Then
        For Each vField As CDBField In pFields
          Item(vField.Name).Value = vField.Value
        Next
      Else
        For Each vField As CDBField In pFields
          With vField
            Add(.Name, .FieldType, .Value, .WhereOperator)
          End With
        Next
      End If
    End Sub
  End Class

End Namespace
