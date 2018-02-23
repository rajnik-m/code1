Namespace Access

  Partial Public Class CDBDataTable

    Public Sub AddRowFromList(ByVal pList As String)
      Dim vRow As New CDBDataRow(mvColumns, mvRows.Count + 1)
      Dim vValues() As String = Split(pList, ",")
      Dim vIndex As Integer = 1
      For Each vString As String In vValues
        vRow.Item(vIndex) = vString
        vIndex += 1
      Next
      mvRows.Add(vRow)
    End Sub
  End Class
End Namespace
