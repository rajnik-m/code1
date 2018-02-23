Imports System.Linq

Public Class ExamCertificateMergeEngine

  Private mvDataTable As New DataTable
  Private mvDocuments As New List(Of String)

  Public Sub New(pFilename As String)
    Using vCsvReader As New CsvReader(pFilename)
      mvDataTable.Load(vCsvReader)
    End Using
    If mvDataTable.Columns.Contains("Standard_Document") Then
      mvDocuments.AddRange(From vDataRow As DataRow In mvDataTable.AsEnumerable
                           Select CStr(vDataRow("Standard_Document")) Distinct)
    End If
  End Sub

  Public Sub ProduceDocuments()
    For Each vDocument As String In mvDocuments
      Using vDataView As New DataView(mvDataTable, "Standard_document = '" & vDocument & "'", Nothing, DataViewRowState.CurrentRows)
        MergeDocument(vDataView.ToTable, vDocument)
      End Using
    Next vDocument
  End Sub

  Private Sub MergeDocument(pData As DataTable, pDocument As String)
    Dim vFilename As String = WriteTempFile(pData)
    Try
      Dim vLookupList As New ParameterList(True)
      vLookupList("StandardDocument") = pDocument
      Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments,
                                                     vLookupList).Rows(0)
      Dim vApplication As ExternalApplication = GetDocumentApplication(vRow.Item("DocfileExtension").ToString)
      AddHandler vApplication.ActionComplete, Sub(pAction As ExternalApplication.DocumentActions, pFilename As String)
                                                Try
                                                  File.Delete(pFilename)
                                                Catch vEx As Exception
                                                End Try
                                              End Sub
      vApplication.MergeStandardDocument(vRow.Item("StandardDocument").ToString,
                                         vRow.Item("DocfileExtension").ToString,
                                         vFilename,
                                         BooleanValue(vRow.Item("InstantPrint").ToString),
                                         True,
                                         True,
                                         True)
    Catch vEx As Exception
      File.Delete(vFilename)
    End Try
  End Sub

  Private Function WriteTempFile(pData As DataTable) As String
    Dim vFilename = Path.GetTempFileName
    Try
      Using vTempFile As New StreamWriter(vFilename)
        Dim vFirst = True
        For Each vColumn As DataColumn In pData.Columns
          If Not vFirst Then
            vTempFile.Write(",")
          End If
          vTempFile.Write(vColumn.ColumnName)
          vFirst = False
        Next vColumn
        vTempFile.Write(vbCrLf)
        For Each vDataRow As DataRow In pData.Rows
          vFirst = True
          For Each vColumn As DataColumn In pData.Columns
            If Not vFirst Then
              vTempFile.Write(",")
            End If
            vTempFile.Write("""" & CStr(vDataRow(vColumn.ColumnName)) & """")
            vFirst = False
          Next vColumn
          vTempFile.Write(vbCrLf)
        Next vDataRow
      End Using
      Return vFilename
    Catch vEx As Exception
      File.Delete(vFilename)
      Throw
    End Try
  End Function

End Class
