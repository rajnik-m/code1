Public Class ViewReceiptDetails
    Inherits CareWebControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctPrintReceipt, tblContent, "", "")

      Dim vList As New ParameterList(HttpContext.Current)
      vList("ReportCode") = "RCPTMM"
      If (Session("BatchNumber") IsNot Nothing) Then
        vList("RP1") = Session("BatchNumber")
        vList("RP2") = Session("TransactionNumber")
        Dim vFileName As String = DataHelper.GetReportFile(vList)
        Dim vReader As New System.IO.StreamReader(vFileName)
        Dim ColumnNames As String = vReader.ReadLine()
        Dim ColumnData As String = vReader.ReadToEnd()


        vReader.Close()
        vReader.Dispose()
                Dim vIdx As Integer = 0

                While vIdx < ColumnNames.Split(","c).Length
                    Dim ColumnValue As String = ColumnData.Split(","c)(vIdx).Replace(Environment.NewLine, "<br />")

                    If (ColumnValue.StartsWith("""") And ColumnValue.EndsWith("""")) Then
                        ColumnValue = ColumnValue.Substring(1, ColumnValue.Length - 2)
                    ElseIf (ColumnValue.StartsWith("""") And ColumnValue.EndsWith("""<br />")) Then
                        ColumnValue = ColumnValue.Substring(1, ColumnValue.Length - 8)
                    End If

                    Dim ColumnName As String = ColumnNames.Split(","c)(vIdx)
                    HTML = HTML.Replace("<<" + ColumnName + ">>", ColumnValue).Replace("&lt;<" + ColumnName + ">&gt;", ColumnValue).Replace("&lt;&lt;" + ColumnName + "&gt;&gt;", ColumnValue)
                    ColumnName = ColumnName.ToUpper()
                    HTML = HTML.Replace("<<" + ColumnName + ">>", ColumnValue).Replace("&lt;<" + ColumnName + ">&gt;", ColumnValue).Replace("&lt;&lt;" + ColumnName + "&gt;&gt;", ColumnValue)
                    vIdx += 1
                End While

        IO.File.Delete(vFileName)
      End If


      Dim vHTMLRow As New HtmlTableRow
      Dim vHTMLCell As New HtmlTableCell
      vHTMLRow.Cells.Add(New HtmlTableCell)
      vHTMLCell.InnerHtml = HTML
      vHTMLRow.Cells.Add(vHTMLCell)
      tblContent.Rows.Insert(0, vHTMLRow)
      tblContent.Attributes("Class") = GetClass()

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try

  End Sub

    Public Sub New()
        mvNeedsAuthentication = True
    End Sub
End Class
