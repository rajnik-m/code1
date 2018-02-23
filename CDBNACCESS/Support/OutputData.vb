Namespace Access

  Public Class OutputData

    Public Shared Sub OutputReportData(ByVal pEnv As CDBEnvironment, pWriter As IO.StreamWriter, ByVal pReportNumber As Integer)
      OutputValues(pEnv, pWriter, "reports", pReportNumber)
      OutputValues(pEnv, pWriter, "report_sections", pReportNumber)
      OutputValues(pEnv, pWriter, "report_parameters", pReportNumber)
      OutputValues(pEnv, pWriter, "report_items", pReportNumber)
      OutputValues(pEnv, pWriter, "fp_controls", pReportNumber)
      OutputValues(pEnv, pWriter, "report_version_history", pReportNumber)
    End Sub

    Private Shared Sub OutputValues(ByVal pEnv As CDBEnvironment, pWriter As IO.StreamWriter, ByVal pTable As String, Optional ByVal pIdentifier As Integer = 0, Optional ByVal pColl As CDBCollection = Nothing)
      Dim vRecordSet As CDBRecordSet
      Dim vWhere As String
      Dim vField As CDBField
      Dim vData As String
      Dim vRecord As Long
      Dim vOrder As String
      Dim vPass As Integer
      Dim vSQL As String

      Do
        If pTable = "fp_controls" Then
          If pIdentifier > 0 Then
            vWhere = "(fp_application = '" & pIdentifier & "' AND fp_page_type = 'USER') OR fp_page_type IN (SELECT report_code FROM reports WHERE report_number = " & pIdentifier & " AND report_code <> 'USER' )"
          Else
            If vPass > 0 Then
              vWhere = "fp_page_type IN (SELECT DISTINCT report_code FROM reports WHERE report_number < 10000) AND fp_page_type <> 'USER'"
            Else
              vWhere = "fp_page_type = 'USER' AND " & pEnv.Connection.DBLength("fp_application") & " < 5"
            End If
          End If
          vOrder = " ORDER BY fp_page_type, Cast(fp_application AS INTEGER), sequence_number"
        Else

          If pIdentifier > 0 Then
            vWhere = "report_number = " & pIdentifier
            vOrder = " ORDER BY report_number"
          Else
            vWhere = "report_number < 10000"
            vOrder = " ORDER BY report_number"
          End If
          Select Case pTable
            Case "report_sections"
              vOrder = vOrder & ",section_number"
            Case "report_items"
              vOrder = vOrder & ",section_number,item_number"
            Case "report_parameters"
              vOrder = vOrder & ",parameter_number"
            Case "report_version_history"
              vOrder = vOrder & ",version_number"
          End Select
        End If

        If pTable = "report_items" Then
          'Need to list attrs so as to exclude attributes
          vSQL = "SELECT report_number,section_number,item_number,report_item_type,caption,attribute_name,parameter_name,item_alignment,item_width,item_newline,amended_by,amended_on,suppress_blanks,item_format"
        ElseIf pTable = "fp_controls" Then
          'Need to list attrs so as to exclude attributes
          vSQL = "SELECT DISTINCT Cast(fp_application AS INTEGER) AS fp_application,fp_page_type,sequence_number,control_type,table_name,attribute_name,control_top,control_left,control_width,control_height,control_caption,caption_width,help_text,visible,resource_id,contact_group,parameter_name,mandatory_item,readonly_item,default_value"
        ElseIf pTable = "reports" Then
          'Need to list attrs so as to exclude attributes
          vSQL = "SELECT report_number,report_name,report_code,client,header,footer,mail_merge_output,detail_exclusive,file_output,landscape,amended_by,amended_on,mailmerge_header,application_name"
        Else
          'Select all attributes
          vSQL = "SELECT *"
        End If
        vSQL = vSQL & " FROM " & pTable & " WHERE " & vWhere & vOrder


        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        While vRecordSet.Fetch
          If vRecord = 0 Then
            pWriter.Write(pTable)
            For Each vField In vRecordSet.Fields
              pWriter.Write(",")
              pWriter.Write(vField.Name)
            Next
            pWriter.WriteLine()
          End If
          If Not pColl Is Nothing Then
            If Not pColl.Exists(vRecordSet.Fields("application_name").Value & vRecordSet.Fields("search_area").Value) Then
              RaiseError(DataAccessErrors.daeUnknowSearchArea)
            End If
          End If
          pWriter.Write(pTable)
          For Each vField In vRecordSet.Fields
            vData = vField.Value.Replace(vbCrLf, "^"c)
            vData = vField.Value.Replace(vbLf, "^"c)
            vData = vData.Replace(","c, Chr(22))
            pWriter.Write(",")
            pWriter.Write(vData)
          Next
          pWriter.WriteLine()
          vRecord = vRecord + 1
        End While
        vRecordSet.CloseRecordSet()
        vPass = vPass + 1
      Loop While pTable = "fp_controls" And pIdentifier = 0 And vPass < 2
    End Sub

  End Class

End Namespace