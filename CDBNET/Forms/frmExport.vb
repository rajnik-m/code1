Public Class frmExport

  Public Enum ExportType
    etReports = 0
    etCustomForm
    etTraderApp
  End Enum

  Private mvExportType As ExportType
  Private mvMsg As String = ""
  Sub New(ByVal pReportType As ExportType)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pReportType)
  End Sub

  Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
    Try
      If eplExport.FindTextBox("ReportDestination").Text.Length = 0 Then
        eplExport.SetErrorField("ReportDestination", InformationMessages.ImExportFile) 'You must specify an Export File
      ElseIf Not eplExport.FindTextBox("ReportDestination").Text.StartsWith("\\") Then
        eplExport.SetErrorField("ReportDestination", InformationMessages.ImUNCPathOnly) 'UNC path
      else
        ExportData()
      End If

    Catch vEx As Exception
      ShowInformationMessage(vEx.Message)
    End Try
  End Sub

  Private Sub ExportData()
    Dim vFileNum As Integer
    Dim vParams As New ParameterList(True)
    Dim vDataTable As DataTable = Nothing
    Try
      vFileNum = FreeFile()
      Select Case mvExportType
        Case ExportType.etReports
          vParams.IntegerValue("ID") = CInt(IIf(eplExport.FindTextLookupBox("Report").TextBox.Text.Length > 0, eplExport.FindTextLookupBox("Report").TextBox.Text, 0))
          vParams("FileName") = eplExport.FindTextBox("ReportDestination").Text
          vParams("Append") = CStr(IIf(eplExport.FindCheckBox("AppendReport").Checked, "Y", "N"))
          vDataTable = DataHelper.ExportReport(vParams)
        Case ExportType.etCustomForm
          vParams.IntegerValue("CustomNumber") = CInt(eplExport.FindTextLookupBox("CustomNumber").TextBox.Text)
          vParams("FileName") = eplExport.FindTextBox("ReportDestination").Text
          vParams("Append") = CStr(IIf(eplExport.FindCheckBox("AppendReport").Checked, "Y", "N"))
          vDataTable = DataHelper.ExportCustomForm(vParams)
        Case ExportType.etTraderApp
          vParams.IntegerValue("ID") = CInt(eplExport.FindTextLookupBox("ID").TextBox.Text)
          vParams("FileName") = eplExport.FindTextBox("ReportDestination").Text
          vParams("Append") = CStr(IIf(eplExport.FindCheckBox("AppendReport").Checked, "Y", "N"))
          vDataTable = DataHelper.ExportTraderApplication(vParams)
      End Select
      If vDataTable IsNot Nothing Then ShowInformationMessage(GetInformationMessage(InformationMessages.ImExportComplete, mvMsg))
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enUNCPathOnly
          ShowInformationMessage(InformationMessages.ImUNCPathOnly)
        Case Else
          ShowInformationMessage(vEx.Message)
      End Select
    End Try
  End Sub

  Private Sub InitialiseControls(ByVal pReportType As ExportType)
    bpl.Refresh()
    Try
      SetExportType(pReportType)
      Select Case mvExportType
        Case ExportType.etReports
          eplExport.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optExportReport))
        Case ExportType.etCustomForm
          eplExport.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optExportCustomForm))
          cmdExport.Enabled = False
        Case ExportType.etTraderApp
          eplExport.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optExportTraderApp))
          cmdExport.Enabled = False
      End Select
      Me.Text = GetInformationMessage(ControlText.ChkAppendFile, mvMsg) 'Export %s to Text File
      My.Computer.FileSystem.CurrentDirectory = AppValues.DefaultOutputDirectory
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub SetExportType(ByVal pNewValue As ExportType)
    mvExportType = pNewValue

    Select Case mvExportType
      Case ExportType.etReports
        mvMsg = ControlText.LblReportText 'Report
      Case ExportType.etCustomForm
        mvMsg = ControlText.LblCustomFormText 'Custom Form
      Case ExportType.etTraderApp
        mvMsg = ControlText.LblTraderApplicationText  'Trader Application
    End Select
  End Sub

  Private Sub eplExport_ButtonClicked(ByVal sender As Object, ByVal pParameterName As String) Handles eplExport.ButtonClicked
    Dim vDLG As New SaveFileDialog
    Dim vFileName As String = ""
    Try
      With vDLG
        .Title = ControlText.LblExportText      'Export File
        .Filter = mvMsg & " Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .DefaultExt = "txt"
        .OverwritePrompt = False
        .FileName = eplExport.FindTextBox("ReportDestination").Text
        .CheckPathExists = True
        If eplExport.FindTextBox("ReportDestination").Text.Length = 0 Then
          Select Case mvExportType
            Case ExportType.etReports
              .FileName = "dbReports.txt"
            Case ExportType.etCustomForm
              .FileName = "dbCustom.txt"
            Case ExportType.etTraderApp
              .FileName = "dbTraderApp.txt"
          End Select
        End If
        If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
          If .FileName.StartsWith("\\") Then
            eplExport.FindTextBox("ReportDestination").Text = .FileName
          Else
            'only UNC files are supported allow the user to reselect a file
            eplExport.FindTextBox("ReportDestination").Text = .FileName
            eplExport.SetErrorField("ReportDestination", InformationMessages.ImUNCPathOnly)
          End If
          vFileName = .FileName
        End If
      End With
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enUNCPathOnly
          If vFileName.Length > 0 Then eplExport.FindTextBox("ReportDestination").Text = vFileName
          eplExport.SetErrorField("ReportDestination", InformationMessages.ImUNCPathOnly)
        Case Else
          DataHelper.HandleException(vEx)
      End Select
    Catch vEx As Exception
      ShowInformationMessage(vEx.Message)
    End Try
  End Sub

  Private Sub eplExport_ValueChanged(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles eplExport.ValueChanged
    Select Case pParameterName
      Case "CustomNumber", "ID"
        If mvExportType = ExportType.etCustomForm Or mvExportType = ExportType.etTraderApp Then
          If eplExport.FindTextLookupBox(pParameterName).Text.Length > 0 Then cmdExport.Enabled = True Else cmdExport.Enabled = False
        End If
    End Select
  End Sub

  Private Sub frmExport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    bpl.RepositionButtons()
  End Sub
End Class