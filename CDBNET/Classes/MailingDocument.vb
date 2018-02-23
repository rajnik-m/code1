Imports System.IO
Imports System.Threading

Public Class MailingDocument
  Private Const WAITING_TIME As Integer = 60
  Private Shared mvWaitingFor As Integer = WAITING_TIME
  Private Shared mvRestartWaiting As Boolean

  Private Shared mvParams As ParameterList
  Private Shared mvPrinters As CollectionList(Of String)
  Private Shared mvLastMailingTemplate As String
  Private Shared mvLastStandardDocument As String
  Private Shared mvLastMailmergeHeader As String
  Private Shared mvLastSelectedParagraphs As String
  Private Shared mvTempDataFile As String
  Private Shared mvMMHFileName As String
  Private Shared mvNextMMHFile As String
  Private Shared mvNextTempDataFile As String
  Private Shared mvHeader As String
  Private Shared mvReader As FileReader
  Private Shared mvTempDataFileInfo As FileInfo
  Private Shared mvCount As Integer

  Public Sub New()
    mvWaitingFor = WAITING_TIME
    mvParams = Nothing
    mvPrinters = New CollectionList(Of String)
    mvLastMailingTemplate = ""
    mvLastStandardDocument = ""
    mvLastMailmergeHeader = ""
    mvLastSelectedParagraphs = ""
    mvTempDataFile = ""
    mvMMHFileName = ""
    mvNextMMHFile = ""
    mvNextTempDataFile = ""
    mvHeader = ""
    mvReader = Nothing
    mvTempDataFileInfo = Nothing
    mvCount = 0
  End Sub
  Public Sub RunMailingDocumentProduction()

    Try
      Dim vList As New ParameterList(True)
      Dim vResult As DialogResult = DialogResult.Cancel
      Dim vDefaults As New ParameterList()
      vDefaults("OutputFilename") = String.Format("{0}\Fulfillment.csv", AppValues.DefaultMailingDirectory())
      mvParams = FormHelper.ShowApplicationParameters(CareServices.TaskJobTypes.tjtMailingDocumentProduction, vDefaults)
      If mvParams.Count > 0 Then
        If (mvParams.ContainsKey("Checkbox") AndAlso mvParams("Checkbox") = "Y") AndAlso (mvParams.ContainsKey("Checkbox3") AndAlso mvParams("Checkbox3") = "N") Then
          vResult = FormHelper.ScheduleTask(mvParams)
        Else
          vResult = DialogResult.No
        End If
        Select Case vResult
          Case DialogResult.Yes
            FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtMailingDocumentProduction, mvParams, False, FormHelper.ProcessTaskScheduleType.ptsAlwaysSchedule)
          Case DialogResult.No
            If mvParams IsNot Nothing AndAlso mvParams.Count > 0 Then
              Dim vParamName As String = "DeviceName"
              For vIndex As Integer = 1 To 3
                If vIndex > 1 Then vParamName = "DeviceName" & vIndex
                If mvParams.Contains(vParamName) Then
                  If mvParams(vParamName) = "Default" Then
                    mvPrinters.Add(vIndex.ToString, (New Printing.PrinterSettings).PrinterName)
                  Else
                    mvPrinters.Add(vIndex.ToString, mvParams(vParamName).Replace("(Default)", ""))
                  End If
                End If
              Next
              ProcessDocuments()
            End If
          Case DialogResult.Cancel
            'Do nothing
        End Select
      End If

    Catch vEx As Exception
      Throw vEx
    End Try

  End Sub
  Private Shared Sub ProcessDocuments()
    'Get document and all other info
    Dim vWriter As StreamWriter = Nothing
    Dim vResult As ParameterList = DataHelper.GetMDPDocumentInfo(mvParams)
    Try
      If vResult.Count > 0 Then
        Dim vDataFileName As String = vResult("DataFileName")
        Dim vMMHeaders As String() = vResult("MailMergeHeaders").Split(","c)
        Dim vStandardDocuments As String() = vResult("StandardDocuments").Split(","c)
        Dim vPrinterNumbers As String() = vResult("Printers").Split(","c)
        Dim vExtension As String = vResult("Extension")
        Dim vBookmarks As String() = vResult("Bookmarks").Split(","c)
        Dim vFulfillmentNumber As Integer = vResult.IntegerValue("FulfillmentNumber")
        Dim vDocumentList As String = vResult("DocumentList")
        Dim vMailingDocumentNumber As Integer
        If vResult.Contains("MailingDocumentNumber") Then vMailingDocumentNumber = vResult.IntegerValue("MailingDocumentNumber")

        vResult.Remove("DataFileName")
        vResult.Remove("MailMergeHeader")
        vResult.Remove("StandardDocuments")
        vResult.Remove("Printers")
        vResult.Remove("Extension")
        vResult.Remove("Bookmarks")
        vResult.Remove("FulfillmentNumber")
        vResult.Remove("DocumentList")

        If mvLastMailingTemplate <> vResult("MailingTemplate") Then
          If mvTempDataFileInfo IsNot Nothing AndAlso mvTempDataFileInfo.Exists Then
            If IsFileInUse(mvTempDataFileInfo.FullName) Then Thread.Sleep(1000)
            mvTempDataFileInfo.Delete() 'Delete any previous temp data file
          End If
          vResult.AddConnectionData()
          mvTempDataFile = DataHelper.GetTempFile(".csv")
          'Get  Data File
          DataHelper.GetReportFile(vResult, mvTempDataFile)

          mvTempDataFileInfo = New FileInfo(mvTempDataFile)
          mvReader = New FileReader(mvTempDataFile)
          mvHeader = mvReader.ReadLine()
        End If

        vWriter = New StreamWriter(vDataFileName, True, Encoding.Default) 'Use Encoding.Default to read the Accents correctly
        vWriter.WriteLine(mvHeader)
        For vIndex As Integer = 0 To UBound(vDocumentList.Split(","c))
          vWriter.WriteLine(mvReader.ReadLine)
        Next
        vWriter.Close()

        If BooleanValue(mvParams("Checkbox")) Then       'If not generating data files only
          mvWaitingFor = -1
        Else
          If mvLastMailmergeHeader <> vMMHeaders(0) Then
            If mvMMHFileName.Length > 0 AndAlso File.Exists(mvMMHFileName) Then File.Delete(mvMMHFileName) 'Delete any previous temp data file
            mvMMHFileName = DataHelper.GetDocumentMergeData(vMMHeaders(0)) 'Get Mail Merge Header info
          End If


          'Process Documents
          Dim vIndex As Integer
          Dim vPrinter As String = ""
          Dim vNextMMHeader As String = ""
          Dim vNewMMHFile As String
          Dim vNewDataFile As String
          Try
            For Each vSD As String In vStandardDocuments
              If mvPrinters.ContainsKey(vPrinterNumbers(vIndex)) Then vPrinter = mvPrinters(vPrinterNumbers(vIndex))

              'Check Mailmerge header for AncillaryDocuments
              If vIndex > 0 AndAlso vMMHeaders(vIndex) <> vMMHeaders(0) Then
                If vMMHeaders(vIndex) <> vNextMMHeader Then
                  'Get New Mailmerge Header info
                  If mvNextMMHFile.Length > 0 AndAlso File.Exists(mvNextMMHFile) Then File.Delete(mvNextMMHFile)
                  mvNextMMHFile = DataHelper.GetDocumentMergeData(vMMHeaders(vIndex))

                  'Get New  Data File
                  If mvNextTempDataFile.Length > 0 AndAlso File.Exists(mvNextTempDataFile) Then File.Delete(mvNextTempDataFile)
                  mvNextTempDataFile = DataHelper.GetTempFile(".csv")
                  Dim vList As New ParameterList(True)
                  vList.Add("ReportCode", vMMHeaders(vIndex))
                  vList.Add("RP1", vDocumentList)
                  DataHelper.GetReportFile(vList, mvNextTempDataFile)
                  vNextMMHeader = vMMHeaders(vIndex)
                End If
                vNewMMHFile = mvNextMMHFile
                vNewDataFile = mvNextTempDataFile
              Else
                vNewMMHFile = mvMMHFileName
                vNewDataFile = vDataFileName
              End If
              vIndex += 1
              GetDocumentApplication(vExtension)
              If TypeOf DocumentApplication Is WordApplication Then '
                AddHandler DocumentApplication.ActionComplete, AddressOf ActionComplete
                AddHandler DocumentApplication.CloseAppConfirmation, AddressOf CloseAppConfirmation
                DocumentApplication.CanCloseApplication = False 'Make sure not to close Word application when its loaded and the user clicks SC main window before the system process the required items below
                If BooleanValue(mvParams("Checkbox2")) Then DocumentApplication.ConfirmClosing = True 'Display a confirmation message before closing the Word application on getting the focus on SC main window
                DirectCast(DocumentApplication, WordApplication).BuildMDPDocument(vSD, vExtension, vMailingDocumentNumber, vBookmarks, vNewMMHFile, vNewDataFile, BooleanValue(mvParams("Checkbox2")), vPrinter)
              End If
            Next
          Catch vEx As Exception
            RemoveTempFiles(vWriter, mvNextMMHFile, mvNextTempDataFile)
            Exit Sub
          End Try
        End If

        'Update mailing records with FulfillmentNumber/add FulfillmentHistory record/update incentives
        mvCount += (vDocumentList.Split(","c)).Length
        DataHelper.SetMailingDocumentsFulfilled(vFulfillmentNumber, vDocumentList, BooleanValue(mvParams("Checkbox3")), vDataFileName)

        mvLastStandardDocument = vResult("StandardDocument")
        mvLastMailingTemplate = vResult("MailingTemplate")
        mvLastMailmergeHeader = vMMHeaders(0)
        mvLastSelectedParagraphs = vResult("SelectedParagraphs")

        If BooleanValue(mvParams("Checkbox2")) = True OrElse mvWaitingFor <> -1 Then mvWaitingFor = WAITING_TIME
        RunWaitProcess()
        If TypeOf DocumentApplication Is WordApplication Then DocumentApplication.CanCloseApplication = True
      Else
        RemoveTempFiles(vWriter, mvNextMMHFile, mvNextTempDataFile)
        ShowInformationMessage(InformationMessages.ImMailingDocumentProductionSuccessfull, mvCount.ToString)
      End If
    Catch vEx As Exception
      RemoveTempFiles(vWriter, mvNextMMHFile, mvNextTempDataFile)
      Throw vEx
    End Try
  End Sub
  ''' <summary>
  ''' This method will check if the file is in use by trying to open it.
  ''' </summary>
  ''' <param name="pPath">File Path</param>
  ''' <returns>True if the file is in use else False</returns>
  ''' <remarks></remarks>
  Private Shared Function IsFileInUse(ByVal pPath As String) As Boolean

    Try
      Using vStream As New FileStream(pPath, FileMode.Open)
        vStream.Close()
      End Using
    Catch vEx As IOException
      If mvReader IsNot Nothing Then mvReader.CloseFile()
      Return True
    End Try
    Return False
  End Function

  Private Shared Sub RemoveTempFiles(ByVal vWriter As StreamWriter, ByVal pNextMMHFile As String, ByVal pNextTempDataFile As String)
    If mvReader IsNot Nothing Then mvReader.CloseFile()
    If vWriter IsNot Nothing Then vWriter.Close()
    If mvTempDataFileInfo IsNot Nothing AndAlso mvTempDataFileInfo.Exists Then
      If IsFileInUse(mvTempDataFileInfo.FullName) Then Thread.Sleep(1000)
      mvTempDataFileInfo.Delete()
    End If
    If mvMMHFileName.Length > 0 AndAlso File.Exists(mvMMHFileName) Then File.Delete(mvMMHFileName)
    If pNextMMHFile.Length > 0 AndAlso File.Exists(pNextMMHFile) Then File.Delete(pNextMMHFile)
    If pNextTempDataFile.Length > 0 AndAlso File.Exists(pNextTempDataFile) Then File.Delete(pNextTempDataFile)
  End Sub
  Private Shared Sub RunWaitProcess()
    Dim vWP As New System.ComponentModel.BackgroundWorker
    AddHandler vWP.DoWork, AddressOf DoWait
    AddHandler vWP.RunWorkerCompleted, AddressOf DoComplete
    vWP.RunWorkerAsync()
  End Sub
  Private Shared Sub DoWait(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
    Try
      Thread.Sleep(1000)
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub
  Private Shared Sub DoComplete(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
    Try
      If mvRestartWaiting Then
        'The user responded to the confirmation message displayed on closing the Word application. The timer should be restarted if the Word application is not already closed.
        mvRestartWaiting = False
        If mvWaitingFor > -1 Then mvWaitingFor = WAITING_TIME 'Do not reset the timer when the Word application is closed (ActionComplete event is handled) because we want to continue processing any remaining documents
      End If
      If mvWaitingFor > 0 Then mvWaitingFor -= 1
      If mvWaitingFor > 0 Then
        RunWaitProcess()
      ElseIf mvWaitingFor = 0 Then
        If TypeOf DocumentApplication Is WordApplication AndAlso DocumentApplication.CanCloseApplication = False Then
          'The confirmation message on closing the Word application is already displayed, so just restart the timer.
          mvWaitingFor = WAITING_TIME
          RunWaitProcess()
        Else
          If TypeOf DocumentApplication Is WordApplication Then DocumentApplication.CanCloseApplication = False 'Do not allow the Word application to be closed (automatically) or display the confirmation message
          Dim vDialogResult As DialogResult = ShowQuestion(QuestionMessages.QmMSWordKeepWaiting, MessageBoxButtons.YesNoCancel)
          If TypeOf DocumentApplication Is WordApplication Then DocumentApplication.CanCloseApplication = True
          If mvWaitingFor = -1 Then
            'The Word application has been closed after displaying the above question message and the ActionComplete event is handled, just process any remaining documents ignoring the user input.
            'Not likely to happen with current implementation but just in case
            ProcessDocuments()
          Else
            Select Case vDialogResult
              Case DialogResult.Yes
                mvWaitingFor = WAITING_TIME
                RunWaitProcess()
              Case DialogResult.No
                ProcessDocuments()
              Case DialogResult.Cancel
                'Do nothing
            End Select
          End If
        End If
      Else
        ProcessDocuments()
      End If
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub

  Private Shared Sub ActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFilename As String)
    If mvWaitingFor > -1 Then mvWaitingFor = -1
  End Sub

  Private Shared Sub CloseAppConfirmation()
    mvRestartWaiting = True
  End Sub
End Class
