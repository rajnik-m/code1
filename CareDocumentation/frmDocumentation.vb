Imports System.Xml
Imports System.IO
Imports CARE.XMLAccess
Imports CARE.Access
Imports CARE.Data
Imports System.Text
Imports System.Collections.Generic
Imports CARE.Utilities
Imports System.Reflection

Public Class frmDocumentation

  Private mvCS As New CareServices.CDBDataAccess
  Private mvNS As New CareNetServices.NDataAccess
  Private mvWA As New WebAccess.WebAccess
  Private mvEA As New ExamsAccess.ExamsDataAccess

  Private mvSourceLocation As String = "C:\dev32"
  Private mvMajorVersion As Integer = 14
  Private mvMinorVersion As Integer = 2
  Private mvVersionNumber As String
  Private mvWSSetupFile As String = mvSourceLocation & "\WEB\CareServicesSetup\CareServicesSetup.vdproj"
  Private mvWSDocumentationPath As String = mvSourceLocation & "\WEB\CareServices\Documentation"
  Private mvWSSummaryFile As String = mvSourceLocation & "\WEB\CareServices\Documentation\WEBServicesSummary.HTM"
  Private mvExamsWSSummaryFile As String = mvSourceLocation & "\WEB\CareServices\Documentation\ExamsWEBServicesSummary.HTM"
  Private mvWebServices As SortedList(Of String, WebServiceData)
  Private mvNetWebServices As SortedList(Of String, WebServiceData)
  Private mvWebWebServices As SortedList(Of String, WebServiceData)
  Private mvExamsWebServices As SortedList(Of String, WebServiceData)
  Private mvJiraReports As New CARE.Collections.CollectionList(Of SortedList(Of String, DefectReport))
  Private mvStreamWriter As StreamWriter
  Private mvProductName As String = "Advanced NFP NG"
  Private mvCopyrightName As String = "Advanced NFP"

  Private Const CARE_STYLE_SHEET As String = "<link rel=Stylesheet type=""text/css"" media=all href=""http://www.care.co.uk/site/styles/Care.css"">"

  Private Sub cmdGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGo.Click
    Try
      If chkJira.Checked AndAlso txtJira.Text.Length = 0 Then Throw New Exception("Jira file has not been selected")
      GetVersionNumber()
      If mvMajorVersion = 0 Then Exit Sub
      If My.Settings.SaveMessagesToFile Then
        mvStreamWriter = New StreamWriter(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\CareDocumentationErrors.txt", False)
        mvStreamWriter.WriteLine("Started on " & Now.ToString)
        mvStreamWriter.Close()
      End If
      If chkWebServices.Checked Or chkBoth.Checked Or chkDifferencesOnly.Checked Then ReadWebServices()
      If chkWebServices.Checked Or chkBoth.Checked Then BuildWebServicesDocumentation(WebServiceSummaryType.DataAccess)
      If chkWebServices.Checked Or chkBoth.Checked Then BuildWebServicesDocumentation(WebServiceSummaryType.ExamsAccess)
      If chkBoth.Checked Or chkDifferencesOnly.Checked Then BuildWebServiceDifferences()
      If chkWSSetupFile.Checked Then CheckWebServicesSetupFile()
      If chkBuildVersion.Checked Then BuildVersionDocumentation()
      If My.Settings.SaveMessagesToFile Then
        mvStreamWriter = New StreamWriter(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\CareDocumentationErrors.txt", True)
        mvStreamWriter.WriteLine("Completed on " & Now.ToString)
        mvStreamWriter.Close()
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message)
    End Try
  End Sub

  Private Sub CheckWebServicesSetupFile()
    'Read all files from the documenation directory and check if they are in the setup file
    Dim vFiles As New List(Of String)
    Dim vReader As New StreamReader(mvWSSetupFile)
    While Not vReader.EndOfStream
      Dim vLine As String = vReader.ReadLine
      If vLine.Contains("SourcePath") Then
        Dim vIndex As Integer = vLine.IndexOf("Documentation\\")
        If vIndex >= 0 Then
          Dim vFilename As String = vLine.Substring(vIndex + 15)
          vFilename = vFilename.Substring(0, vFilename.Length - 1).ToLower
          If vFiles.Contains(vFilename) Then
            ShowError("Web Services Setup duplicate entry for " & vFilename)
          Else
            vFiles.Add(vFilename)
          End If
        End If
      End If
    End While
    vReader.Close()

    Dim vExistingFiles As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    vExistingFiles = My.Computer.FileSystem.GetFiles(mvWSDocumentationPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
    For Each vFileName As String In vExistingFiles
      vFileName = Path.GetFileName(vFileName).ToLower
      If Not vFiles.Contains(vFileName) Then
        ShowError("Web Services Setup is missing " & vFileName)
      End If
    Next
    ssl.Text = String.Format("Processing Complete")
  End Sub

  Private Sub ShowError(ByVal pString As String)
    If My.Settings.SaveMessagesToFile Then
      mvStreamWriter = New StreamWriter(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\CareDocumentationErrors.txt", True)
      mvStreamWriter.WriteLine(pString)
      mvStreamWriter.Close()
    Else
      MessageBox.Show(pString)
    End If
  End Sub

  Private Sub chkDifferencesOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDifferencesOnly.CheckedChanged
    If chkDifferencesOnly.Checked Then chkBoth.Checked = False
  End Sub

  Private Sub chkWebServices_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBoth.CheckedChanged
    If chkBoth.Checked Then chkDifferencesOnly.Checked = False
  End Sub

  Private Sub GetVersionNumber()
    mvVersionNumber = String.Format("{0}.{1}", mvMajorVersion, mvMinorVersion)
    'Try
    '  Dim vReader As StreamReader = New StreamReader(mvSourceLocation & "\CDBNACCESS\CDBEnvironment.vb")
    '  Do While vReader.EndOfStream = False
    '    Dim vLine As String = vReader.ReadLine
    '    If vLine.Contains("VERSION_NUMBER") Then
    '      Dim vVersion As String = vLine.Substring(vLine.IndexOf("=") + 1).Trim(" """.ToCharArray)
    '      Dim vItems() As String = vVersion.Split(".".ToCharArray)
    '      mvMajorVersion = CInt(vItems(0))
    '      mvMinorVersion = CInt(vItems(1))
    '      mvVersionNumber = String.Format("{0}.{1}", mvMajorVersion, mvMinorVersion)
    '      Exit Do
    '    End If
    '  Loop
    '  vReader.Close()
    'Catch vEx As Exception
    '  MessageBox.Show(vEx.Message)
    'End Try
  End Sub

  Private Sub ReadWebServices()
    mvWebServices = New SortedList(Of String, WebServiceData)
    mvNetWebServices = New SortedList(Of String, WebServiceData)
    mvWebWebServices = New SortedList(Of String, WebServiceData)
    mvExamsWebServices = New SortedList(Of String, WebServiceData)
    ReadWebServiceFile(mvWebServices, mvSourceLocation & "\web\Careservices\dataaccess.asmx.vb")
    ReadWebServiceFile(mvNetWebServices, mvSourceLocation & "\web\Careservices\ndataaccess.asmx.vb")
    ReadWebServiceFile(mvWebWebServices, mvSourceLocation & "\web\Careservices\webaccess.asmx.vb")
    ReadWebServiceFile(mvExamsWebServices, mvSourceLocation & "\web\Careservices\examsdataaccess.asmx.vb")
  End Sub

  Private Sub ReadWebServiceFile(ByVal pWebServices As SortedList(Of String, WebServiceData), ByVal pFileName As String)
    Dim vReader As New StreamReader(pFileName)
    While Not vReader.EndOfStream
      Dim vLine As String = vReader.ReadLine
      If vLine.Contains("WebMethod") Then
        Dim vDocument As String = ""
        Dim vDescription As String = ""
        Dim vStartDesc As Integer = vLine.IndexOf("""") + 1
        Dim vEndDesc As Integer = vLine.IndexOf("<A HREF=", vStartDesc)
        If vEndDesc < 0 Then
          vDescription = vLine.Substring(vStartDesc)
        Else
          vDescription = vLine.Substring(vStartDesc, vEndDesc - vStartDesc)
          Dim vStartDoc As Integer = vLine.IndexOf("documentation\") + 14
          Dim vEndDoc As Integer = vLine.IndexOf(">", vStartDoc)
          vDocument = vLine.Substring(vStartDoc, vEndDoc - vStartDoc)
        End If
        vLine = vReader.ReadLine.Trim
        Dim vStartFunction As Integer = vLine.IndexOf("Function") + 9
        Dim vEndFunction As Integer = vLine.IndexOf("(", vStartFunction)
        Dim vFunction As String = vLine.Substring(vStartFunction, vEndFunction - vStartFunction)
        If vDocument.Length > 0 AndAlso Not vDocument.StartsWith(vFunction) Then
          ShowError(String.Format("Documentation Reference {0} Incorrect for Web Service {1}", vDocument, vFunction))
        ElseIf vDocument.Length = 0 AndAlso Not vDescription.StartsWith("DEPRECATED") AndAlso Not vDescription.StartsWith("NOT CURRENTLY SUPPORTED") Then
          ShowError(String.Format("No Documentation Reference for Web Service {0}", vFunction))
        End If
        ssl.Text = String.Format("Processing {0}", vFunction)
        sts.Refresh()
        Dim vWebService As New WebServiceData(vFunction, vDescription.Trim)
        If vLine.StartsWith("Public Function") AndAlso vLine.EndsWith("As String") Then
          vWebService.Syntax = vLine.Substring(16, vLine.Length - (16 + 9)).Replace("ByVal ", "").Trim
          While Not vLine.StartsWith("End Function")
            vLine = vReader.ReadLine.Trim
            If vLine.Contains("AddressOf ") Then
              If vLine.Contains(".AddRecord") Then
                vWebService.CallType = WebServiceData.CallTypes.AddRecord
              ElseIf vLine.Contains(".UpdateRecord") Then
                vWebService.CallType = WebServiceData.CallTypes.UpdateRecord
              ElseIf vLine.Contains(".DeleteRecord") Then
                vWebService.CallType = WebServiceData.CallTypes.DeleteRecord
              End If
              If vWebService.CallType <> WebServiceData.CallTypes.None Then
                Dim vIndex As Integer = vLine.IndexOf("CARE.Access")
                If vIndex > 0 Then
                  vWebService.ClassName = vLine.Substring(vIndex + 12).Replace("(Nothing))", "")
                End If
              End If
            End If
          End While
        End If



        pWebServices.Add(vFunction, vWebService)
      End If
    End While
    vReader.Close()
  End Sub

  Private Sub BuildVersionDocumentation()
    Dim vEnv As New CDBEnvironment("DSN=defects", "care_admin", "care_admin")
    Dim vRelease As New Release(vEnv)
    Dim vReleases As New List(Of Release)
    mvJiraReports = New CARE.Collections.CollectionList(Of SortedList(Of String, DefectReport))

    Dim vSQLStatement As New SQLStatement(vEnv.Connection, vRelease.GetRecordSetFields(), "releases", New CDBField("supported", "Y"), "release")
    Dim vRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet
    With vRecordSet
      While .Fetch
        vRelease = New Release(vEnv)
        vRelease.InitFromRecordSet(vRecordSet)
        vReleases.Add(vRelease)
      End While
    End With
    vRecordSet.CloseRecordSet()

    If chkJira.Checked Then
      'If we are processing Jira data, read all the data out of the csv file into a SortedList
      Dim vJiraDataFile As New FileReader(txtJira.Text, FileReader.FileReaderTypes.rftCharSeparated, CType(",", Char()), True)
      If vJiraDataFile.FileOpen = False Then Throw New Exception(String.Format("Unable to open Jira file {0}", txtJira.Text))
      Dim vLastRelease As String = ""
      Dim vJiraList As New SortedList(Of String, DefectReport)
      If vJiraDataFile.EndOfFile = False Then vJiraDataFile.ReadLine() 'First line is header so skip it
      While Not vJiraDataFile.EndOfFile
        vJiraDataFile.ReadLine()
        If Not vJiraDataFile.EndOfFile Then
          Dim vDefectReport As New DefectReport(vEnv, vJiraDataFile.Fields)
          If vLastRelease.Length > 0 AndAlso vLastRelease <> vDefectReport.FixedInRelease Then
            mvJiraReports.Add(vLastRelease, vJiraList)
            vJiraList = New SortedList(Of String, DefectReport)
          End If
          vLastRelease = vDefectReport.FixedInRelease
          vJiraList.Add(vDefectReport.ToString, vDefectReport)
        End If
      End While
      If vJiraList.Count > 0 Then mvJiraReports.Add(vLastRelease, vJiraList)
    End If

    For Each vRelease In vReleases
      With vRelease
        CreateVersionDocument(vEnv, .FieldValueString("destination_file_name") & ".html", .FieldValueString("destination_dir"), .FieldValueInteger("from_build"), .FieldValueInteger("first_release_build"), .FieldValueString("release_desc"), .FieldValueString("release"))
        CreateVersionDocument(vEnv, .FieldValueString("destination_file_name") & "patch.html", .FieldValueString("destination_dir"), .FieldValueInteger("first_release_build") + 1, .FieldValueInteger("to_build"), .FieldValueString("release_desc") & " Patch", .FieldValueString("release"))
        CreateDBChangeDocument(vEnv, "dbmods" & .FieldValueString("release") & ".html", .FieldValueString("destination_dir"), .FieldValueString("release_desc"))
      End With
    Next
  End Sub

  Private Sub CreateVersionDocument(ByVal pEnv As CDBEnvironment, ByVal pDestFilename As String, ByVal pDestDir As String, ByVal pStartVersion As Integer, ByVal pEndVersion As Integer, ByVal pVDesc As String, ByVal pReleaseNUmber As String)

    ssl.Text = String.Format("Processing {0}", pVDesc)
    sts.Refresh()

    Dim vDefectReports As New List(Of DefectReport)
    Dim vJiraList As New SortedList(Of String, DefectReport)
    If mvJiraReports.Count > 0 Then
      'See if we have any Jira's for this release, if so retrieve the SortedList contining them
      Dim vKey As String = pVDesc
      If vKey.EndsWith("Patch") Then
        vKey = vKey.Substring(0, (vKey.Length - 5)).Trim
      End If
      If mvJiraReports.ContainsKey(vKey) Then vJiraList = mvJiraReports.Item(vKey)
    End If

    Dim vAttrs As String = "version,workaround,report_type,reported_on,area,description,fixed_in_version,fixed_on,resolution,log_number,r.report_number,finder,change_configuration,change_database,change_reports"
    Dim vWhereFields As New CDBFields()
    vWhereFields.Add("fixed_in_version", pStartVersion, CDBField.FieldWhereOperators.fwoBetweenFrom)
    vWhereFields.Add("fixed_in_version#1", pEndVersion, CDBField.FieldWhereOperators.fwoBetweenTo)
    vWhereFields.AddJoin("rf.report_number", "r.report_number")
    vWhereFields.Add("version_history", "N", CDBField.FieldWhereOperators.fwoNotEqual)
    'Exclude Rave fixes as these are in Jira as well
    vWhereFields.Add("responsibility", "", CType(CDBField.FieldWhereOperators.fwoEqual + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
    vWhereFields.Add("responsibility#2", "Rave", CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoNotEqual + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
    Dim vSQLStatement As New SQLStatement(pEnv.Connection, vAttrs, "report_fixes rf, reports r", vWhereFields, "fixed_in_version DESC, rf.report_number DESC")
    Dim vDefectRecordSet As CDBRecordSet = vSQLStatement.GetRecordSet
    If vJiraList.Count > 0 Then
      'Have Jira data so merge the Jira & BR data into one collection, ordered by FixedinVersion desc
      'For a particular version, all BR's will be added followed by all Jira's
      Dim vDefectReport As New DefectReport(pEnv)
      If vDefectRecordSet.Fetch() Then vDefectReport.InitFromRecordSet(vDefectRecordSet)
      'Debug.WriteLine("BR Fixed in version:" & vDefectReport.FixedInVersion)
      For vIndex As Integer = (vJiraList.Count - 1) To 0 Step -1
        Dim vJiraDefect As DefectReport = vJiraList.Values(vIndex)
        'Debug.WriteLine("Jira Fixed in version: " & vJiraDefect.FixedInVersion)
        If vJiraDefect.FixedInVersion > 0 AndAlso vJiraDefect.FixedInVersion >= pStartVersion AndAlso vJiraDefect.FixedInVersion <= pEndVersion Then
          If vDefectReport.Existing Then
            If vJiraDefect.FixedInVersion > vDefectReport.FixedInVersion Then
              'Debug.WriteLine("BR Fixed in version:" & vDefectReport.FixedInVersion)
              If vJiraDefect.Existing Then vDefectReports.Add(vJiraDefect)
            Else
              While vDefectReport.Existing AndAlso (vJiraDefect.FixedInVersion <= vDefectReport.FixedInVersion)
                'Debug.WriteLine("BR Fixed in version:" & vDefectReport.FixedInVersion)
                vDefectReports.Add(vDefectReport)
                vDefectReport = New DefectReport(pEnv)
                If vDefectRecordSet.Fetch Then vDefectReport.InitFromRecordSet(vDefectRecordSet)
              End While
              If vJiraDefect.Existing Then vDefectReports.Add(vJiraDefect)
            End If
          Else
            If vJiraDefect.Existing Then vDefectReports.Add(vJiraDefect)
          End If
        End If
      Next
      If vDefectReport.Existing Then
        'Gone through all the Jira'a and still have some BR's left unprocessed
        'Debug.WriteLine("BR Fixed in version:" & vDefectReport.FixedInVersion)
        vDefectReports.Add(vDefectReport)
        While vDefectRecordSet.Fetch
          vDefectReport = New DefectReport(pEnv)
          vDefectReport.InitFromRecordSet(vDefectRecordSet)
          'Debug.WriteLine("BR Fixed in version:" & vDefectReport.FixedInVersion)
          vDefectReports.Add(vDefectReport)
        End While
      End If
    Else
      'No Jira data so just process BR's
      While vDefectRecordSet.Fetch
        Dim vDefectReport As New DefectReport(pEnv)
        vDefectReport.InitFromRecordSet(vDefectRecordSet)
        vDefectReports.Add(vDefectReport)
      End While
    End If
    vDefectRecordSet.CloseRecordSet()

    Dim vFileName As String = "c:\temp\releases.html"   '"C:\Users\stephen.smith\AppData\Local\TempDocumentation\releases.html"
    Dim vWriter As New StreamWriter(vFileName)
    vWriter.WriteLine("<HTML><HEAD>")
    vWriter.WriteLine(CARE_STYLE_SHEET)
    vWriter.WriteLine("</HEAD>")
    vWriter.WriteLine("<BODY>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<H1>" & mvProductName & " " & pVDesc & " Change Details</H1>")
    vWriter.WriteLine("Please note that this includes builds that have not yet been approved for general release to clients.<BR>For the latest information on builds that have been formally released, see the ""latest Build details"" page.")
    Dim vFixedInVersion As String
    Dim vLastBuild As String = ""
    Dim vReportNumber As String
    Dim vLogNumber As String
    Dim vInList As Boolean

    For Each vDefectReport As DefectReport In vDefectReports
      vFixedInVersion = vDefectReport.FixedInVersion.ToString
      If vFixedInVersion <> vLastBuild Then
        If vInList Then vWriter.WriteLine("</TABLE>")
        vWriter.WriteLine("<HR width=""100%"" align=""left"">")
        vWriter.WriteLine("<a name=""" & vFixedInVersion & """</a>")
        vWriter.WriteLine("<STRONG>Changes in build " & vFixedInVersion & " (last change date " & vDefectReport.FieldValueString("fixed_on") & ")</STRONG><P>")
        vLastBuild = vFixedInVersion
        vWriter.WriteLine("<TABLE width=""100%"" BORDER>")
        vWriter.WriteLine("<TR>")
        vWriter.WriteLine("<TH width=""60"">Support No</TH>")
        vWriter.WriteLine("<TH width=""90"">Report No</TH>")
        vWriter.WriteLine("<TH width=""90"">Report Type</TH>")
        vWriter.WriteLine("<TH width=""160"">Area</TH>")
        vWriter.WriteLine("<TH width=""490"">Description</TH>")
        vWriter.WriteLine("</TR>")
        vInList = True
      End If
      vWriter.WriteLine("<TR>")
      vReportNumber = vDefectReport.FieldValueString("report_number")
      ssl.Text = String.Format("Processing {0} Report {1}", pVDesc, vReportNumber)
      sts.Refresh()
      vLogNumber = vDefectReport.FieldValueString("log_number")
      If vLogNumber = "" Then vLogNumber = "n/a"
      vWriter.WriteLine("<TD align=""center"" valign=""top"">" & vLogNumber & "</TH>")
      If vDefectReport.ReportType = DefectReport.DefectReportType.BugReport Then
        vWriter.WriteLine("<TD align=""center"" valign=""Top""><A HREF=""br" & vReportNumber & ".html"">" & vReportNumber & "</TH>")
      Else
        'Jira
        vWriter.WriteLine("<TD align=""center"" valign=""Top""><A HREF=""" & vReportNumber & ".html"">" & vReportNumber & "</TH>")
      End If
      vWriter.WriteLine("<TD align=""center"" valign=""top"">" & vDefectReport.FieldValueString("report_type") & "</TH>")
      vWriter.WriteLine("<TD valign=""top"">" & vDefectReport.FieldValueString("area") & "</TH>")
      vWriter.WriteLine("<TD valign=""top"">" & vDefectReport.FieldValueString("description") & "</TH>")
      Dim vImpactNotes As String = ""     'TODO impactnotes
      vWriter.WriteLine("<TD valign=""top"">" & vImpactNotes & "</TD>")         'Impact Notes
      vWriter.WriteLine("</TR>")

      Dim vBRFileName As String = "c:\temp\"    '"C:\Users\stephen.smith\AppData\Local\TempDocumentation\"
      If vDefectReport.ReportType = DefectReport.DefectReportType.BugReport Then
        vBRFileName &= "br" & vReportNumber & ".html"
      Else
        'Jira
        vBRFileName &= vReportNumber & ".html"
      End If
      Dim vBRWriter As New StreamWriter(vBRFileName)

      vBRWriter.WriteLine("<HTML><HEAD>")
      vBRWriter.WriteLine(CARE_STYLE_SHEET)
      vBRWriter.WriteLine("</HEAD>")
      vBRWriter.WriteLine("<BODY>")
      vBRWriter.WriteLine("<p>")
      vBRWriter.WriteLine("<H1>" & mvProductName & " Development Report " & vReportNumber & "</H1>")
      vBRWriter.WriteLine("<HR width=""100%"" align=""left"">")
      vBRWriter.WriteLine("<p>")

      vBRWriter.WriteLine("<TABLE width=""100%"">")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD width=""160""><b>Report Number:</b> </TD>")
      vBRWriter.WriteLine("<TD width=""60"">" & vReportNumber & "</TD>")
      vBRWriter.WriteLine("<TD width=""160""><b>Report Type:</b> </TD>")
      vBRWriter.WriteLine("<TD width=""40"">" & vDefectReport.FieldValueString("report_type") & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD><b>Reported In:</b> </TD>")
      vBRWriter.WriteLine("<TD>" & vDefectReport.FieldValueString("version") & "</TD>")
      vBRWriter.WriteLine("<TD><b>On:</b> </TD>")
      vBRWriter.WriteLine("<TD>" & vDefectReport.FieldValueString("reported_on") & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("</TABLE>")
      vBRWriter.WriteLine("<TABLE width=""100%"">")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD width=""160""><b>Found By:</b> </TD>")
      If vDefectReport.FieldValueString("finder") = "C" Then
        vBRWriter.WriteLine("<TD>Client</TD>")
      Else
        vBRWriter.WriteLine("<TD>Internal</TD>")
      End If
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD><b>Related Support Nos:</b> </TD>")

      Dim vSupportLogs As String = ""
      If vDefectReport.ReportType = DefectReport.DefectReportType.BugReport Then
        Dim vSQLLogs As CDBRecordSet = New SQLStatement(pEnv.Connection, "log_number", "report_links", New CDBField("report_number", vReportNumber)).GetRecordSet
        While vSQLLogs.Fetch
          If vSupportLogs.Length > 0 Then vSupportLogs &= ", "
          vSupportLogs &= vSQLLogs.Fields(1).Value
        End While
        vSQLLogs.CloseRecordSet()
      Else
        'Jira
        vSupportLogs = vDefectReport.FieldValueString("log_number")
      End If
      If vSupportLogs.Length = 0 Then vSupportLogs = "None."
      vBRWriter.WriteLine("<TD colspan=""4"">" & vSupportLogs & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD align=""top""><b>Description:</b> </TD>")
      vBRWriter.WriteLine("<TD colspan=""4"">" & vDefectReport.FieldValueString("description") & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("</TABLE><p><p>")

      vBRWriter.WriteLine("<TABLE width=""100%"">")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD width=""160"" valign=""top""><b>Fixed in versions: </b></TD>")
      vBRWriter.WriteLine("<TD align=""left"">")
      vBRWriter.WriteLine("<TABLE BORDER width=""300"">")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TH>Version  </TH>")
      vBRWriter.WriteLine("<TH>Build    </TH>")
      vBRWriter.WriteLine("</TR>")
      Dim vHTML As String = ""
      If vDefectReport.ReportType = DefectReport.DefectReportType.BugReport Then
        vWhereFields.Clear()
        vWhereFields.Add("report_number", vReportNumber)
        vWhereFields.AddJoin("rf.fixed_in_version", "v.version")
        vWhereFields.AddJoin("v.release", "r.release")
        Dim vRSFixedIn As CDBRecordSet = New SQLStatement(pEnv.Connection, "v.release,fixed_in_version,destination_dir", "report_fixes rf, versions v, releases r", vWhereFields, "fixed_in_version DESC").GetRecordSet
        While vRSFixedIn.Fetch
          vBRWriter.WriteLine("<TR>")
          Dim vHyperlinkVersion As String = vRSFixedIn.Fields("release").Value.Trim
          Dim vHyperlinkBuild As String = vRSFixedIn.Fields("fixed_in_version").Value.Trim
          Dim vHyperlinkFilePath As String = String.Empty
          If vHyperlinkVersion <> pReleaseNUmber Then
            Dim vDirectory As New DirectoryInfo(vRSFixedIn.Fields("destination_dir").Value)
            If vDirectory.Exists Then
              vHyperlinkFilePath = vDirectory.Parent.Name
              vHyperlinkFilePath = "..\..\" & vHyperlinkFilePath & "\" & vDirectory.Name
            Else
              vHyperlinkFilePath = vRSFixedIn.Fields("destination_dir").Value
            End If
            vHyperlinkFilePath &= "\"
          End If
          vBRWriter.WriteLine("<TD align=""center"">" & "<A HREF=""" & vHyperlinkFilePath & "version" & vHyperlinkVersion.Replace(".", "") & vHTML & ".html" & """>" & vHyperlinkVersion & "</TD>")
          vBRWriter.WriteLine("<TD align=""center"">" & "<A HREF=""" & vHyperlinkFilePath & "version" & vHyperlinkVersion.Replace(".", "") & vHTML & ".html#" & vHyperlinkBuild.Replace(".", "") & """>" & vHyperlinkBuild & "</TD>")
          vBRWriter.WriteLine("</TR>")
          vHTML = "patch"
        End While
        vRSFixedIn.CloseRecordSet()
      Else
        'Jira
        vBRWriter.WriteLine("<TR>")
        Dim vHyperlinkVersion As String = vDefectReport.FixedInRelease
        Dim vHyperlinkBuild As String = vDefectReport.FieldValueString("fixed_in_version")
        Dim vHyperlinkFilePath As String = String.Empty
        If vHyperlinkVersion <> pReleaseNUmber Then
          Dim vDirectory As New DirectoryInfo(pDestDir)
          If vDirectory.Exists Then
            vHyperlinkFilePath = vDirectory.Parent.Parent.Name
            vHyperlinkFilePath = "..\..\" & vHyperlinkFilePath
          Else
            vHyperlinkFilePath = pDestDir
          End If
          vHyperlinkFilePath &= "\"
        End If
        vBRWriter.WriteLine("<TD align=""center"">" & "<A HREF=""" & vHyperlinkFilePath & "version" & vHyperlinkVersion.Replace(".", "") & vHTML & ".html" & """>" & vHyperlinkVersion & "</TD>")
        vBRWriter.WriteLine("<TD align=""center"">" & "<A HREF=""" & vHyperlinkFilePath & "version" & vHyperlinkVersion.Replace(".", "") & vHTML & ".html#" & vHyperlinkBuild.Replace(".", "") & """>" & vHyperlinkBuild & "</TD>")
        vBRWriter.WriteLine("</TR>")
      End If

      vBRWriter.WriteLine("</TABLE><p><p>")
      vBRWriter.WriteLine("</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("</TABLE><p><p>")

      vBRWriter.WriteLine("<STRONG>Details of solution:" & "</STRONG>")
      vBRWriter.WriteLine("<TABLE BORDER width=""100%"">")
      vBRWriter.WriteLine("<TR>")
      Dim vResolution As String = GetPlainText(vDefectReport.FieldValueString("resolution"))
      If vResolution.Length = 0 Then vResolution = "None."
      vBRWriter.WriteLine("<TD>" & vResolution & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("</TABLE><p><p>")

      Dim vWorkAround As String = GetPlainText(vDefectReport.FieldValueString("workaround"))
      If vWorkAround.Length = 0 Then vWorkAround = "None."
      vBRWriter.WriteLine("<STRONG>Available workarounds:" & "</STRONG>")
      vBRWriter.WriteLine("<TABLE BORDER width=""100%"">")
      vBRWriter.WriteLine("<TR>")
      vBRWriter.WriteLine("<TD>" & vWorkAround & "</TD>")
      vBRWriter.WriteLine("</TR>")
      vBRWriter.WriteLine("</TABLE><p><p>")
      vBRWriter.WriteLine("<p><p><p><p><p>")
      vBRWriter.WriteLine("<FN><CENTER>" & "This page was last updated automatically on:" & Date.Today & "</CENTER></FN>")
      vBRWriter.WriteLine("</BODY></HTML>")
      vBRWriter.Close()
      FileCopy(vBRFileName, pDestDir & "\" & vBRFileName.Substring((vBRFileName.LastIndexOf("\") + 1)))
    Next
    If vInList Then vWriter.WriteLine("</TABLE>")

    vWriter.WriteLine("<HR width=""100%"" align=""left"">")
    vWriter.WriteLine("<FN><CENTER>" & "This page was last updated automatically on:" & Date.Today & "</CENTER></FN>")
    vWriter.WriteLine("</BODY></HTML>")
    vWriter.Close()
    FileCopy(vFileName, pDestDir & "\" & pDestFilename)
    ssl.Text = "Complete"
    sts.Refresh()
  End Sub

  Private Function GetPlainText(ByVal pText As String) As String
    If pText.Contains("\rtf") Then
      rtb.Rtf = pText
    Else
      rtb.Text = pText
    End If
    Dim vText As String = rtb.Text.Trim
    vText = vText.Replace(vbLf, "<p>")
    While vText.EndsWith("<p>")
      vText.Remove(vText.Length - 3, 3)
    End While
    Return vText.Trim
  End Function

  Public Sub CreateDBChangeDocument(ByVal pEnv As CDBEnvironment, ByVal pDestFilename As String, ByVal pDestDir As String, ByVal pVDesc As String)
    ssl.Text = String.Format("Processing the Data Structure Changes for {0}", pVDesc)
    sts.Refresh()
    Dim vNewAttrText As String = "New attribute - "
    Dim vAlterAttrText As String = "Existing attribute  -"

    Dim vSQL As New SQLStatement(pEnv.Connection, "config_value", "config", New CDBField("config_name", "dbupgrade_location"))
    Dim vDBUpgradeFile As String = vSQL.GetValue()
    If vDBUpgradeFile.Length > 0 Then
      vDBUpgradeFile = vDBUpgradeFile & "\dbupgrade.txt"
    Else
      vDBUpgradeFile = "//NTDEV4/GUIBUILD/ADMIN\dbupgrade.txt"
    End If

    Dim vEnv As New CDBEnvironment("CDBWEBSERVER", "care_admin", "care_admin")

    Dim vFileName As String = "c:\temp\dbmods.htm"
    Dim vWriter As New StreamWriter(vFileName)
    vWriter.WriteLine("<HTML><HEAD>")
    vWriter.WriteLine(CARE_STYLE_SHEET)
    vWriter.WriteLine("</HEAD><BODY>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<H1>" & mvProductName & " Version " & pVDesc & " Database Structure Changes</H1>")

    'New Tables
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<h4>New Tables</h4>")
    vWriter.WriteLine("<p style='text-align: justify;'>The following database tables have been added in this version:</p>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<TABLE BORDER>")
    vWriter.WriteLine("<TR>")
    vWriter.WriteLine("<TH>Table Name</TH>")
    vWriter.WriteLine("<TH>Summary Description</TH>")
    vWriter.WriteLine("</TR>")

    Dim vReader As New StreamReader(vDBUpgradeFile)
    Dim vDocumentChange As Boolean
    Dim vWhereFields As New CDBFields()
    vWhereFields.Add("table_name")
    Dim vSQLTable As New SQLStatement(vEnv.Connection, "table_notes", "maintenance_tables", vWhereFields)
    Dim vTableName As String = ""
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine().TrimStart
      Dim vItems() As String = vLine.Split(","c)
      Select Case vItems(0)
        Case "CHANGEINFO"
          If vItems(1).Trim = pVDesc Then vDocumentChange = True
        Case "ENDCHANGE"
          vDocumentChange = False
        Case "CREATETABLE"
          If vDocumentChange Then
            vTableName = vItems(1).Trim
            vWriter.WriteLine("<TR>")
            vWriter.WriteLine("<TD>" & vTableName & "</TD>")
            vWhereFields(1).Value = vTableName
            Dim vComment As String = vSQLTable.GetValue
            If vComment.Length > 0 Then
              vWriter.WriteLine("<TD>" & vComment & "</TD>")
            End If
          End If
      End Select
    End While
    vWriter.WriteLine("</TABLE>")

    'Altered Tables
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<h4>Amended Tables</h4>")
    vWriter.WriteLine("<p style='text-align: justify;'>The following database tables have been amended in this version:</p>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<TABLE BORDER>")
    vWriter.WriteLine("<TR>")
    vWriter.WriteLine("<TH>Table Name</TH>")
    vWriter.WriteLine("<TH>Attribute Name</TH>")
    vWriter.WriteLine("<TH>Summary Description</TH>")
    vWriter.WriteLine("</TR>")
    vReader.BaseStream.Seek(0, SeekOrigin.Begin)
    vWhereFields.Add("attribute_name")
    Dim vSQLAttribute As New SQLStatement(vEnv.Connection, "attribute_notes", "maintenance_attributes", vWhereFields)

    Dim vProcessTable As Boolean
    Dim vHeading As Boolean
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine().TrimStart
      Dim vItems() As String = vLine.Split(","c)
      Select Case vItems(0)
        Case "CHANGEINFO"
          If vItems(1).Trim = pVDesc Then vDocumentChange = True
        Case "ENDCHANGE"
          vDocumentChange = False
        Case "BEGINTABLE"
          If vDocumentChange Then
            vProcessTable = True
            vHeading = False
            vTableName = vItems(1).Trim
            vWriter.WriteLine("<TR>")
          End If
        Case "CREATETABLE", "ENDTABLE"
          vProcessTable = False
        Case "CREATEATTRIBUTE", "ALTERATTRIBUTE"
          If vDocumentChange And vProcessTable Then
            vWriter.WriteLine("<TR>")
            If Not vHeading Then
              vWriter.WriteLine("<TD>" & vTableName & "</TD>")
              vHeading = True
            Else
              vWriter.WriteLine("<TD></TD>")
            End If
            vWriter.WriteLine("<TD>" & vItems(1) & "</TD>")
            vWhereFields(1).Value = vTableName
            vWhereFields(2).Value = vItems(1)
            Dim vAttrComment As String = vSQLAttribute.GetValue

            If vItems(0) = "CREATEATTRIBUTE" Then
              vWriter.WriteLine("<TD>" & vNewAttrText & vAttrComment & "</TD>")
            Else
              vWriter.WriteLine("<TD>" & vAlterAttrText & vAttrComment & "</TD>")
            End If
            vWriter.WriteLine("</TR>")
          End If
        Case "CREATEINDEX", "DROPINDEX"
          If vDocumentChange And vProcessTable Then
            vWriter.WriteLine("<TR>")
            If Not vHeading Then
              vWriter.WriteLine("<TD>" & vTableName & "</TD>")
              vHeading = True
            Else
              vWriter.WriteLine("<TD></TD>")
            End If
            vWriter.WriteLine("<TD></TD>")
            Dim vText As String = vItems(4) & "," & vItems(5) & "," & vItems(6) & "," & vItems(7)
            vText = vText.TrimEnd(","c).Replace(",", ", ")
            vWriter.WriteLine("<TD>" & IIf(vItems(0) = "CREATEINDEX", "Created", "Dropped").ToString & " " & vItems(1) & " index on " & vText & "</TD>") 'List of Attributes
            vWriter.WriteLine("</TR>")
          End If
      End Select
    End While
    vWriter.WriteLine("</TABLE>")

    'New Records in New Tables
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<h4>New Records in New Tables</h4>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<TABLE BORDER>")
    vWriter.WriteLine("<TR>")
    vWriter.WriteLine("<TH>Table Name</TH>")
    vWriter.WriteLine("<TH>Record Created</TH>")
    vWriter.WriteLine("<TH>Record Description</TH>")
    vWriter.WriteLine("</TR>")
    vReader.BaseStream.Seek(0, SeekOrigin.Begin)
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine().TrimStart
      Dim vItems() As String = vLine.Split(","c)
      Select Case vItems(0)
        Case "CHANGEINFO"
          If vItems(1).Trim = pVDesc Then vDocumentChange = True
        Case "ENDCHANGE"
          vDocumentChange = False
        Case "CREATETABLE"
          If vDocumentChange Then
            vTableName = vItems(1)
            vHeading = False
            vProcessTable = True
          End If
        Case "BEGINTABLE"
          vProcessTable = False
        Case "INSERT"
          If vDocumentChange And vProcessTable Then
            WriteInsertRecordData(vWriter, vReader, vTableName, vHeading)
          End If
      End Select
    End While
    vWriter.WriteLine("</TABLE>")

    'New Records in Existing Tables
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<h4>New Records in Existing Tables</h4>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<TABLE BORDER>")
    vWriter.WriteLine("<TR>")
    vWriter.WriteLine("<TH>Table Name</TH>")
    vWriter.WriteLine("<TH>Record Created</TH>")
    vWriter.WriteLine("<TH>Record Description</TH>")
    vWriter.WriteLine("</TR>")
    vReader.BaseStream.Seek(0, SeekOrigin.Begin)
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine().TrimStart
      Dim vItems() As String = vLine.Split(","c)
      Select Case vItems(0)
        Case "CHANGEINFO"
          If vItems(1).Trim = pVDesc Then vDocumentChange = True
        Case "ENDCHANGE"
          vDocumentChange = False
        Case "CREATETABLE"
          vProcessTable = False
        Case "BEGINTABLE"
          If vDocumentChange Then
            vTableName = vItems(1)
            vHeading = False
            vProcessTable = True
          End If
        Case "INSERT"
          If vDocumentChange And vProcessTable Then
            WriteInsertRecordData(vWriter, vReader, vTableName, vHeading)
          End If
      End Select
    End While
    vWriter.WriteLine("</TABLE>")

    'Updated Records
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<h4>Updated Records in Existing Tables</h4>")
    vWriter.WriteLine("<p>")
    vWriter.WriteLine("<TABLE BORDER>")
    vWriter.WriteLine("<TR>")
    vWriter.WriteLine("<TH>Table Name</TH>")
    vWriter.WriteLine("<TH>Record Updated</TH>")
    vWriter.WriteLine("<TH>New Value</TH>")
    vWriter.WriteLine("</TR>")
    vReader.BaseStream.Seek(0, SeekOrigin.Begin)
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine().TrimStart
      Dim vItems() As String = vLine.Split(","c)
      Select Case vItems(0)
        Case "CHANGEINFO"
          If vItems(1).Trim = pVDesc Then vDocumentChange = True
        Case "ENDCHANGE"
          vDocumentChange = False
        Case "CREATETABLE"
          vProcessTable = False
        Case "BEGINTABLE"
          If vDocumentChange Then
            vTableName = vItems(1)
            vHeading = False
            vProcessTable = True
          End If
        Case "UPDATE"
          If vDocumentChange And vProcessTable Then
            WriteInsertRecordData(vWriter, vReader, vTableName, vHeading)
          End If
      End Select
    End While
    vWriter.WriteLine("</TABLE>")

    vReader.Close()
    'Finish the page
    vWriter.WriteLine("<HR>")
    vWriter.WriteLine("<FN><CENTER>" & "This page was last updated automatically on:" & Date.Today & "</CENTER></FN>")
    vWriter.WriteLine("</BODY></HTML>")
    vWriter.Close()
    FileCopy(vFileName, pDestDir & "\" & pDestFilename)
    FileCopy(vFileName, "c:\temp\" & pDestFilename)
    ssl.Text = "Complete"
    sts.Refresh()
  End Sub

  Private Shared Sub WriteInsertRecordData(ByVal pWriter As StreamWriter, ByVal pReader As StreamReader, ByVal pTableName As String, ByRef pHeading As Boolean)
    pWriter.WriteLine("<TR>")
    If Not pHeading Then
      pWriter.WriteLine("<TD>" & pTableName & "</TD>")
      pHeading = True
    Else
      pWriter.WriteLine("<TD></TD>")
    End If
    'read next line for code
    Dim vLine As String = pReader.ReadLine.TrimStart
    If Not pReader.EndOfStream Then
      Dim vItems() As String = vLine.Split(","c)
      pWriter.WriteLine("<TD>" & vItems(2) & "</TD>")
      'read next line for description
      vLine = pReader.ReadLine.TrimStart
      If Not pReader.EndOfStream Then
        vItems = vLine.Split(","c)
        pWriter.WriteLine("<TD>" & vItems(2) & "</TD>")
      End If
    End If
    pWriter.WriteLine("</TR>")
  End Sub

  Private Enum WebServiceSummaryType
    DataAccess
    ExamsAccess
  End Enum

  Private Sub BuildWebServicesDocumentation(pWebServiceSummaryType As WebServiceSummaryType)
    'Build the Web Services documentation
    'Read the CDBDAXML.XML file 
    'For each method in the file create a documentation page
    Dim vSummaryWriter As StreamWriter = Nothing
    Dim vItemWriter As StreamWriter = Nothing
    Dim vXMLSourceFile As String
    Dim vWSSummaryFile As String
    If pWebServiceSummaryType = WebServiceSummaryType.ExamsAccess Then
      vXMLSourceFile = mvSourceLocation & "\WEB\CareServices\CDBEXAMS.XML"
      vWSSummaryFile = mvExamsWSSummaryFile
    Else
      vXMLSourceFile = mvSourceLocation & "\WEB\CareServices\CDBDAXML.XML"
      vWSSummaryFile = mvWSSummaryFile
    End If
    Try
      Dim vDoc As New XmlDocument
      vDoc.Load(vXMLSourceFile)
      Dim vNodeList As XmlNodeList = vDoc.GetElementsByTagName("method")

      If pWebServiceSummaryType = WebServiceSummaryType.ExamsAccess Then
        CheckWebServiceInXMLFile(vDoc, mvExamsWebServices, vNodeList)
      Else
        CheckCAREUSEONLY(mvWebServices, mvNetWebServices)

        CheckWebServiceInXMLFile(vDoc, mvWebServices, vNodeList)
        CheckWebServiceInXMLFile(vDoc, mvNetWebServices, vNodeList)
        CheckWebServiceInXMLFile(vDoc, mvWebWebServices, vNodeList)
      End If
      'My.Computer.FileSystem.DeleteFile("C:\MissingWebServices")

      CheckAllMethodsInWebServiceFile(vNodeList)

      vSummaryWriter = New StreamWriter(vWSSummaryFile)
      With vSummaryWriter
        .WriteLine("<HTML>")
        .WriteLine("<head>")
        Dim vPageTitle As String = "<title>{0} {1}WEB Services Summary {2}</title>"
        Dim vPageHeader As String = "<H1>{0} {1}WEB Services (Summary) Version {2}</H1>"
        If pWebServiceSummaryType = WebServiceSummaryType.ExamsAccess Then
          vPageTitle = String.Format(vPageTitle, mvProductName, "Exams ", mvVersionNumber)
          vPageHeader = String.Format(vPageHeader, mvProductName, "Exame ", mvVersionNumber)
          .WriteLine(vPageTitle)
        Else
          vPageTitle = String.Format(vPageTitle, mvProductName, "", mvVersionNumber)
          vPageHeader = String.Format(vPageHeader, mvProductName, "", mvVersionNumber)
          .WriteLine(vPageTitle)
        End If
        .WriteLine("<LINK href=""NGWEBServices.css"" type=""text/css"" rel=""stylesheet"">")
        .WriteLine("</head>")
        .WriteLine("<BODY>")
        .WriteLine(vPageHeader)
        If pWebServiceSummaryType <> WebServiceSummaryType.ExamsAccess Then
          .WriteLine("<BR>Differences from previous versions are shown <A HREF=""WEBServiceDifferences.htm"">here</A>")
        End If
        .WriteLine("<BR><BR>A summary of the available WEB Service operations is given below")
        .WriteLine("<TABLE border=1>")
        .WriteLine("<TR><TH>Operation</TH><TH>Description</TH></TR>")
      End With
      Dim vLastItemName As String = ""
      For Each vNode As XmlNode In vNodeList
        Dim vNotes As String
        Dim vItem As XmlNode = vNode.Attributes.GetNamedItem("WEBService")
        If vItem Is Nothing Then
          If mvWebServices.ContainsKey(vNode.Attributes.GetNamedItem("ID").InnerText) Then
            If Not vNode.Attributes.GetNamedItem("ID").InnerText = "CheckLicenseData" Then ShowError("Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is not marked as a WEB Service")
          ElseIf mvNetWebServices.ContainsKey(vNode.Attributes.GetNamedItem("ID").InnerText) Then
            If Not vNode.Attributes.GetNamedItem("ID").InnerText = "CheckLicenseData" Then ShowError("Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is not marked as a WEB Service")
          ElseIf mvExamsWebServices.ContainsKey(vNode.Attributes.GetNamedItem("ID").InnerText) Then
            ShowError("Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is not marked as a WEB Service")
          ElseIf vNode.Attributes.GetNamedItem("NotReleased") IsNot Nothing AndAlso vNode.Attributes.GetNamedItem("NotReleased").InnerText = "True" Then
            ShowError("NOT RELEASED: Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is not defined in DataAccess.asmx or NDataAccess.asmx.")
          Else
            ShowError("Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is not defined in DataAccess.asmx or NDataAccess.asmx")
          End If
        Else
          If Not vNode.Attributes.GetNamedItem("NotReleased") Is Nothing Then
            MsgBox("Item " & vNode.Attributes.GetNamedItem("ID").InnerText & " Is marked as WEB Service and Not Released", vbCritical)
          End If
          Dim vWebServiceName As String = vItem.InnerText
          Dim vWebService As Object
          If vWebServiceName.ToLower = "dataaccess.asmx" Then
            vWebService = mvCS
          ElseIf vWebServiceName.ToLower = "webaccess.asmx" Then
            vWebService = mvWA
          ElseIf vWebServiceName.ToLower = "examsdataaccess.asmx" Then
            vWebService = mvEA
          Else
            vWebService = mvNS
          End If

          Dim vItemName As String = vNode.Attributes.GetNamedItem("ID").InnerText
          If vItemName.ToLower < vLastItemName.ToLower Then
            ShowError(String.Format("Item {0} Is not in sequence", vItemName))
          End If
          vLastItemName = vItemName
          ssl.Text = String.Format("Processing {0}", vItemName)
          Debug.Print(String.Format("Processing {0}", vItemName))
          sts.Refresh()

          vItemWriter = New StreamWriter(String.Format("{0}\{1}.htm", mvWSDocumentationPath, vItemName))
          With vItemWriter
            .WriteLine("<HTML>")
            .WriteLine("<head>")
            .WriteLine(String.Format("<title>{0} WEB Services - Detail for {1}</title>", mvProductName, vItemName))
            .WriteLine("<LINK href=""NGWEBServices.css"" type=""text/css"" rel=""stylesheet"">")
            .WriteLine("</head>")
            .WriteLine("<BODY>")
            .WriteLine(String.Format("<H1>{0}</H1>", vItemName))
            .WriteLine("<TABLE border=1>")
            .WriteLine(String.Format("<TR><TD><B>Description</B></TD><TD>{0}</TD></TR>", vNode.Attributes.GetNamedItem("Description").InnerText))
            .WriteLine(String.Format("<TR><TD><B>Syntax</B></TD><TD>{0}</TD></TR>", vNode.Attributes.GetNamedItem("Syntax").InnerText))
          End With
          Dim vEnumType As String = ""
          Dim vHasReturns As Boolean = False

          If vNode.HasChildNodes Then
            For Each vChild As XmlNode In vNode.ChildNodes
              If vChild.Name = "type" Then
                For vIndex As Integer = 0 To vChild.ChildNodes.Count - 1
                  Dim vEnum As XmlNode = vChild.ChildNodes(vIndex)
                  If vEnum.Name = "enum" Then
                    vEnumType = vEnum.ChildNodes(0).Value.Trim
                    vItemWriter.WriteLine(String.Format("<TR><TD><B>{0}</B></TD><TD>", vEnumType))
                    If vEnumType = "HistoryActions" Or vEnumType = "PaymentPlanType" Then
                      vItemWriter.WriteLine("The type of data to be added<TABLE border =1>")
                    Else
                      vItemWriter.WriteLine("The type of data to be selected<TABLE border =1>")
                    End If
                    vItemWriter.WriteLine("<TR><TH>Name</TH><TH>Value</TH><TH>Description</TH></TR>")
                  ElseIf vEnum.Name = "enumvalue" Then
                    Dim vDesc As String = vEnum.Attributes("Description").InnerText
                    If vDesc.Length = 0 And vIndex = 1 Then
                      'Ignore none items
                    Else
                      vItemWriter.WriteLine("<TR><TD>" & vEnum.InnerText & "</TD><TD>" & vIndex - 1 & "</TD><TD>" & vDesc & "</TD></TR>")
                    End If
                  End If
                Next
                vItemWriter.WriteLine("</TABLE>")

              ElseIf vChild.Name = "parameters" Then
                If vChild.HasChildNodes Then
                  vItemWriter.WriteLine("<TR><TD><B>XML Parameters</B></TD><TD>")
                  vItemWriter.WriteLine("<TABLE border =1>")
                  vItemWriter.WriteLine("<TR><TH>Name</TH><TH>Data Type</TH><TH>Length</TH><TH>Mandatory</TH><TH>Notes</TH></TR>")
                  Dim vParameterList As New CDBParameters
                  For Each vParam As XmlNode In vChild.ChildNodes
                    vParameterList.Add(vParam.InnerText)
                  Next
                  Dim vXMLHelper As New XMLHelper
                  Try
                    vXMLHelper.ValidateParameterList(vParameterList, "", True, False, True)
                  Catch ex As Exception
                    ShowError(vItemName & ": " & ex.Message)
                  End Try
                  For Each vParam As XmlNode In vChild.ChildNodes
                    Dim vDataType As String = vParam.Attributes.GetNamedItem("DataType").InnerText
                    Select Case vDataType
                      Case "C"
                        vDataType = "String"
                      Case "L"
                        vDataType = "Long"
                      Case "D"
                        vDataType = "Date String"
                      Case "T"
                        vDataType = "DateTime String"
                      Case "M"
                        vDataType = "Multi Line String"
                      Case "N"
                        vDataType = "Decimal/Numeric"
                      Case Else
                        MsgBox("Unknown DataType " & vDataType, vbExclamation)
                    End Select

                    Dim vMandatory As String = vParam.Attributes.GetNamedItem("Mandatory").InnerText
                    Select Case vMandatory
                      Case "Y"
                        vMandatory = "Yes"
                      Case "N"
                        vMandatory = "No"
                    End Select
                    If vParam.Attributes.GetNamedItem("Notes") Is Nothing Then
                      vNotes = "&nbsp"
                    Else
                      vNotes = vParam.Attributes.GetNamedItem("Notes").InnerText
                    End If
                    vItemWriter.WriteLine("<TR><TD>" & vParam.InnerText & "</TD><TD>" & vDataType & "</TD><TD>" & vParameterList(vParam.InnerText).Value & "</TD><TD>" & vMandatory & "</TD><TD>" & vNotes & "</TD></TR>")
                  Next
                  vItemWriter.WriteLine("</TABLE>")
                Else
                  vItemWriter.WriteLine("<TR><TD><B>XML Parameters</B></TD><TD>none required")
                End If
              ElseIf vChild.Name = "resultset" Then
                Dim vXML As String = "<Parameters><UserLogname>care_admin</UserLogname>"
                For Each vParam As XmlNode In vChild.ChildNodes
                  vXML = vXML & "<" & vParam.InnerText & ">" & vParam.Attributes.GetNamedItem("Value").InnerText & "</" & vParam.InnerText & ">"
                Next
                vXML = vXML & "</Parameters>"
                Dim vResult As String = ""
                Try
                  If vEnumType.Length > 0 Then
                    vResult = CStr(CallByName(vWebService, vItemName, CallType.Method, 1, vXML))
                  Else
                    vResult = CStr(CallByName(vWebService, vItemName, CallType.Method, vXML))
                  End If
                  vNotes = GetResultItems(vResult)
                Catch ex As Exception
                  MessageBox.Show(ex.Message)
                  vNotes = ""
                End Try
                vItemWriter.WriteLine("<TR><TD><B>XML ResultSet</B></TD><TD>" & vNotes & "</TD></TR>")

              ElseIf vChild.Name = "returns" Then
                vHasReturns = True
                vItemWriter.WriteLine("<TR><TD><B>XML Return</B></TD><TD>")
                If vChild.ChildNodes.Count > 0 Then
                  vItemWriter.WriteLine("<TABLE border =1>")
                  vItemWriter.WriteLine("<TR><TH>Name</TH><TH>Data Type</TH><TH>Notes</TH></TR>")
                  For Each vParam As XmlNode In vChild.ChildNodes
                    Dim vDataType As String = vParam.Attributes.GetNamedItem("DataType").InnerText
                    Select Case vDataType
                      Case "C"
                        vDataType = "String"
                      Case "L"
                        vDataType = "Long"
                      Case "D"
                        vDataType = "Date String"
                      Case "T"
                        vDataType = "DateTime String"
                      Case "M"
                        vDataType = "Multi Line String"
                      Case "N"
                        vDataType = "Decimal/Numeric"
                      Case Else
                        MsgBox("Unknown DataType " & vDataType, vbExclamation)
                    End Select
                    If vParam.Attributes.GetNamedItem("Notes") Is Nothing Then
                      vNotes = "&nbsp"
                    Else
                      vNotes = vParam.Attributes.GetNamedItem("Notes").InnerText
                    End If
                    vItemWriter.WriteLine(String.Format("<TR><TD>{0}</TD><TD>{1}</TD><TD>{2}</TD></TR>", vParam.InnerText, vDataType, vNotes))
                  Next
                  vItemWriter.WriteLine("</TABLE>")
                Else
                  vItemWriter.WriteLine("No Return values")
                End If
              End If
            Next
          End If

          With vItemWriter
            .WriteLine("</TD></TR>")
            If Len(vEnumType) > 0 And vHasReturns = False Then
              .WriteLine("<TR><TD><B>XML Return</B></TD><TD>A result set element containing data rows with items appropriate to the selection. See the table below</TD></TR>")
            End If
            If vNode.Attributes.GetNamedItem("Notes") Is Nothing Then
              vNotes = "&nbsp"
            Else
              vNotes = vNode.Attributes.GetNamedItem("Notes").InnerText
            End If
            .WriteLine(String.Format("<TR><TD><B>Notes</B></TD><TD>{0}</TD></TR>", vNotes))
            .WriteLine("</TABLE>")
          End With

          If Len(vEnumType) > 0 AndAlso (vEnumType.EndsWith("DataSelectionType") OrElse vEnumType = "LookupDataType") Then
            vItemWriter.WriteLine("<BR><TABLE border =1><TR><TH>" & vEnumType & "</TH><TH>Items Returned</TH></TR>")
            Dim vTypeNode As XmlNode = vNode.SelectSingleNode("type")
            Dim vBaseXML As New StringBuilder
            For vIndex As Integer = 0 To vTypeNode.ChildNodes.Count - 1
              Dim vEnum As XmlNode = vTypeNode.ChildNodes(vIndex)
              If vEnum.Name = "enum" Then
                Dim vTestParams As XmlNodeList = vEnum.SelectNodes("testparam")
                For Each vTestParam As XmlNode In vTestParams
                  vBaseXML.Append("<")
                  vBaseXML.Append(vTestParam.InnerText)
                  vBaseXML.Append(">")
                  vBaseXML.Append(vTestParam.Attributes("Value").InnerText)
                  vBaseXML.Append("</")
                  vBaseXML.Append(vTestParam.InnerText)
                  vBaseXML.Append(">")
                Next
              ElseIf vEnum.Name = "enumvalue" Then
                Dim vDesc As String = vEnum.Attributes("Description").InnerText
                If vDesc.Length = 0 And vIndex < 3 Then
                  'Ignore none items
                Else
                  Dim vXML As New StringBuilder("<Parameters><DocumentColumns>Y</DocumentColumns>")
                  Dim vEnumName As String = vEnum.ChildNodes(0).Value.Trim
                  Dim vTestParams As XmlNodeList = vEnum.SelectNodes("testparam")
                  Dim vBaseParameters As String = vBaseXML.ToString
                  For Each vTestParam As XmlNode In vTestParams
                    If vBaseParameters.Contains(vTestParam.InnerText) Then
                      If vTestParam.Attributes.Count > 0 Then
                        Dim vPos1 As Integer = vBaseParameters.IndexOf(vTestParam.InnerText) + vTestParam.InnerText.Length + 1
                        Dim vPos2 As Integer = vBaseParameters.IndexOf(vTestParam.InnerText, vPos1) - 2
                        vBaseParameters = vBaseParameters.Replace(vBaseParameters.Substring(vPos1, vPos2 - vPos1), vTestParam.Attributes("Value").InnerText)
                      Else
                        Dim vPos1 As Integer = vBaseParameters.IndexOf(vTestParam.InnerText) - 1
                        Dim vPos2 As Integer = vBaseParameters.IndexOf(vTestParam.InnerText, vPos1 + 2) + vTestParam.InnerText.Length + 1
                        vBaseParameters = vBaseParameters.Remove(vPos1, vPos2 - vPos1)
                      End If
                    Else
                      vXML.Append("<")
                      vXML.Append(vTestParam.InnerText)
                      vXML.Append(">")
                      vXML.Append(vTestParam.Attributes("Value").InnerText)
                      vXML.Append("</")
                      vXML.Append(vTestParam.InnerText)
                      vXML.Append(">")
                    End If
                  Next
                  vXML.Append(vBaseParameters)
                  If Not vXML.ToString.Contains("<UserLogname>") Then vXML.Append("<UserLogname>care_admin</UserLogname>")

                  vXML.Append("</Parameters>")
                  Dim vResult As String = ""
                  Try
                    If vEnumName = "xldtGenericLookup" Then
                      vItemWriter.WriteLine("<TR><TD><B>" & vEnumName & "</B></TD><TD>Return values are dependant on supplied parameters</TD></TR>")
                    ElseIf vEnumName = "xcdtContactDashboard" Then
                      vItemWriter.WriteLine("<TR><TD><B>" & vEnumName & "</B></TD><TD>Deprecated</TD></TR>")
                    Else
                      'If vItemName = "SelectContactData" AndAlso vEnumName = "xcdtContactLegacy" Then Stop ' AndAlso vDesc = "xcdtContactLegacy" Then Stop
                      vResult = CStr(CallByName(vWebService, vItemName, CallType.Method, vIndex - 1, vXML.ToString))
                      If vResult.Contains("<ErrorMessage>") Then
                        ShowError("Enum value: " & vEnumName & " Got Error" & vbCrLf & vbCrLf & vResult)
                        vItemWriter.WriteLine("<TR><TD><B>" & vEnumName & "</B></TD><TD></TD></TR>")
                      Else
                        vNotes = GetResultItems(vResult)
                        vItemWriter.WriteLine("<TR><TD><B>" & vEnumName & "</B></TD><TD>" & vNotes & "</TD></TR>")
                        If Len(vNotes) = 0 Then ShowError("Enum value: " & vEnumName & " Gave no Results")
                      End If
                    End If
                  Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    vNotes = ""
                  End Try
                End If
              End If
            Next
            vItemWriter.WriteLine("</TABLE>")
          End If

          With vItemWriter
            .WriteLine("<P><HR>")
            If pWebServiceSummaryType = WebServiceSummaryType.ExamsAccess Then
              .WriteLine(String.Format("<A HREF=""ExamsWEBServicesSummary.HTM#{0}_index"">{1} Exams WEB Services Summary</A>", vItemName, mvProductName))
            Else
              .WriteLine(String.Format("<A HREF=""WEBServicesSummary.HTM#{0}_index"">{1} WEB Services Summary</A>", vItemName, mvProductName))
            End If
            .WriteLine("<HR>")
            .WriteLine(String.Format("&#169 {0}, 2014<br>", mvCopyrightName))
            .WriteLine("All Rights Reserved")
            .WriteLine("<HR>")
            .WriteLine("Last Updated: " & Now.ToShortDateString)
            .WriteLine("</BODY>")
            .WriteLine("</HTML>")
          End With
          vItemWriter.Close()
          vItemWriter = Nothing
          If vWebServiceName.ToLower <> "webaccess.asmx" Then
            vSummaryWriter.WriteLine(String.Format("<TR><TD><A NAME=""{0}_index""></A>", vItemName))
            vSummaryWriter.WriteLine(String.Format("<A HREF=""{0}.htm"">{0}</A></TD><TD>{1}</TD></TR>", vItemName, vNode.Attributes.GetNamedItem("Description").InnerText))
          End If
        End If
      Next
      vSummaryWriter.Close()
      vSummaryWriter = Nothing
      ssl.Text = String.Format("Web Services Processing Complete")

    Catch ex As Exception
      If vSummaryWriter IsNot Nothing Then vSummaryWriter.Close()
      If vItemWriter IsNot Nothing Then vItemWriter.Close()
      Throw ex
    End Try
  End Sub

  Private Sub CheckCAREUSEONLY(ByVal pWebServices As SortedList(Of String, WebServiceData), ByVal pNetWebServices As SortedList(Of String, WebServiceData))
    For Each vWSItem As WebServiceData In pWebServices.Values
      If pNetWebServices.ContainsKey(vWSItem.Name) Then
        If vWSItem.Description.StartsWith("NOT CURRENTLY SUPPORTED") AndAlso Not pNetWebServices(vWSItem.Name).Description.StartsWith("NOT CURRENTLY SUPPORTED") Then
          MessageBox.Show(String.Format("Web Service {0} Not {1} USE ONLY In .NET Web Services", vWSItem.Name, mvProductName))
        End If
      End If
    Next
  End Sub

  Private Sub CheckAllMethodsInWebServiceFile(ByVal pNodeList As XmlNodeList)
    For Each vNode As XmlNode In pNodeList
      If vNode.Attributes.GetNamedItem("WebServiceFound") Is Nothing Then
        ShowError("XML File refers to method " & vNode.Attributes.GetNamedItem("ID").InnerText & " missing from Web services")
      End If
    Next
  End Sub

  Private Sub CheckWebServiceInXMLFile(pDoc As XmlDocument, ByVal pWebServices As SortedList(Of String, WebServiceData), ByVal pNodeList As XmlNodeList)
    For Each vWSItem As WebServiceData In pWebServices.Values
      Dim vFound As Boolean = False
      For Each vNode As XmlNode In pNodeList
        If vNode.Attributes.GetNamedItem("ID").InnerText = vWSItem.Name Then
          If vNode.Attributes.GetNamedItem("WebServiceFound") Is Nothing Then
            Dim vAttr As XmlAttribute = pDoc.CreateAttribute("WebServiceFound")
            vAttr.Value = "Y"
            vNode.Attributes.Append(vAttr)
          End If
          vFound = True
          Exit For
        End If
      Next
      If Not vFound AndAlso Not vWSItem.Description.StartsWith("DEPRECATED") AndAlso Not vWSItem.Description.StartsWith("NOT CURRENTLY SUPPORTED") Then
        'My.Computer.FileSystem.WriteAllText("c:\MissingWebServices", vWSItem.Key & vbCrLf, True)
        ShowError("Web Service " & vWSItem.Name & " missing from XML file")
      End If
    Next
  End Sub

  Private Sub BuildWebServiceDifferences()
    Dim vFileName As String = mvWSDocumentationPath & "\WEBServiceDifferences.HTM"
    Dim vWriter As StreamWriter = New StreamWriter(vFileName)
    With vWriter

      .WriteLine("<HTML>")
      .WriteLine("<head>")
      .WriteLine(String.Format("<title>{0} WEB Service Differences Summary {1}</title>", mvProductName, mvVersionNumber))
      .WriteLine("<LINK href=""NGWEBServices.css"" type=""text/css"" rel=""stylesheet"">")
      .WriteLine("</head>")
      .WriteLine("<BODY>")
      .WriteLine(String.Format("<H1>{0} WEB Services Differences</H1>", mvProductName))

      Dim vSourceDirectoryPath As String = "D:\"
      If My.Settings.UseOldSourceFromNTDEV4 Then vSourceDirectoryPath = "\\ntdev4\D$\"
      Dim vSourceDirectory As String = vSourceDirectoryPath & "GUI BUILD{0} SOURCE\WEB\CareServices\Documentation\WEBServicesSummary.htm"
      If Not New FileInfo(String.Format(vSourceDirectory, "54")).Exists Then
        vSourceDirectory = "\\NTDEV4\GUI{0}SOURCE\WEB\CareServices\Documentation\WEBServicesSummary.htm"
      End If
      Dim vNewVersion As Integer = mvMajorVersion * 10 + mvMinorVersion
      'Build 7.2 to latest differences
      WriteDifferencePage(vWriter, 72, vNewVersion, vSourceDirectory)
      'Build 11.2 onwards differences
      If mvMajorVersion > 11 OrElse (mvMajorVersion = 11 AndAlso mvMinorVersion = 3) Then
        Dim vLastVersion As Integer = If(mvMinorVersion = 1, (mvMajorVersion - 1) * 10 + 3, vNewVersion - 1)
        For vVersion As Integer = 112 To vLastVersion
          If vVersion Mod 10 = 4 Then vVersion += 10 - 3
          Select Case vVersion
            Case 123, 133
              'Do nothing as 12.3 & 13.3 are not valid releases
            Case Else
              WriteDifferencePage(vWriter, vVersion, vNewVersion, vSourceDirectory)
          End Select
          'If vVersion <> 123 Then WriteDifferencePage(vWriter, vVersion, vNewVersion, vSourceDirectory)
        Next
      End If
      .WriteLine("<p><HR>")
      .WriteLine(String.Format("&#169 {0}, 2014<br>", mvCopyrightName))
      .WriteLine("All Rights Reserved")
      .WriteLine("<HR>")
      .WriteLine("Last Updated: " & Now)
      .WriteLine("</BODY>")
      .WriteLine("</HTML>")
    End With
    vWriter.Close()
    vWriter = Nothing
    ssl.Text = String.Format("Processing Complete")
  End Sub

  Private Sub WriteDifferencePage(ByVal pStreamWriter As StreamWriter, ByVal pVersion As Integer, ByVal pNewVersion As Integer, ByVal pSourceDirectory As String)
    Dim vVersionString As String
    Dim vNewVersionString As String

    If pVersion > 110 Then
      vVersionString = CStr(pVersion \ 10) & "R" & CStr(pVersion Mod 10)
    Else
      vVersionString = pVersion.ToString
    End If
    If pNewVersion > 110 Then
      vNewVersionString = CStr(pNewVersion \ 10) & "R" & CStr(pNewVersion Mod 10)
    Else
      vNewVersionString = pNewVersion.ToString
    End If
    Dim vDifferenceName As String = String.Format("Differences{0}to{1}.HTM", pVersion, pNewVersion)
    pStreamWriter.Write("<BR><A HREF=""{0}"">", vDifferenceName)
    pStreamWriter.WriteLine("Differences {0}.{1} to {2}</A>", pVersion \ 10, pVersion Mod 10, mvVersionNumber)
    MakeDifferencePages(String.Format(pSourceDirectory, vVersionString), _
                mvWSSummaryFile, mvWSDocumentationPath & "\" & vDifferenceName, String.Format("{0}.{1}", pVersion \ 10, pVersion Mod 10), mvVersionNumber)
  End Sub

  Private Function GetResultItems(ByVal pResult As String) As String
    Dim vResultDoc As New XmlDocument
    vResultDoc.LoadXml(pResult)
    Dim vResultList As XmlNodeList = vResultDoc.GetElementsByTagName("DataRow")
    Dim vResultNode As XmlNode = vResultList(0)
    Dim vItems As New StringBuilder
    If Not vResultNode Is Nothing Then
      For Each vResultChild As XmlNode In vResultNode.ChildNodes
        If vItems.Length > 0 Then vItems.Append(", ")
        vItems.Append(vResultChild.Name)
      Next
    End If
    Return vItems.ToString
  End Function

  Private Sub MakeDifferencePages(ByVal pOldFile As String, ByVal pNewFile As String, ByVal pDiffFileName As String, ByVal pOldVersion As String, ByVal pNewVersion As String)
    Dim vWriter As StreamWriter = Nothing
    Try
      Dim vOldFileInfo As FileInfo = New FileInfo(pOldFile)
      Dim vNewFileInfo As FileInfo = New FileInfo(pNewFile)
      'Read methods from latest file.
      If Not vNewFileInfo.Exists Then
        MessageBox.Show("WEBServicesSummary.HTM file not found: " & pNewFile)
        Exit Sub
      End If
      If Not vOldFileInfo.Exists Then
        MessageBox.Show("WEBServicesSummary.HTM file not found: " & pOldFile)
        Exit Sub
      End If
      ssl.Text = "Reading new summary"
      sts.Refresh()
      Dim vNewServices As List(Of String) = GetWebServiceMethods(pNewFile)
      ssl.Text = "Reading old summary"
      sts.Refresh()
      Dim vOldServices As List(Of String) = GetWebServiceMethods(pOldFile)

      vWriter = New StreamWriter(pDiffFileName)
      With vWriter
        .WriteLine("<HTML>")
        .WriteLine("<head>")
        .WriteLine(String.Format("<title>{0} WEB Services Differences</title>", mvProductName))
        .WriteLine("<LINK href=""NGWEBServices.css"" type=""text/css"" rel=""stylesheet"">")
        .WriteLine("</head>")
        .WriteLine("<BODY>")
        .WriteLine(String.Format("<H1>{0} WEB Services (Differences)</H1>", mvProductName))
        .WriteLine(String.Format("A list of the differences between the WEB Services provided in versions {0} and {1} is given below<BR><BR>", pOldVersion, pNewVersion))
        .WriteLine("<TABLE border=1>")
        .WriteLine("<TR><TH>Operation</TH><TH>Description</TH></TR>")
        For Each vItem As String In vNewServices
          If Not vOldServices.Contains(vItem) Then
            .WriteLine("<TR><TD><A NAME=""" & vItem & "_index""></A>")
            .WriteLine("<A HREF=""" & vItem & ".htm"">" & vItem & "</A></TD><TD>has been <font color=""#00CC66"">ADDED</font> for this version.</TD></TR>")
          End If
        Next
        For Each vItem As String In vOldServices
          If Not vNewServices.Contains(vItem) Then
            .WriteLine("<TR><TD><A NAME=""" & vItem & "_index""></A>")
            .WriteLine(String.Format("<A HREF=""" & vItem & ".htm"">" & vItem & "</A></TD><TD>has been <font color=""#CC3333"">REMOVED</font> from this version. (May be {0} use only or previously documented in error)</TD></TR>", mvProductName))
          End If
        Next
        .WriteLine("</TABLE>")
        .WriteLine("<P><HR>")
        .WriteLine(String.Format("A list of the differences between the WEB Service provided in versions {0} and {1} is given below<BR><BR>", pOldVersion, pNewVersion))

        For Each vItem As String In vNewServices
          If vOldServices.Contains(vItem) Then
            ssl.Text = String.Format("Reading New {0}", vItem)
            sts.Refresh()
            Dim vNew As New WebServiceData(vItem, "")
            GetWebServiceItems(vNewFileInfo.DirectoryName & "\" & vItem & ".htm", vNew)
            ssl.Text = String.Format("Reading Old {0}", vItem)
            sts.Refresh()
            Dim vOld As New WebServiceData(vItem, "")
            Dim vIgnore As Boolean = False
            Try
              GetWebServiceItems(vOldFileInfo.DirectoryName & "\" & vItem & ".htm", vOld)
            Catch vEx As Exception
              ShowError(String.Format("File {0} not found", vOldFileInfo.DirectoryName & "\" & vItem & ".htm"))
              vIgnore = True
            End Try
            If Not vIgnore Then
              'Process parameters, returns, notes and enum returns
              Dim vStdHeader As New StringBuilder
              vStdHeader.Append("<A HREF=""")
              vStdHeader.Append(vItem & ".htm")
              vStdHeader.Append(""">")
              vStdHeader.Append(vItem)
              vStdHeader.Append("</A>&nbsp;")
              vStdHeader.Append("Parameters")
              vStdHeader.Append("<TABLE width = ""800"" border=1>")
              Dim vHeader As New StringBuilder
              vHeader.Append(vStdHeader)
              vHeader.Append("<TR><TH width = ""15%"">Name</TH><TH width = ""10%"">Data Type</TH><TH width = ""10%"">Length</TH><TH width = ""15%"">Mandatory</TH><TH>Notes</TH><TH width = ""15%"">Difference</TH></TR>")
              WriteDifferences(vWriter, SectionTypes.Parameters, vNew.Parameters, vOld.Parameters, vHeader)

              vHeader = New StringBuilder
              vHeader.Append(vStdHeader.Replace("Parameters", "Returns"))
              vHeader.Append("<TR><TH width = ""15%"">Name</TH><TH width = ""10%"">Data Type</TH><TH>Notes</TH><TH width = ""15%"">Difference</TH></TR>")
              WriteDifferences(vWriter, SectionTypes.Returns, vNew.Returns, vOld.Returns, vHeader)

              vHeader = New StringBuilder
              vHeader.Append(vStdHeader.Replace("Parameters", "Notes"))
              vHeader.Append("<TR><TH>Notes</TH><TH width = ""15%"">Difference</TH></TR>")
              If vNew.Notes.Count > 0 And vOld.Notes.Count = 0 Then
                vWriter.WriteLine(vHeader)
                vWriter.WriteLine("<TR><TD>" & vNew.Notes(0) & "</TD><TD><font color=""#00CC66"">ADDED</font></TD></TR>")
                vWriter.WriteLine("</TABLE>")
                vWriter.WriteLine("<P>")
              ElseIf vNew.Notes.Count = 0 And vOld.Notes.Count > 0 Then
                vWriter.WriteLine(vHeader)
                vWriter.WriteLine("<TR><TD>" & vOld.Notes(0) & "</TD><TD><font color=""#CC3333"">DELETED</font></TD></TR>")
                vWriter.WriteLine("</TABLE>")
                vWriter.WriteLine("<P>")
              ElseIf vNew.Notes.Count > 0 And vOld.Notes.Count > 0 Then
                If vNew.Notes(0) <> vOld.Notes(0) Then
                  vWriter.WriteLine(vHeader)
                  vWriter.WriteLine("</TD><TD>" & "<strong>OLD: </strong>" & vOld.Notes(0) & "<BR><BR><strong>NEW: </strong>" & vNew.Notes(0) & "</TD><TD><font color=""#CC9900"">CHANGED</font><BR></TD></TR>")
                  vWriter.WriteLine("</TABLE>")
                  vWriter.WriteLine("<P>")
                End If
              End If

              vHeader = New StringBuilder
              vHeader.Append(vStdHeader.Replace("Parameters", "Enumerations"))
              vHeader.Append("<TR><TH width = ""15%"">Enumeration</TH><TH>Items Returned</TH><TH width = ""15%"">Difference</TH></TR>")
              WriteDifferences(vWriter, SectionTypes.EnumReturns, vNew.EnumReturns, vOld.EnumReturns, vHeader)
            End If
          End If
        Next
        .WriteLine(String.Format("&#169 {0}, 2003-2014<BR>", mvCopyrightName))
        .WriteLine("All Rights Reserved")
        .WriteLine("<HR>")
        .WriteLine("Last Updated: " & Now)
        .WriteLine("</BODY>")
        .WriteLine("</HTML>")
        .Close()
      End With
    Catch ex As Exception
      If vWriter IsNot Nothing Then vWriter.Close()
      Throw ex
    End Try
  End Sub

  Private Shared Sub WriteDifferences(ByVal pWriter As StreamWriter, ByVal pSectionType As SectionTypes, ByVal vNew As SortedList(Of String, String()), ByVal vOld As SortedList(Of String, String()), ByVal vHeader As StringBuilder)
    Dim vWrittenHeader As Boolean = False

    For Each vParam As KeyValuePair(Of String, String()) In vNew
      If Not vOld.ContainsKey(vParam.Key) Then
        If Not vWrittenHeader Then
          pWriter.WriteLine(vHeader)
          vWrittenHeader = True
        End If
        pWriter.Write("<TR><TD>")     ' New Added
        pWriter.Write(vParam.Key)
        pWriter.Write("</TD><TD>")
        pWriter.Write(vParam.Value(1))
        If pSectionType = SectionTypes.Parameters OrElse pSectionType = SectionTypes.Returns Then
          pWriter.Write("</TD><TD>")
          pWriter.Write(vParam.Value(2))
          If pSectionType = SectionTypes.Parameters Then
            pWriter.Write("</TD><TD>")
            pWriter.Write(vParam.Value(3))
            pWriter.Write("</TD><TD>")
            pWriter.Write(vParam.Value(4))
          End If
        End If
        pWriter.WriteLine("</TD><TD><font color=""#00CC66"">ADDED</font></TD></TR>")
      Else
        Dim vParameterChanged As Integer = 0
        Dim vLine As New StringBuilder
        Dim vOldParam As String() = vOld(vParam.Key)
        Dim vItemCount As Integer
        Select Case pSectionType
          Case SectionTypes.Parameters
            vItemCount = 4
          Case SectionTypes.Returns
            vItemCount = 2
          Case SectionTypes.EnumReturns
            vItemCount = 1
        End Select
        For vLoop As Integer = 1 To vItemCount
          If vParam.Value(vLoop) <> vOldParam(vLoop) Then
            vLine.Append("</TD><TD><strong>OLD: </strong>" & vOldParam(vLoop) & "<BR><BR><strong>NEW: </strong>" & vParam.Value(vLoop))
            vParameterChanged = CInt(vParameterChanged + 2 ^ (vLoop - 1))
          Else
            vLine.Append("</TD><TD>")
            If vParam.Value(vLoop).Length > 0 Then
              vLine.Append(vParam.Value(vLoop))
            Else
              vLine.Append("&nbsp;")
            End If
          End If
        Next
        vLine.Append("</TD><TD>")
        If vParameterChanged > 0 Then vLine.Append("<font color=""#CC9900"">")
        Select Case pSectionType
          Case SectionTypes.Parameters
            If (vParameterChanged And 1) > 0 Then vLine.Append("Data Type, ")
            If (vParameterChanged And 2) > 0 Then vLine.Append("Length, ")
            If (vParameterChanged And 4) > 0 Then vLine.Append("Mandatory, ")
            If (vParameterChanged And 8) > 0 Then vLine.Append("Notes, ")
            If vParameterChanged > 0 Then vLine.Remove(vLine.Length - 2, 2)
          Case SectionTypes.Returns
            If (vParameterChanged And 1) > 0 Then vLine.Append("Data Type, ")
            If (vParameterChanged And 2) > 0 Then vLine.Append("Notes, ")
            If vParameterChanged > 0 Then vLine.Remove(vLine.Length - 2, 2)
        End Select
        If vParameterChanged > 0 Then vLine.Append(" CHANGED</font>")
        vLine.Append("</TD></TR>")
        If vParameterChanged > 0 Then
          If Not vWrittenHeader Then
            pWriter.WriteLine(vHeader)
            vWrittenHeader = True
          End If
          pWriter.Write("<TR><TD>")
          pWriter.Write(vParam.Key)
          pWriter.WriteLine(vLine)
        End If
      End If
    Next
    For Each vParam As KeyValuePair(Of String, String()) In vOld
      If Not vNew.ContainsKey(vParam.Key) Then
        If Not vWrittenHeader Then
          pWriter.WriteLine(vHeader)
          vWrittenHeader = True
        End If
        pWriter.Write("<TR><TD>")     'Deleted
        pWriter.Write(vParam.Key)
        pWriter.Write("</TD><TD>")
        pWriter.Write(vParam.Value(1))
        If pSectionType = SectionTypes.Parameters OrElse pSectionType = SectionTypes.Returns Then
          pWriter.Write("</TD><TD>")
          pWriter.Write(vParam.Value(2))
          If pSectionType = SectionTypes.Parameters Then
            pWriter.Write("</TD><TD>")
            pWriter.Write(vParam.Value(3))
            pWriter.Write("</TD><TD>")
            pWriter.Write(vParam.Value(4))
          End If
        End If
        pWriter.WriteLine("</TD><TD><font color=""#CC3333"">DELETED</font></TD></TR>")
      End If
    Next
    If vWrittenHeader Then
      pWriter.WriteLine("</TABLE>")
      pWriter.WriteLine("<P>")
    End If
  End Sub

  Private Enum SectionTypes
    None
    Parameters
    Returns
    Notes
    EnumReturns
  End Enum

  Private Sub GetWebServiceItems(ByVal pFilename As String, ByVal pWebServiceData As WebServiceData)
    Dim vSectionType As SectionTypes = SectionTypes.None
    Dim vReader As New StreamReader(pFilename)
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine

      If vLine.StartsWith("<TR><TD><B>XML Parameters</B></TD><TD>") Then
        vSectionType = SectionTypes.Parameters    'Parameters section
      ElseIf vLine.StartsWith("<TR><TD><B>XML Return</B></TD><TD>") Then
        vSectionType = SectionTypes.Returns       'Returns section
      ElseIf vLine.StartsWith("<TR><TD><B>Notes</B></TD><TD>") Then
        vSectionType = SectionTypes.Notes         'Notes section
      ElseIf vLine.StartsWith("<BR><TABLE border =1><TR><TH>") Then
        vSectionType = SectionTypes.EnumReturns   'Items returned (enums)
      End If

      Select Case vSectionType
        Case SectionTypes.Parameters
          If vLine <> "<TR><TD><B>XML Parameters</B></TD><TD>none required" Then
            vLine = vReader.ReadLine
            vLine = vReader.ReadLine
            vLine = vReader.ReadLine
            While vLine <> "</TABLE>"
              vLine = Mid$(vLine, 9)
              vLine = Replace(vLine, "<TD></TD>", "<TD>&nbsp;</TD>")
              vLine = Replace(vLine, "</TD>", "")
              vLine = Replace(vLine, "</TR>", "")
              Dim vItems() As String = Split(vLine, "<TD>")
              pWebServiceData.AddParameterData(vItems)
              vLine = vReader.ReadLine
            End While
          End If
          vSectionType = SectionTypes.None

        Case SectionTypes.EnumReturns
          vLine = vReader.ReadLine
          While vLine <> "</TABLE>"
            vLine = Mid$(vLine, 12)
            vLine = Replace(vLine, "</B>", "")
            vLine = Replace(vLine, "<TD></TD>", "<TD>&nbsp;</TD>")
            vLine = Replace(vLine, "</TD>", "")
            vLine = Replace(vLine, "</TR>", "")
            Dim vItems() As String = Split(vLine, "<TD>")
            pWebServiceData.AddEnumReturnData(vItems)
            vLine = vReader.ReadLine
          End While
          vSectionType = SectionTypes.None

        Case SectionTypes.Notes
          vLine = Mid$(vLine, 30, Len(vLine) - 30 - 9)
          pWebServiceData.AddNotes(vLine)
          vSectionType = SectionTypes.None

        Case SectionTypes.Returns
          If vLine.Contains("A result set element") Then
            '
          Else
            vLine = vReader.ReadLine
            If vLine <> "No Return values" Then
              vLine = vReader.ReadLine
              If vLine <> "</TABLE>" Then
                vLine = vReader.ReadLine
                While vLine <> "</TABLE>"
                  vLine = Mid$(vLine, 9)
                  vLine = Replace(vLine, "</TD>", "")
                  vLine = Replace(vLine, "</TR>", "")
                  Dim vItems() As String = Split(vLine, "<TD>")
                  pWebServiceData.AddReturnData(vItems)
                  vLine = vReader.ReadLine
                End While
              End If
            End If
          End If
          vSectionType = SectionTypes.None
      End Select
    End While
    vReader.Close()
  End Sub

  Private Function GetWebServiceMethods(ByVal pFileName As String) As List(Of String)
    Dim vReader As New StreamReader(pFileName)
    Dim vServices As New List(Of String)
    While vReader.EndOfStream = False
      Dim vLine As String = vReader.ReadLine
      If vLine.StartsWith("<A HREF=") Then
        Dim vStart As Integer = vLine.IndexOf(">", 8) + 1
        Dim vEnd As Integer = vLine.IndexOf("<", vStart)
        Dim vWebService As String = vLine.Substring(vStart, vEnd - vStart)
        vServices.Add(vWebService)
      End If
    End While
    vReader.Close()
    Return vServices
  End Function

  Private Class WebServiceData

    Friend Enum CallTypes
      None
      AddRecord
      UpdateRecord
      DeleteRecord
    End Enum

    Private mvName As String
    Private mvDesc As String
    Private mvParameters As New SortedList(Of String, String())
    Private mvReturns As New SortedList(Of String, String())
    Private mvEnumReturns As New SortedList(Of String, String())
    Private mvNotes As New List(Of String)
    Public Syntax As String
    Public CallType As CallTypes
    Public ClassName As String

    Public Sub AddParameterData(ByVal pData As String())
      Try
        mvParameters.Add(pData(0), pData)
      Catch vEx As Exception
        MessageBox.Show(String.Format("Web Service {0} AddParameter {1} Error: {2}", mvName, pData(0), vEx.Message))
      End Try
    End Sub
    Public Sub AddReturnData(ByVal pData As String())
      Try
        mvReturns.Add(pData(0), pData)
      Catch vEx As Exception
        MessageBox.Show(String.Format("Web Service {0} AddReturnData {1} Error: {2}", mvName, pData(0), vEx.Message))
      End Try
    End Sub
    Public Sub AddEnumReturnData(ByVal pData As String())
      Try
        mvEnumReturns.Add(pData(0), pData)
      Catch vEx As Exception
        MessageBox.Show(String.Format("Web Service {0} AddEnumReturnData {1} Error: {2}", mvName, pData(0), vEx.Message))
      End Try
    End Sub
    Public Sub AddNotes(ByVal pData As String)
      mvNotes.Add(pData)
    End Sub

    Public ReadOnly Property Parameters() As SortedList(Of String, String())
      Get
        Return mvParameters
      End Get
    End Property
    Public ReadOnly Property Returns() As SortedList(Of String, String())
      Get
        Return mvReturns
      End Get
    End Property
    Public ReadOnly Property EnumReturns() As SortedList(Of String, String())
      Get
        Return mvEnumReturns
      End Get
    End Property
    Public ReadOnly Property Notes() As List(Of String)
      Get
        Return mvNotes
      End Get
    End Property
    Public ReadOnly Property Name() As String
      Get
        Return mvName
      End Get
    End Property
    Public ReadOnly Property Description() As String
      Get
        Return mvDesc
      End Get
    End Property
    Public Sub New(ByVal pName As String, pDesc As String)
      mvName = pName
      mvDesc = pDesc
      CallType = CallTypes.None
    End Sub
  End Class

  Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
    Dim vFileName As String = txtJira.Text
    With ofdJira
      .CheckFileExists = True
      .CheckPathExists = True
      .Multiselect = False
      .Title = "Select Jira csv file"
      .Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      .FileName = vFileName
      If .ShowDialog() = Windows.Forms.DialogResult.OK Then
        txtJira.Text = .FileName
      End If
    End With
  End Sub

  Private Sub chkBuildVersion_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBuildVersion.CheckStateChanged
    If chkBuildVersion.Checked = False Then
      chkJira.Checked = False
      txtJira.Text = ""
    End If
    chkJira.Enabled = chkBuildVersion.Checked
    txtJira.Enabled = chkBuildVersion.Checked
    cmdBrowse.Enabled = chkBuildVersion.Checked
  End Sub

  Private Sub cmdExams_Click(sender As System.Object, e As System.EventArgs) Handles cmdExams.Click
    Dim vEnv As New CDBEnvironment("CDBWEBSERVER", "care_admin", "care_admin")

    mvExamsWebServices = New SortedList(Of String, WebServiceData)
    ReadWebServiceFile(mvExamsWebServices, mvSourceLocation & "\web\Careservices\examsdataaccess.asmx.vb")


    Dim vLogFile As String = mvSourceLocation & "\WEB\CareServices\NOEXAMDOCS.TXT"
    Dim vLogWriter As New StreamWriter(vLogFile)

    Dim vWriter As StreamWriter = Nothing
    Dim vXMLSourceFile As String = mvSourceLocation & "\WEB\CareServices\CDBEXAMS.XML"
    vWriter = New StreamWriter(vXMLSourceFile)
    With vWriter
      .WriteLine("<methods>")
      For Each vWebServiceData As WebServiceData In mvExamsWebServices.Values
        .WriteLine(String.Format("  <method ID=""{0}"" Description=""{1}"" Syntax=""XMLReturn = {2}"" WebService=""ExamsDataAccess.asmx"">", vWebServiceData.Name, vWebServiceData.Description, vWebServiceData.Syntax))
        .Write("    ")
        .WriteLine(GetMethodNameDesc(vWebServiceData.Name))
        .WriteLine("    <parameters>")
        .WriteLine("      <parameter DataType=""C"" Mandatory=""N"" Notes=""Defines which CBS.INI entry will be used - Default is CDBWEBSERVER"">Database</parameter>")
        .WriteLine("      <parameter DataType=""C"" Mandatory=""N"" Notes=""Defines which Users record will be used - Default is guest"">UserLogname</parameter>")
        .WriteLine("      <parameter DataType=""C"" Mandatory=""N"" Notes=""Defines the value to be used as amended_by for anonymous users - Defaults to the UserLogname - Should be set to a contact number"">UserID</parameter>")
        Dim vWriteEndParameters As Boolean = True
        If Not String.IsNullOrEmpty(vWebServiceData.ClassName) Then
          'Instantiate instance of the class
          Dim vType As Type = Nothing
          Try
            vType = Type.GetType("CARE.Access." & vWebServiceData.ClassName & ",CDBNACCESS", True)
          Catch vEx As Exception
            ShowError("Cannot Find Class " & vWebServiceData.ClassName)
          End Try
          If vType IsNot Nothing Then
            Dim vArgs(0) As Object
            vArgs(0) = vEnv
            Dim vTable As CARERecord = CType(vType.InvokeMember(Nothing, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.CreateInstance Or BindingFlags.Instance, Nothing, Nothing, vArgs), CARERecord)
            Dim vAttrs As CARE.Collections.CollectionList(Of MaintenanceAttribute) = vTable.MaintenanceAttributes
            For Each vAttr As MaintenanceAttribute In vAttrs
              Dim vName As String = CARE.Utilities.Common.ProperName(vAttr.AttributeName)

              Dim vDataType As String = vAttr.Type
              Dim vMandatory As String = CARE.Utilities.Common.BooleanString(vAttr.NullsInvalid)
              If vWebServiceData.CallType = WebServiceData.CallTypes.UpdateRecord AndAlso IsTableID(vWebServiceData.ClassName, vName) Then
                'For an update the id is mandatory
              Else
                If Not IsTableIDForReturn(vWebServiceData.ClassName, vName) Then vMandatory = "N" 'nothing else should be mandatory
              End If
              If vWebServiceData.CallType = WebServiceData.CallTypes.AddRecord OrElse vWebServiceData.CallType = WebServiceData.CallTypes.UpdateRecord OrElse vAttr.PrimaryKey Then
                Select Case vName
                  Case "AmendedBy", "AmendedOn", "CreatedBy", "CreatedOn"
                    'Don't document
                  Case Else
                    If vWebServiceData.CallType = WebServiceData.CallTypes.AddRecord AndAlso IsTableID(vWebServiceData.ClassName, vName) Then
                      'Don't add primary key for insert
                    Else
                      .WriteLine(String.Format("      <parameter DataType=""{0}"" Mandatory=""{1}"">{2}</parameter>", vDataType, vMandatory, vName))
                    End If
                End Select
              End If
            Next
            If vWebServiceData.CallType = WebServiceData.CallTypes.AddRecord OrElse vWebServiceData.CallType = WebServiceData.CallTypes.UpdateRecord Then
              .WriteLine("    </parameters>")
              vWriteEndParameters = False
              .WriteLine("    <returns>")
              For Each vAttr As MaintenanceAttribute In vAttrs
                Dim vName As String = CARE.Utilities.Common.ProperName(vAttr.AttributeName)
                Dim vDataType As String = vAttr.Type
                If IsTableIDForReturn(vWebServiceData.ClassName, vName) Then
                  .WriteLine(String.Format("      <parameter DataType=""{0}"">{1}</parameter>", vDataType, vName))
                End If
              Next
              .WriteLine("    </returns>")
            ElseIf vWebServiceData.CallType = WebServiceData.CallTypes.DeleteRecord Then
              .WriteLine("    </parameters>")
              vWriteEndParameters = False
              .WriteLine("    <returns>")
              .WriteLine("    <parameter DataType=""C"" Notes=""Value of 'OK' if successful"">Result</parameter>")
              .WriteLine("    </returns>")
            End If
          End If
        Else
          vLogWriter.WriteLine("No class for " & vWebServiceData.Name)
        End If
        If vWriteEndParameters Then .WriteLine("    </parameters>")
        .WriteLine("  </method>")
      Next
      .WriteLine("</methods>")
      vWriter.Close()
      vLogWriter.Close()
    End With

  End Sub

  Private Function IsTableID(pClassName As String, pName As String) As Boolean
    If pName = pClassName & "Id" Then
      Return True
    Else
      Return False
    End If
  End Function

  Private Function IsTableIDForReturn(pClassName As String, pName As String) As Boolean
    If pName = pClassName & "Id" Then
      Return True
    ElseIf pClassName = "ExamUnitLink" AndAlso (pName.EndsWith("Id1") OrElse pName.EndsWith("Id2")) Then
      Return True
    Else
      Return False
    End If
  End Function

  Private Function GetMethodNameDesc(pName As String) As String
    Dim vDesc As New StringBuilder
    vDesc.Append(pName.Substring(0, 1))
    For vIndex As Integer = 2 To pName.Length
      Dim vChar As String = pName.Substring(vIndex - 1, 1)
      If vChar.ToUpper = vChar Then vDesc.Append(" ")
      vDesc.Append(vChar)
    Next
    Return vDesc.ToString
  End Function
End Class
