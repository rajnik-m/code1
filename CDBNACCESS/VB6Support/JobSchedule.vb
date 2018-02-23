Namespace Access

  Partial Public Class JobSchedule

    Private mvRecordsProcessed As Integer
    Private mvReportRecordsProcessed As Integer
    Private mvRecordType As String = ""
    Private mvReportName As String = "" 'Currently running report
    Private mvLastEventTime As Date
    Private mvLastUpdateTime As Date
    Private mvInfoMessage As String
    Private mvStartTime As Date
    Private mvClientCursor As Boolean

    Public Event StatusChange(ByVal pJobStatus As JobStatuses)

    Public Property Parameters() As CDBParameters
      Get
        If mvParameters Is Nothing Then
          If CommandLine.Length > 0 Then
            'set up the params from the command line
            mvParameters = GetCDBParameterList(GetParameterList(CommandLine).XMLParameterString())
          Else
            mvParameters = New CDBParameters
          End If
        End If
        Return mvParameters
      End Get
      Set(ByVal value As CDBParameters)
        mvParameters = value
      End Set
    End Property

    Public Property InfoMessage() As String
      Get
        InfoMessage = mvInfoMessage
      End Get
      Set(ByVal Value As String)
        mvInfoMessage = Value
        If DateDiff(Microsoft.VisualBasic.DateInterval.Second, mvLastUpdateTime, Now) > 1 Then 'Update record every second
          UpdateStatus(mvInfoMessage)
        End If
      End Set
    End Property

    Private Sub UpdateStatus(ByRef pMsg As String)
      If mvExisting Then
        mvClassFields(JobScheduleFields.ErrorStatus).Value = pMsg
        If mvClientCursor = False Then
          Save() 'Update job schedule record with records processed status
          mvLastUpdateTime = Now
        End If
      End If
    End Sub

    Public Property ReportRecordsProcessed() As Integer
      Get
        Return mvReportRecordsProcessed
      End Get
      Set(ByVal Value As Integer)
        mvReportRecordsProcessed = Value
        If DateDiff(Microsoft.VisualBasic.DateInterval.Second, mvLastEventTime, Now) > 5 Then  'Update record every 5 seconds
          RaiseEvent StatusChange(JobStatuses.jbsReportProcessed)
          mvLastEventTime = Now
        End If
        If DateDiff(Microsoft.VisualBasic.DateInterval.Second, mvLastUpdateTime, Now) > 5 Then 'Update record every 5 seconds
          UpdateStatus("Report '" & ReportName & "' Processed " & ReportRecordsProcessed & " Rows")
        End If
      End Set
    End Property

    Public ReadOnly Property ReportName() As String
      Get
        Return mvReportName
      End Get
    End Property

    Public Sub SetReportStatus(ByVal pStatus As JobStatuses, ByVal pReportName As String)
      mvReportName = pReportName
      RaiseEvent StatusChange(pStatus)
      If DateDiff(Microsoft.VisualBasic.DateInterval.Second, mvLastUpdateTime, Now) > 5 Then 'Update record every 5 seconds
        Dim vMessage As String
        If pStatus = JobStatuses.jbsReportStarted Then
          vMessage = "Started Report '" & pReportName & "'"
          UpdateStatus(vMessage)
        ElseIf pStatus = JobStatuses.jbsReportCompleted Then
          vMessage = "Completed Report '" & pReportName & "'"
          UpdateStatus(vMessage)
        End If
      End If
    End Sub

    Public Property RecordsProcessed() As Integer
      Get
        Return mvRecordsProcessed
      End Get
      Set(ByVal Value As Integer)
        mvRecordsProcessed = Value
        If DateDiff(Microsoft.VisualBasic.DateInterval.Second, mvLastEventTime, Now) > 1 Then 'Update status every second
          RaiseEvent StatusChange(JobStatuses.jbsProcessed)
          mvLastEventTime = Now
          If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, mvLastUpdateTime, Now) > 1 Then 'Update record every minute
            UpdateStatus("Processed " & mvRecordsProcessed & " " & mvRecordType)
          End If
        End If
      End Set
    End Property

    Public Property RecordType() As String
      Get
        Return mvRecordType
      End Get
      Set(ByVal value As String)
        mvRecordType = value
      End Set
    End Property

    Public Sub Resubmit()
      If mvExisting Then
        mvClassFields(JobScheduleFields.JobStatus).Value = JobStatusCode(JobStatuses.jbsWaiting) 'Waiting
        mvClassFields(JobScheduleFields.EndDate).Value = ""                          'Clear End date
        mvClassFields(JobScheduleFields.RunDate).Value = ""                          'Clear Run date
        mvClassFields(JobScheduleFields.ErrorStatus).Value = ""                      'Clear Error Status
        Save()                                                          'Update
      End If
    End Sub

    Private Function GetParameterList(ByVal pParamList As String) As Collections.ParameterList
      'Build Parameter array from list, ignoring spaces in quoted strings
      Dim vParameterList As New Collections.ParameterList
      Dim vIndex As Integer
      Dim vFound As Boolean
      Dim vParameters As String() = pParamList.Split(" "c)
      'Parameters are MODULE INISECTION CLIENT LOGNAME PASSWORD [PARAMETERS]
      Do
        vFound = False
        For vIndex = 0 To vParameters.Length - 1
          If UnevenQuotes(vParameters(vIndex), """") Or UnevenQuotes(vParameters(vIndex), "'") Then
            If (vIndex + 1) < vParameters.Length Then
              vParameters(vIndex) = vParameters(vIndex) & " " & vParameters(vIndex + 1)
              While vIndex + 1 < vParameters.Length - 1
                vIndex += 1
                vParameters(vIndex) = vParameters(vIndex + 1)
              End While
              ReDim Preserve vParameters(vIndex)
              vFound = True
              Exit For
            End If
          End If
        Next
      Loop While vFound
      ' Now we have the array - get it into the parameterlist
      Dim vPos As Integer
      For vCount As Integer = 0 To vParameters.Length - 1
        Select Case vCount
          Case Is < 7
            ' do nothing
          Case Else
            Dim vParam As String = vParameters(vCount)
            vPos = vParam.IndexOf("=")
            If vPos >= 0 Then
              vParameterList.Add(vParam.Substring(0, vPos), vParam.Substring(vPos + 1, vParam.Length - (vPos + 1)).Trim(""""c))
            Else
              vParameterList.Add(vParam.Trim(""""c), "")
            End If
        End Select
      Next
      Return vParameterList
    End Function

    Private Function UnevenQuotes(ByVal pString As String, ByVal pQuoteString As String) As Boolean
      Dim vCount As Integer
      Dim vPos As Integer
      Dim vStartPos As Integer

      vStartPos = 0
      Do
        vPos = pString.IndexOf(pQuoteString, vStartPos)
        If vPos >= 0 Then
          vStartPos = vPos + 1
          vCount += 1
        End If
      Loop While vPos >= 0
      If vCount Mod 2 = 1 Then UnevenQuotes = True
    End Function

    Private Function GetCDBParameterList(ByVal pXMLString As String) As CDBParameters
      Dim vParams As New CDBParameters
      Dim vDoc As New Xml.XmlDocument

      If Len(pXMLString) > 0 Then
        vDoc.LoadXml(pXMLString)
        Dim vXMLRoot As Xml.XmlElement = vDoc.DocumentElement 'Parameters
        If vXMLRoot.Name = "Parameters" Then
          For Each vNode As Xml.XmlNode In vXMLRoot.ChildNodes
            Select Case vNode.Name
              Case "TraderAnalysisLine", "InvoiceLine", "PPDLine", "OPSLine", "MemberLine", "WebPageItemControl", "POSLine", "PPALine", "PISLine", "NonFinActivity", "NonFinSuppression", "FdePageItemControl", "FinderControl", "MaintenanceControl", "IncentiveLine", "SelectedInvoice", "EventBookingLine", "ExamBookingLine"
                'Do Nothing
              Case Else
                If vNode.HasChildNodes Then
                  vParams.Add((vNode.Name), CDBField.FieldTypes.cftCharacter, vNode.ChildNodes(0).InnerText.TrimEnd)
                Else
                  vParams.Add((vNode.Name), CDBField.FieldTypes.cftCharacter)
                End If
            End Select
          Next vNode
        Else
          'Error?
        End If
      End If
      Return vParams
    End Function

  End Class
End Namespace