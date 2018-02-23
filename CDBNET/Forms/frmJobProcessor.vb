Public Class frmJobProcessor

#Region "Private Members"

  Private mvInterval As Long                'Number of minutes between checks
  Private mvMinutes As Long                 'Number of minutes since last check
  Private mvPeriod As PeriodOptionTypes     'Period type for history selection
  Private mvDataSet As DataSet
  Private Const MAX_ROWS As Integer = 5000

  Private Enum PeriodOptionTypes As Integer
    potWeek = 0
    potMonth
    potQuarter
    potCustom
  End Enum

#End Region

#Region "Constructor"

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

#End Region

#Region "Private Methods"

  Private Sub InitialiseControls()
    SetControlTheme()

    'Set tab captions
    tab.TabPages(0).Text = ControlText.TbpJobHistory
    tab.TabPages(1).Text = ControlText.TbpOptions
    tab.TabPages(2).Text = ControlText.TbpJobProcessors
    tab.SetItemSizes()

    'Date formats
    dtpFrom.CustomFormat = AppValues.DateFormat
    dtpTo.CustomFormat = AppValues.DateFormat

    'Set control based on user permissions
    Dim vEnabled As Boolean
    If DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualDatabaseAdministrator OrElse DataHelper.UserInfo.AccessLevel = UserInfo.UserAccessLevel.ualSupervisor Then vEnabled = True
    cmdDelete.Enabled = vEnabled
    cmdReSubmit.Enabled = vEnabled

    'Set MaxRows
    txtMaxRows.Text = MAX_ROWS.ToString

    AddHandler txtTimerInterval.KeyPress, AddressOf IntegerKeyPressHandler
    AddHandler txtMaxRows.KeyPress, AddressOf IntegerKeyPressHandler
  End Sub

  ''' <summary>
  ''' Get the selected job number from the job history grid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetJobNumber() As Double
    If dgrJobSchedule.RowCount > 0 Then
      Return DoubleValue(dgrJobSchedule.GetValue(dgrJobSchedule.CurrentRow, dgrJobSchedule.GetColumn("JobNo")))
    Else
      Return -1
    End If
  End Function

  ''' <summary>
  ''' Validate the controls on the Options Tab
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ValidateOptions() As Boolean
    Dim vValid As Boolean
    Dim vInterval As Integer
    Dim vErrorControl As Control = Nothing

    vInterval = IntegerValue(txtTimerInterval.Text)
    If vInterval < 1 Or vInterval > 999 Then
      vErrorControl = txtTimerInterval
    Else
      mvInterval = vInterval
      mvMinutes = 0
      SetOption("TimerInterval", txtTimerInterval.Text)
    End If

    If mvPeriod = PeriodOptionTypes.potCustom Then
      If dtpFrom.Checked = False AndAlso dtpTo.Checked = False Then vErrorControl = dtpFrom
      If dtpTo.Value < dtpFrom.Value Then vErrorControl = dtpFrom
    End If

    If Not vErrorControl Is Nothing Then
      Beep()
      tab.SelectedTab = tbpOptions
      vErrorControl.Focus()
    Else
      vValid = True
      tab.SelectedTab = tbpJobHistory
    End If

    Return vValid
  End Function

  ''' <summary>
  ''' Save user settings
  ''' </summary>
  ''' <param name="pOption">The key for the options</param>
  ''' <param name="pValue">The value</param>
  ''' <remarks></remarks>
  Private Sub SetOption(ByVal pOption As String, ByVal pValue As String)
    SaveSetting("CARE", "Job Processor", pOption, pValue)
  End Sub

  ''' <summary>
  ''' Fetch the job history and processor data based on the options
  ''' and populatet the grids
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub LoadJobs()
    Dim vFromDate As Date = Date.MinValue
    Dim vToDate As Date = Date.MinValue

    Select Case mvPeriod
      Case PeriodOptionTypes.potWeek                    'Last Week
        vFromDate = Today.AddDays(-7)
      Case PeriodOptionTypes.potMonth                   'Last Month
        vFromDate = Today.AddMonths(-1)
      Case PeriodOptionTypes.potQuarter                 'Last Quarter
        vFromDate = Today.AddMonths(-3)
      Case PeriodOptionTypes.potCustom                  'Custom
        If dtpFrom.Checked Then vFromDate = dtpFrom.Value
        If dtpTo.Checked Then vToDate = dtpTo.Value.AddDays(1)
    End Select

    'Fetch data only if criteria is specified
    If vFromDate > Date.MinValue OrElse vToDate > Date.MinValue Then
      'Job processor
      Dim vList As New ParameterList(True)
      Dim vJobProcessors As DataSet = DataHelper.GetJobProcessorsData(vList)
      dgrJobProcessors.Populate(vJobProcessors)
      SetJobColours(vJobProcessors)

      'Job history
      If vFromDate > Date.MinValue Then vList("FromDate") = vFromDate.ToString(AppValues.DateFormat)
      If vToDate > Date.MinValue Then vList("ToDate") = vToDate.ToString(AppValues.DateFormat)
      If txtMaxRows.Text.Length > 0 Then vList("MaxRows") = txtMaxRows.Text
      mvDataSet = DataHelper.GetJobScheduleData(vList)
      dgrJobSchedule.Populate(mvDataSet)
    End If
  End Sub

  ''' <summary>
  ''' Set the colour for the jobs that are not polling or inactive
  ''' </summary>
  ''' <param name="pDataSet"></param>
  ''' <remarks></remarks>
  Private Sub SetJobColours(ByVal pDataSet As DataSet)
    Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(pDataSet)
    If Not vDataTable Is Nothing Then
      For vIndex As Integer = 0 To vDataTable.Rows.Count - 1
        With vDataTable.Rows(vIndex)
          If .Item("Polling").ToString = "N" Then
            For vColumn As Integer = 0 To dgrJobProcessors.ColumnCount - 1
              dgrJobProcessors.SetBackgroundColour(vIndex, vColumn, Color.LightPink)
            Next
          ElseIf .Item("Active").ToString = "N" Then
            For vColumn As Integer = 0 To dgrJobProcessors.ColumnCount - 1
              dgrJobProcessors.SetBackgroundColour(vIndex, vColumn, Color.LightGray)
            Next
          End If
        End With
      Next
    End If
  End Sub

  Private Sub SelectOption()
    Select Case mvPeriod
      Case PeriodOptionTypes.potWeek
        optPrevWeek.Checked = True
      Case PeriodOptionTypes.potMonth
        optPrevMonth.Checked = True
      Case PeriodOptionTypes.potQuarter
        optPrevQuater.Checked = True
      Case PeriodOptionTypes.potCustom
        optCustomPeriod.Checked = True
    End Select
  End Sub

  ''' <summary>
  ''' Retrive a settings value 
  ''' </summary>
  ''' <param name="pOption">The key</param>
  ''' <param name="pDefault">default value</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetOption(ByVal pOption As String, Optional ByVal pDefault As String = "") As String
    Dim vReturn As String = GetSetting("CARE", "Job Processor", pOption, "")
    If vReturn = "" Then
      vReturn = GetSetting("Contacts Database", "Job Processor", pOption, pDefault)
    End If
    Return vReturn
  End Function

#End Region

#Region "Control Events"

  Private Sub frmJobProcessor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    tim.Enabled = False
  End Sub

  Private Sub frmJobProcessor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      'Now get the preferences settings
      txtTimerInterval.Text = GetOption("TimerInterval", "5")       '5 minutes
      mvInterval = IntegerValue(txtTimerInterval.Text)
      mvMinutes = 0

      Dim vPeriod As Integer = IntegerValue(GetOption("HistoryPeriod", "0"))
      If vPeriod > 3 Then vPeriod = 0
      mvPeriod = CType(vPeriod, PeriodOptionTypes)
      SelectOption()

      Dim vDate As String = GetOption("HistoryFrom", Today.ToString(AppValues.DateFormat))
      If IsDate(vDate) Then dtpFrom.Text = vDate

      vDate = GetOption("HistoryTo", Today.ToString(AppValues.DateFormat))
      If IsDate(vDate) Then dtpTo.Text = vDate

      LoadJobs()

      tim.Interval = 60000    '1 minute
      tim.Enabled = True
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      If ValidateOptions() Then
        Dim vJobNumber As Double = GetJobNumber()
        If vJobNumber > -1 Then
          If ShowQuestion(String.Format(QuestionMessages.QmDeleteJob, vJobNumber), MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            Dim vList As New ParameterList(True)
            vList("JobNumber") = vJobNumber.ToString
            DataHelper.DeleteJobSchedule(vList)
            dgrJobSchedule.DeleteRow(dgrJobSchedule.CurrentRow)
          End If
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdReSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReSubmit.Click
    Try
      If ValidateOptions() Then
        Dim vJobNumber As Double = GetJobNumber()
        If vJobNumber > -1 Then
          If ShowQuestion(String.Format(QuestionMessages.QmResubmitJob, vJobNumber), MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            Dim vList As New ParameterList(True)
            vList("JobNumber") = vJobNumber.ToString
            DataHelper.ResubmitJobSchedule(vList)
            LoadJobs()
          End If
        End If
      End If
      Exit Sub
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    Try
      If ValidateOptions() Then LoadJobs()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub opt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPrevWeek.CheckedChanged, optPrevMonth.CheckedChanged, optPrevQuater.CheckedChanged, optCustomPeriod.CheckedChanged
    Try
      If optPrevWeek.Checked Then
        mvPeriod = PeriodOptionTypes.potWeek
      ElseIf optPrevMonth.Checked Then
        mvPeriod = PeriodOptionTypes.potMonth
      ElseIf optPrevQuater.Checked Then
        mvPeriod = PeriodOptionTypes.potQuarter
      Else
        mvPeriod = PeriodOptionTypes.potCustom
      End If
      SetOption("HistoryPeriod", CInt(mvPeriod).ToString)
      'enable/disable From/To dates
      dtpFrom.Enabled = optCustomPeriod.Checked
      dtpTo.Enabled = optCustomPeriod.Checked
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged
    Try
      Dim vFrom As String = String.Empty
      If dtpFrom.Checked Then vFrom = dtpFrom.Value.ToString(AppValues.DateFormat)
      SetOption("HistoryFrom", vFrom)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub dtpTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged
    Try
      Dim vTo As String = String.Empty
      If dtpTo.Checked Then vTo = dtpTo.Value.ToString(AppValues.DateFormat)
      SetOption("HistoryTo", vTo)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub tim_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tim.Tick
    Try
      mvMinutes = mvMinutes + 1
      If mvMinutes >= mvInterval Then
        LoadJobs()
        mvMinutes = 0
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub txtTimerInterval_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTimerInterval.TextChanged
    Try
      If Not mvInterval.ToString = txtTimerInterval.Text Then
        mvInterval = IntegerValue(txtTimerInterval.Text)
        mvMinutes = 0
        SetOption("TimerInterval", txtTimerInterval.Text)
        LoadJobs()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub
#End Region

End Class