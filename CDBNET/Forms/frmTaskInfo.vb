Public Class frmTaskInfo
  Private mvTaskType As CareServices.TaskJobTypes
  Private mvTaskName As String
  Private mvJobNumber As Integer
  Private WithEvents mvTimer As System.Windows.Forms.Timer
  Private Const DBUPGRADE_TIMER_INTERVAL As Integer = 2000        '2 second interval
  Private Const TIMER_INTERVAL As Integer = 5000                  '5 second interval
  Public Event RefreshStatus(ByVal Sender As Object, ByVal pJobNumber As String, ByVal Status As String)

  'This form could be used to display the processing status of a specific task
  'At present is only used for the database upgrade process

  Public Sub New(ByVal pTaskType As CareServices.TaskJobTypes)
    InitializeComponent()
    InitialiseControls(pTaskType)
  End Sub

  Private Sub InitialiseControls(ByVal pTaskType As CareServices.TaskJobTypes)
    SetControlTheme()
    Me.Text = ControlText.FrmTaskStatus
    MainHelper.SetMDIParent(Me)
    mvTaskType = pTaskType
    mvTaskName = FormHelper.GetTaskJobTypeName(mvTaskType)
    lblTask.Text = String.Format("The '{0}' task has been set to run asynchronously on the server", mvTaskName)
    mvTimer = New System.Windows.Forms.Timer
    If mvTaskType = CareNetServices.TaskJobTypes.tjtDatabaseUpgrade Then
      mvTimer.Interval = DBUPGRADE_TIMER_INTERVAL
    Else
      mvTimer.Interval = TIMER_INTERVAL
      lblUpgrade.Visible = False
    End If
    DoRefresh()
    StartTimer()
  End Sub

  Public Sub DoRefresh()
    Try
      Dim vList As New ParameterList(True)
      vList("SubmittedBy") = DataHelper.UserInfo.Logname
      vList("JobStatus") = "C,H"
      Dim vStatus As String = "Searching for Job Status"
      Dim vJobNoStatus As String = "Searching for Job Number"
      Dim vJobFound As Boolean
      Dim vDataSet As DataSet = DataHelper.GetJobScheduleData(vList)
      If vDataSet IsNot Nothing Then
        Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
        If vDataTable IsNot Nothing Then
          'Debug.Print("Got " & vDataTable.Rows.Count & " Rows")
          For Each vRow As DataRow In vDataTable.Rows
            If vRow("Description").ToString = mvTaskName Then
              'Here we have found a running task with the correct name - read the job number
              Dim vJobNumber As Integer = IntegerValue(vRow("JobNo").ToString)
              'Debug.Print("Job " & vJobNumber)
              'Find the highest job number in case a previous run crashed and was left running
              If mvJobNumber = 0 OrElse vJobNumber > mvJobNumber Then
                mvJobNumber = vJobNumber
              End If
              If vJobNumber = mvJobNumber Then
                vStatus = vRow("Information").ToString
                vJobNoStatus = mvJobNumber.ToString
                'Debug.Print("Status " & vStatus)
                vJobFound = True
              End If
            End If
          Next
        End If

        lblStatus.Text = vStatus
        lblJobNumber.Text = vJobNoStatus
        lblStatus.Refresh()


        If vJobFound = False And mvJobNumber > 0 Then
          'We found the job in the past but now cannot find it so must assume it has completed
          Me.Close()
        End If
      End If
      RaiseEvent RefreshStatus(Me, vJobNoStatus, vStatus)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub mvTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvTimer.Tick
    DoRefresh()
  End Sub

  Private Sub frmTaskInfo_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    StopTimer()
  End Sub
  ''' <summary>
  ''' This method will stop the timer.Public methods so that it can be called from outside classes (DataImport) 
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub StopTimer()
    mvTimer.Stop()
  End Sub
  ''' <summary>
  ''' Start timer. Public methods so that it can be called from outside classes (DataImport)
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub StartTimer()
    mvTimer.Start()
  End Sub
End Class