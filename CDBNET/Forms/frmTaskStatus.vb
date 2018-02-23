Public Class frmTaskStatus

  Public Sub New()
    InitializeComponent()
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()
    Me.Text = ControlText.FrmTaskStatus
    MainHelper.SetMDIParent(Me)
    dgr.AutoSetHeight = False
    DoRefresh()
  End Sub

  Public Sub DoRefresh()
    Dim vList As New ParameterList(True)
    If Me.InvokeRequired Then
      Me.Invoke(New MethodInvoker(AddressOf DoRefresh))
    Else
      vList("SubmittedBy") = DataHelper.UserInfo.Logname
      vList("JobStatus") = "C,H"
      Dim vDataSet As DataSet = DataHelper.GetJobScheduleData(vList)
      UpdateTaskStatus(vDataSet)
    End If
  End Sub

  Private Sub UpdateTaskStatus(ByVal pDataSet As DataSet)
    If pDataSet IsNot Nothing Then
      dgr.Populate(pDataSet)
    End If
  End Sub

  Private Sub cmdRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    Dim vCursor As New BusyCursor
    Try
      DoRefresh()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

  Private Sub frmTaskStatus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    FormHelper.TaskStatusForm = Nothing
  End Sub
End Class