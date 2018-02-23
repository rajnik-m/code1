Public Class frmSelectItems

  Public Sub New(ByVal pDataSet As DataSet)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pDataSet)
  End Sub

  Private Sub InitialiseControls(ByVal pDataSet As DataSet)
    'At present only one use which is for email attachment selection
    'Need to build a list of documents attached to this document which can be saved as attachments
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    dgr.Populate(pDataSet)
    If dgr.RowCount > 0 Then
      dgr.SetCellsEditable()
      dgr.SetCellsReadOnly()
      dgr.SetCheckBoxColumn("Select")
    End If
    Me.Text = "Select related documents to include as attachments"
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Me.Close()
  End Sub

  Private Sub cmdSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectAll.Click
    Dim vCol As Integer = dgr.GetColumn("Select")
    For vRow As Integer = 0 To dgr.RowCount - 1
      dgr.SetValue(vRow, vCol, "True")
    Next
  End Sub

  Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
    Dim vCol As Integer = dgr.GetColumn("Select")
    For vRow As Integer = 0 To dgr.RowCount - 1
      dgr.SetValue(vRow, vCol, "")
    Next
  End Sub

  Private Sub frmSelectItems_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If dgr.RowCount = 0 Then
      Me.DialogResult = Windows.Forms.DialogResult.OK
      Me.Close()
    End If
  End Sub
End Class