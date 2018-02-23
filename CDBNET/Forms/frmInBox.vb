Public Class frmInbox
  Inherits PersistentForm

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls()
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  Friend WithEvents cmdShow As System.Windows.Forms.Button
  Friend WithEvents cmdRefresh As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInbox))
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdShow = New System.Windows.Forms.Button
    Me.cmdRefresh = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(767, 338)
    Me.dgr.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdShow)
    Me.bpl.Controls.Add(Me.cmdRefresh)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 338)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(767, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdShow
    '
    Me.cmdShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdShow.Location = New System.Drawing.Point(169, 6)
    Me.cmdShow.Name = "cmdShow"
    Me.cmdShow.Size = New System.Drawing.Size(96, 27)
    Me.cmdShow.TabIndex = 0
    Me.cmdShow.Text = "&Show"
    '
    'cmdRefresh
    '
    Me.cmdRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdRefresh.Location = New System.Drawing.Point(280, 6)
    Me.cmdRefresh.Name = "cmdRefresh"
    Me.cmdRefresh.Size = New System.Drawing.Size(96, 27)
    Me.cmdRefresh.TabIndex = 3
    Me.cmdRefresh.Text = "&Refresh"
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.Location = New System.Drawing.Point(391, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(96, 27)
    Me.cmdDelete.TabIndex = 1
    Me.cmdDelete.Text = "&Delete"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(502, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Close"
    '
    'frmInbox
    '
    Me.AcceptButton = Me.cmdShow
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(767, 377)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmInbox"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub InitialiseControls()
    SetControlTheme()
    Me.cmdShow.Text = ControlText.CmdShow
    Me.cmdRefresh.Text = ControlText.CmdRefresh
    Me.cmdDelete.Text = ControlText.CmdDelete
    Me.cmdCancel.Text = ControlText.CmdClose

    Me.Text = ControlText.FrmInBox
    SettingsName = "InBox"
    MainHelper.SetMDIParent(Me)
    ReloadEMailData()
  End Sub

  Private Sub ReloadEMailData()
    Dim vCursor As New BusyCursor
    Try
      dgr.Populate(EMailApplication.EmailInterface.GetInBoxData)
      dgr.SetBoldRowsFromColumn("Read")
      dgr.SetIconColumn("Attachments")
      Me.Text = GetInformationMessage(InformationMessages.frmEMailCaption, dgr.RowCount.ToString)
      DoSetStatusMessage("")
      Dim vGotRecords As Boolean = dgr.RowCount > 0
      cmdShow.Enabled = vGotRecords
      cmdDelete.Enabled = vGotRecords
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShow.Click
    ShowEMail()
  End Sub

  Private Sub dgr_RowDoubleClicked(ByVal pSender As Object, ByVal pRow As Integer) Handles dgr.RowDoubleClicked
    ShowEMail()
  End Sub

  Private Sub ShowEMail()
    Try
      Dim vID As String = dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("ID"))
      If vID.Length > 0 Then
        'Need to mark the email as Read
        dgr.SetValue(dgr.CurrentRow, "Read", True.ToString)
        dgr.SetBoldRowsFromColumn("Read")
        Dim vForm As New frmShowEmail(vID)
        vForm.ShowDialog(Me)
        If vForm.EMailAction = EmailInterface.EMailActions.emaDelete Then
          dgr.DeleteRow(dgr.CurrentRow)
          Me.Text = GetInformationMessage(InformationMessages.frmEMailCaption, dgr.RowCount.ToString)
        Else

        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public Sub EMailDeleted(ByVal pID As String)
    If pID = dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("ID")) Then
      dgr.DeleteRow(dgr.CurrentRow)
      Me.Text = GetInformationMessage(InformationMessages.frmEMailCaption, dgr.RowCount.ToString)
    End If
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Dim vID As String = dgr.GetValue(dgr.CurrentRow, dgr.GetColumn("ID"))
    If vID.Length > 0 Then
      If EMailApplication.EmailInterface.ProcessAction(vID, EmailInterface.EMailActions.emaDelete) Then
        EMailDeleted(vID)
      End If
    End If
  End Sub

  Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    ReloadEMailData()
  End Sub
End Class
