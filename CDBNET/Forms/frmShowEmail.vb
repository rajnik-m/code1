
Public Class frmShowEmail
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New(ByVal pID As String)
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pID)
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
  Friend WithEvents epl As CDBNETCL.EditPanel
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdReplyAll As System.Windows.Forms.Button
  Friend WithEvents cmdForward As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
  Friend WithEvents cmdReply As System.Windows.Forms.Button
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmShowEmail))
    Me.epl = New CDBNETCL.EditPanel
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdReply = New System.Windows.Forms.Button
    Me.cmdReplyAll = New System.Windows.Forms.Button
    Me.cmdForward = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    Me.cmdSave = New System.Windows.Forms.Button
    Me.cmdClose = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'epl
    '
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(688, 401)
    Me.epl.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdReply)
    Me.bpl.Controls.Add(Me.cmdReplyAll)
    Me.bpl.Controls.Add(Me.cmdForward)
    Me.bpl.Controls.Add(Me.cmdDelete)
    Me.bpl.Controls.Add(Me.cmdSave)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.Location = New System.Drawing.Point(0, 401)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(688, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdReply
    '
    Me.cmdReply.Location = New System.Drawing.Point(24, 6)
    Me.cmdReply.Name = "cmdReply"
    Me.cmdReply.Size = New System.Drawing.Size(94, 27)
    Me.cmdReply.TabIndex = 0
    Me.cmdReply.Text = "Reply"
    '
    'cmdReplyAll
    '
    Me.cmdReplyAll.Location = New System.Drawing.Point(133, 6)
    Me.cmdReplyAll.Name = "cmdReplyAll"
    Me.cmdReplyAll.Size = New System.Drawing.Size(94, 27)
    Me.cmdReplyAll.TabIndex = 1
    Me.cmdReplyAll.Text = "Reply All"
    '
    'cmdForward
    '
    Me.cmdForward.Location = New System.Drawing.Point(242, 6)
    Me.cmdForward.Name = "cmdForward"
    Me.cmdForward.Size = New System.Drawing.Size(94, 27)
    Me.cmdForward.TabIndex = 2
    Me.cmdForward.Text = "Forward"
    '
    'cmdDelete
    '
    Me.cmdDelete.Location = New System.Drawing.Point(351, 6)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(94, 27)
    Me.cmdDelete.TabIndex = 3
    Me.cmdDelete.Text = "Delete"
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(460, 6)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(94, 27)
    Me.cmdSave.TabIndex = 4
    Me.cmdSave.Text = "Save"
    '
    'cmdClose
    '
    Me.cmdClose.Location = New System.Drawing.Point(569, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(94, 27)
    Me.cmdClose.TabIndex = 5
    Me.cmdClose.Text = "Close"
    '
    'frmShowEmail
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
    Me.ClientSize = New System.Drawing.Size(688, 440)
    Me.Controls.Add(Me.epl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmShowEmail"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private mvEmail As EMailMessage
  Private mvAction As EmailInterface.EMailActions

  Private Sub InitialiseControls(ByVal pID As String)
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    SetControlColors(Me)
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optEMailDisplay))
    Me.ClientSize = New Size(Me.ClientSize.Width, epl.RequiredHeight + bpl.Height)
    epl.SetAnchor("Body", AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom)
    epl.SetAnchor("Attachments", AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom)
    Dim vList As New ParameterList
    mvEmail = EMailApplication.EmailInterface.EmailMessageByID(pID)
    EMailApplication.EmailInterface.MarkRead(mvEmail)
    With mvEmail
      Me.Text = .Subject
      If .OrigDisplayName <> .OrigAddress Then
        vList("From") = .OrigDisplayName & " [" & .OrigAddress & "]"
      Else
        vList("From") = .OrigAddress
      End If
      vList("To") = .ToList
      vList("Subject") = .Subject
      vList("CC") = .CCList
      vList("Received") = .DateReceived
      vList("Body") = .NoteText
      vList("Attachments") = .AttachmentNameList
    End With
    epl.Populate(vList)
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    mvAction = EmailInterface.EMailActions.emaNone
    Me.Close()
  End Sub

  Private Sub cmdReply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReply.Click, cmdReplyAll.Click, cmdDelete.Click, cmdForward.Click, cmdSave.Click
    Dim vCursor As New BusyCursor
    Try
      If sender Is cmdReply Then
        mvAction = EmailInterface.EMailActions.emaReply
      ElseIf sender Is cmdReplyAll Then
        mvAction = EmailInterface.EMailActions.emaReplyAll
      ElseIf sender Is cmdForward Then
        mvAction = EmailInterface.EMailActions.emaForward
      ElseIf sender Is cmdDelete Then
        mvAction = EmailInterface.EMailActions.emaDelete
      ElseIf sender Is cmdSave Then
        mvAction = EmailInterface.EMailActions.emaSave
      End If
      Me.Close()
      If mvAction = EmailInterface.EMailActions.emaSave Then
        mvEmail.DeleteAfterSave = ShowQuestion(QuestionMessages.qmRetainEMailInInbox, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No
        Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctEMailDocument, 0, Nothing)
        vForm.EMailMessage = mvEmail
        vForm.Show()
      Else
        EMailApplication.EmailInterface.ProcessAction(mvEmail, mvAction)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public ReadOnly Property EMailAction() As EmailInterface.EMailActions
    Get
      Return mvAction
    End Get
  End Property

  Private Sub epl_AttachmentNavigate(ByVal pSender As Object, ByVal pIndex As Integer) Handles epl.AttachmentNavigate
    EMailApplication.EmailInterface.ShowAttachment(mvEmail.ID, pIndex)
  End Sub

  Private Sub frmShowEmail_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    If Not DocumentApplication Is Nothing Then DocumentApplication.ProcessAppActive()
  End Sub
End Class
