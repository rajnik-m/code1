<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAboutBox
  Inherits ThemedForm

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing AndAlso components IsNot Nothing Then
      components.Dispose()
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAboutBox))
    Me.pic = New System.Windows.Forms.PictureBox()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.pnl = New CDBNETCL.PanelEx()
    Me.lblWarning = New System.Windows.Forms.Label()
    Me.lblServer = New System.Windows.Forms.Label()
    Me.lblClient = New System.Windows.Forms.Label()
    Me.lblDescription = New System.Windows.Forms.Label()
    CType(Me.pic, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnl.SuspendLayout()
    Me.SuspendLayout()
    '
    'pic
    '
    Me.pic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
    Me.pic.Dock = System.Windows.Forms.DockStyle.Top
    Me.pic.Image = Global.CDBNET.My.Resources.Resources.advanced_banner
    Me.pic.Location = New System.Drawing.Point(0, 0)
    Me.pic.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pic.Name = "pic"
    Me.pic.Size = New System.Drawing.Size(495, 58)
    Me.pic.TabIndex = 3
    Me.pic.TabStop = False
    '
    'cmdOK
    '
    Me.cmdOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(200, 235)
    Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    Me.cmdOK.UseVisualStyleBackColor = True
    '
    'pnl
    '
    Me.pnl.BackColor = System.Drawing.Color.Transparent
    Me.pnl.Controls.Add(Me.lblWarning)
    Me.pnl.Controls.Add(Me.lblServer)
    Me.pnl.Controls.Add(Me.lblClient)
    Me.pnl.Controls.Add(Me.lblDescription)
    Me.pnl.Controls.Add(Me.cmdOK)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 58)
    Me.pnl.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(495, 274)
    Me.pnl.TabIndex = 4
    '
    'lblWarning
    '
    Me.lblWarning.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblWarning.Location = New System.Drawing.Point(0, 106)
    Me.lblWarning.Name = "lblWarning"
    Me.lblWarning.Padding = New System.Windows.Forms.Padding(50, 0, 50, 0)
    Me.lblWarning.Size = New System.Drawing.Size(495, 117)
    Me.lblWarning.TabIndex = 4
    Me.lblWarning.Text = resources.GetString("lblWarning.Text")
    Me.lblWarning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    '
    'lblServer
    '
    Me.lblServer.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblServer.Location = New System.Drawing.Point(0, 80)
    Me.lblServer.Name = "lblServer"
    Me.lblServer.Size = New System.Drawing.Size(495, 26)
    Me.lblServer.TabIndex = 5
    Me.lblServer.Text = "Server Version"
    Me.lblServer.TextAlign = System.Drawing.ContentAlignment.TopCenter
    '
    'lblClient
    '
    Me.lblClient.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblClient.Location = New System.Drawing.Point(0, 52)
    Me.lblClient.Name = "lblClient"
    Me.lblClient.Size = New System.Drawing.Size(495, 28)
    Me.lblClient.TabIndex = 2
    Me.lblClient.Text = "Client Version"
    Me.lblClient.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    '
    'lblDescription
    '
    Me.lblDescription.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblDescription.Location = New System.Drawing.Point(0, 0)
    Me.lblDescription.Name = "lblDescription"
    Me.lblDescription.Size = New System.Drawing.Size(495, 52)
    Me.lblDescription.TabIndex = 1
    Me.lblDescription.Text = "Description"
    Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    '
    'frmAboutBox
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.ClientSize = New System.Drawing.Size(495, 332)
    Me.ControlBox = False
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.pic)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmAboutBox"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "About"
    CType(Me.pic, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents pnl As CDBNETCL.PanelEx
  Friend WithEvents pic As System.Windows.Forms.PictureBox
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents lblDescription As System.Windows.Forms.Label
  Friend WithEvents lblClient As System.Windows.Forms.Label
  Friend WithEvents lblWarning As System.Windows.Forms.Label
  Friend WithEvents lblServer As System.Windows.Forms.Label
End Class
