Public Class frmAboutBox

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
    pnl.BackgroundExtenderType = BackGroundExtender.BackgroundExtenderTypes.betFillPanel
    SetControlTheme()
    cmdOK.Text = ControlText.CmdOK
    lblDescription.Text = My.Application.Info.Title
    lblClient.Text = String.Format(ControlText.LblAboutClient, My.Application.Info.Version.ToString)
    lblServer.Text = String.Format(ControlText.LblAboutServer, DataHelper.GetVersion)
    lblWarning.Text = GetInformationMessage(ControlText.LblCopyrightMsg)
    Me.Text = GetInformationMessage(ControlText.FrmAbout)
  End Sub

  Private Sub llb_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
    Try
      Dim vLink As LinkLabel = DirectCast(sender, LinkLabel)
      Dim vBrowser As LocalBrowser = New LocalBrowser(vLink.Text)
      vLink.LinkVisited = True
    Catch ex As Exception
      'Could not access link
    End Try
  End Sub

  Public Overrides Sub SetControlTheme()
    MyBase.SetControlTheme()
    DisplayTheme.ThemeButton(cmdOK)
  End Sub

End Class