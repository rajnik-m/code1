Public Class frmBrowser

  Public Sub New(ByVal pURL As String, ByVal vShowToolbar As Boolean, Optional ByVal pHelp As Boolean = False)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pURL, vShowToolbar, pHelp)
  End Sub

  Private Sub InitialiseControls(ByVal pURL As String, ByVal pShowToolbar As Boolean, Optional ByVal pHelp As Boolean = False)
    SetControlTheme()
    Me.MdiParent = MDIForm
    If pHelp Then
      Me.Text = "Care Help"
    Else
      Me.Text = pURL
    End If
    web.Init(pURL, pShowToolbar)
  End Sub

  'Updates the title bar with the current document title.
  Private Sub WebDocumentTitleChanged(ByVal sender As Object, ByVal pWebDocumentTitle As String) Handles web.WebDocumentTitleChanged
    Me.Text = pWebDocumentTitle
  End Sub

  'Updates the status bar with the current browser status text.
  Private Sub WebStatusTextChanged(ByVal sender As Object, ByVal pWebStatusText As String) Handles web.WebStatusTextChanged
    tsl.Text = pWebStatusText
  End Sub

End Class