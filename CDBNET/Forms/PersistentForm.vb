Public Class PersistentForm
  Inherits System.Windows.Forms.Form

  Private mvSettingsName As String
  Private mvSettingsBounds As String
  Private mvPersistSize As Boolean

  Friend Property SettingsName() As String
    Get
      Return mvSettingsName
    End Get
    Set(ByVal value As String)
      mvSettingsName = value
      mvSettingsBounds = mvSettingsName & "_Bounds"
      mvPersistSize = mvSettingsName.Length > 0
    End Set
  End Property

  Private Sub PersistentForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    If mvPersistSize AndAlso Me.WindowState = FormWindowState.Normal Then
      AppValues.SetWindowSize(mvSettingsBounds, Me.Left, Me.Top, Me.Width, Me.Height)
    End If
  End Sub

  Private Sub PersistentForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If mvPersistSize Then
      Dim vWindowSize As WindowSize = AppValues.GetWindowSize(mvSettingsBounds)
      If vWindowSize IsNot Nothing Then
        Me.SetBounds(vWindowSize.X, vWindowSize.Y, vWindowSize.Width, vWindowSize.Height)
      End If
    End If
  End Sub

  Private Sub InitializeComponent()
    Me.SuspendLayout()
    '
    'PersistentForm
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.ClientSize = New System.Drawing.Size(292, 260)
    Me.Name = "PersistentForm"
    Me.ResumeLayout(False)

  End Sub
End Class
