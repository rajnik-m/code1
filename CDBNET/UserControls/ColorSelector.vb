Public Class ColorSelector

  Private mvColorDialog As ColorDialog

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
  End Sub

  <System.ComponentModel.Browsable(True)> _
  <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
  Public Overrides Property Text() As String
    Get
      Return lbl.Text
    End Get
    Set(ByVal pValue As String)
      lbl.Text = pValue
    End Set
  End Property

  <System.ComponentModel.Browsable(True)> _
  <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
  Public Property Color() As Color
    Get
      Return pnl.BackColor
    End Get
    Set(ByVal pValue As Color)
      pnl.BackColor = pValue
    End Set
  End Property

  Public Property ColorDialog() As ColorDialog
    Get
      Return mvColorDialog
    End Get
    Set(ByVal value As ColorDialog)
      mvColorDialog = value
    End Set
  End Property

  Private Sub cmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmd.Click
    If mvColorDialog Is Nothing Then mvColorDialog = New ColorDialog
    With mvColorDialog
      .Color = Me.Color
      .AllowFullOpen = True
      .SolidColorOnly = True
      If .ShowDialog = DialogResult.OK Then
        Me.Color = .Color
      End If
    End With
  End Sub

  Private Sub ColorSelector_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FontChanged
    PositionControls()
  End Sub

  Private Sub ColorSelector_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    PositionControls()
  End Sub

  Private Sub PositionControls()
    cmd.Left = Me.Width - (cmd.Width + 2)
    pnl.Left = cmd.Left - (pnl.Width + 4)
    pnl.Top = cmd.Top
  End Sub

End Class
