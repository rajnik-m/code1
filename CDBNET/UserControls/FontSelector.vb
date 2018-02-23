Public Class FontSelector

  Private mvFont As Font

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    txtFont.Text = Me.Font.ToString
  End Sub

  <System.ComponentModel.Browsable(True)> _
  <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
  Public Property SelectedFont() As Font
    Get
      Return mvFont
    End Get
    Set(ByVal pValue As Font)
      mvFont = pValue
      SetFontName()
    End Set
  End Property

  Private Sub SetFontName()
    If mvFont Is Nothing Then
      txtFont.Text = "Default"
      txtFont.Font = New Font("Microsoft sans serif", 8, FontStyle.Regular)
      chk.Checked = False
      chk.Enabled = False
    Else
      txtFont.Text = String.Format("{0}pt. {1}", CInt(mvFont.Size), mvFont.Name)
      txtFont.Font = New Font(Me.Font.FontFamily, Me.Font.Size, mvFont.Style)
      chk.Enabled = True
      chk.Checked = True
    End If
  End Sub

  Private Sub chk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk.Click
    If chk.Checked = False AndAlso mvFont IsNot Nothing Then
      SelectedFont = Nothing
    End If
  End Sub

  Private Sub cmdFont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFont.Click
    If mvFont Is Nothing Then
      fnt.Font = Me.Font
    Else
      fnt.Font = mvFont
    End If
    fnt.AllowVerticalFonts = False
    fnt.AllowScriptChange = False
    If fnt.ShowDialog = DialogResult.OK Then
      mvFont = fnt.Font
      SetFontName()
    End If
  End Sub

  Private Sub FontSelector_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FontChanged
    chk.Left = Me.Width - (AppValues.CheckBoxPixelsX + 4)
  End Sub

  Private Sub FontSelector_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    chk.Left = Me.Width - (AppValues.CheckBoxPixelsX + 4)
  End Sub
End Class
