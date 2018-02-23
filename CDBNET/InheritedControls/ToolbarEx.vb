Public Class ToolbarEx
  Inherits System.Windows.Forms.ToolBar

  Private mvBackgroundExtender As BackGroundExtender
  Private mvButtonExtender As BackGroundExtender
  Private mvMouseDown As Boolean
  Private mvLastButton As ToolBarButton
  Private mvLastMovedOver As Integer

  Public Sub New()
    MyBase.New()
    InitialiseControls()
  End Sub

  Public Sub InitialiseControls()
    SetStyle(System.Windows.Forms.ControlStyles.DoubleBuffer, True)
    SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, False)
    SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw, True)
    SetStyle(System.Windows.Forms.ControlStyles.UserPaint, True)
    SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    Me.BackColor = System.Drawing.Color.Transparent
    mvBackgroundExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betToolbar)
    mvButtonExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betToolbarButton)
  End Sub

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    Me.BackColor = System.Drawing.Color.Transparent
    MyBase.OnPaintBackground(e)
    'Debug.WriteLine("On Paint Background " & Me.ClientRectangle.ToString)
    mvBackgroundExtender.PaintBackGround(Me.ClientRectangle, e)

    If Not Me.ImageList Is Nothing Then
      Dim vImageWidth As Integer = Me.ImageList.ImageSize.Width
      Dim vImageHeight As Integer = Me.ImageList.ImageSize.Height
      Dim vPoint As Point = Me.PointToClient(System.Windows.Forms.Cursor.Position)
      For Each vButton As ToolBarButton In Me.Buttons
        Select Case vButton.Style
          Case ToolBarButtonStyle.Separator
            Dim vTop As Integer = ((vButton.Rectangle.Height - vImageHeight) \ 2) + 1
            Dim vX As Integer = vButton.Rectangle.X + 1
            e.Graphics.DrawLine(SystemPens.Highlight, vX, vTop, vX, vImageHeight - 1)
            e.Graphics.DrawLine(Pens.White, vX + 1, vTop + 1, vX + 1, vImageHeight)
          Case ToolBarButtonStyle.DropDownButton
            'Do nothing
          Case Else
            Dim vBorderRect As New Rectangle(vButton.Rectangle.Location, vButton.Rectangle.Size)
            vBorderRect.Inflate(-1, -1)
            If vButton.Enabled Then
              If vButton.Rectangle.Contains(vPoint) Then       'Cursor is over the button
                mvButtonExtender.DrawPushed = False
                mvButtonExtender.DrawSelected = mvMouseDown
                mvButtonExtender.PaintBackGround(vBorderRect, e)
                mvLastButton = vButton
              ElseIf vButton.Style = ToolBarButtonStyle.ToggleButton And vButton.Pushed Then
                mvButtonExtender.DrawPushed = True
                mvButtonExtender.DrawSelected = False
                mvButtonExtender.PaintBackGround(vBorderRect, e)
              End If
            End If
            Dim vRect As Rectangle = vButton.Rectangle
            Dim vX As Integer = (vRect.Width - vImageWidth) \ 2
            Dim vY As Integer = (vRect.Height - vImageHeight) \ 2
            Dim vDest As New Rectangle(vRect.X + vX, vRect.Y + vY, vImageWidth, vImageHeight)
            If vButton.ImageIndex >= 0 Then
              If vButton.Enabled Then
                e.Graphics.DrawImage(Me.ImageList.Images(vButton.ImageIndex), vDest)
              Else
                ControlPaint.DrawImageDisabled(e.Graphics, Me.ImageList.Images(vButton.ImageIndex), vDest.Left, vDest.Top, Color.Transparent)
              End If
            End If
        End Select
      Next
    End If
  End Sub

  Private Sub ToolbarEx_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseHover
    Dim vPoint As Point = Me.PointToClient(System.Windows.Forms.Cursor.Position)
    For Each vButton As ToolBarButton In Me.Buttons
      Select Case vButton.Style
        Case ToolBarButtonStyle.Separator, ToolBarButtonStyle.DropDownButton
          'Do nothing
        Case Else
          If vButton.Rectangle.Contains(vPoint) Then Me.Invalidate(vButton.Rectangle)
      End Select
    Next
  End Sub

  Private Sub ToolbarEx_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
    If e.Button = System.Windows.Forms.MouseButtons.Left Then
      Dim vPoint As Point = Me.PointToClient(System.Windows.Forms.Cursor.Position)
      For Each vButton As ToolBarButton In Me.Buttons
        Select Case vButton.Style
          Case ToolBarButtonStyle.Separator, ToolBarButtonStyle.DropDownButton
            'Do nothing
          Case Else
            If vButton.Rectangle.Contains(vPoint) Then
              mvMouseDown = True
              Me.Invalidate(vButton.Rectangle)
            End If
        End Select
      Next
    End If
  End Sub

  Private Sub ToolbarEx_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
    mvMouseDown = False
  End Sub

  Private Sub ToolbarEx_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseLeave
    mvMouseDown = False
    If Not mvLastButton Is Nothing Then Me.Invalidate(mvLastButton.Rectangle)
  End Sub

  Private Sub ToolbarEx_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
    Dim vImageIndex As Integer
    Dim vButton As ToolBarButton = Nothing

    If Not DisplayTheme.IsXPThemeActive Then
      Dim vPoint As Point = Me.PointToClient(System.Windows.Forms.Cursor.Position)
      For Each vButton In Me.Buttons
        Select Case vButton.Style
          Case ToolBarButtonStyle.Separator, ToolBarButtonStyle.DropDownButton
            'Do nothing
          Case Else
            If vButton.Rectangle.Contains(vPoint) Then
              vImageIndex = vButton.ImageIndex
              Exit For
            End If
        End Select
      Next
      If mvLastMovedOver <> vImageIndex Then
        mvLastMovedOver = vImageIndex
        If Not mvLastButton Is Nothing And Not vButton Is Nothing Then
          Me.Invalidate(Rectangle.Union(mvLastButton.Rectangle, vButton.Rectangle))
        Else
          Me.Invalidate()
        End If
      End If
    End If
  End Sub
End Class
