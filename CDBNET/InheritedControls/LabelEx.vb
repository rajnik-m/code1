Public Class LabelEx
  Inherits System.Windows.Forms.Label

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    InitialiseControls()
    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

  Dim mvBackGroundExtender As BackGroundExtender

  Private Sub InitialiseControls()
    SetStyle(System.Windows.Forms.ControlStyles.DoubleBuffer, True)
    SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, True)
    SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw, True)
    SetStyle(System.Windows.Forms.ControlStyles.UserPaint, True)
    SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    Me.BackColor = System.Drawing.Color.Transparent
    mvBackGroundExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betDisplayLabel)
  End Sub

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)

    Me.BackColor = System.Drawing.Color.Transparent
    MyBase.OnPaintBackground(e)
    mvBackGroundExtender.PaintBackGround(Me.ClientRectangle, e)
  End Sub

  Public ReadOnly Property CurvatureWidth() As Integer
    Get
      Return mvBackGroundExtender.Curvature
    End Get
  End Property

  Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
    With e.Graphics
      .DrawString(Me.Text, Me.Font, Brushes.Black, New PointF(mvBackGroundExtender.TextXOffset, 0))
    End With
  End Sub
End Class
