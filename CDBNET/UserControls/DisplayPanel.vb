Public Class DisplayPanel
  Inherits System.Windows.Forms.UserControl
  Implements IThemeSettable

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
    '
    'DisplayPanel
    '
    Me.BackColor = System.Drawing.SystemColors.Control
    Me.Name = "DisplayPanel"
    Me.Size = New System.Drawing.Size(304, 144)

  End Sub

#End Region

  Private Const LABEL_WIDTH As Integer = 100
  Private Const TEMP_COL_WIDTH As Integer = 200
  Private Const LABEL_GAP As Integer = 4             'Gap from label to item
  Private Const MINIMUM_COLUMN_WIDTH As Integer = 20

  Private mvDefaultHeight As Integer = 20
  Private mvHeightOffset As Integer = 24       'Distance from one item to the next vertically
  Private mvAutoSetHeight As Boolean = False
  Private mvProcessResize As Boolean
  Private mvTextHeight As Integer
  Private mvLastWidth As Integer
  Private mvColumns As List(Of DisplayColumnInfo)
  Private mvRows As List(Of DisplayRowInfo)
  Private mvDisplayItems(0, 0) As DisplayItem
  Private WithEvents mvBackgroundExtender As BackGroundExtender

  Private Declare Auto Function SendMessage Lib "user32.dll" ( _
      ByVal hWnd As IntPtr, _
      ByVal wMsg As Int32, _
      ByVal wParam As Boolean, _
      ByVal lParam As Int32 _
  ) As Int32

  Private Const WM_SETREDRAW As Int32 = &HB

  Public Sub InitialiseControls()
    SetStyle(System.Windows.Forms.ControlStyles.DoubleBuffer, True)
    SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, True)
    SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw, True)
    SetStyle(System.Windows.Forms.ControlStyles.UserPaint, True)
    SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    Me.BackColor = System.Drawing.Color.Transparent
    mvBackgroundExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betDisplayPanel)
  End Sub

  Public Sub Init(ByVal pDataSet As DataSet)
    Dim vControls(50) As Control
    Dim vControlCount As Integer = 0

    mvProcessResize = False
    Me.SuspendLayout()
    If pDataSet.Tables.Contains("Column") Then
      Dim vTable As DataTable = pDataSet.Tables("Column")
      Dim vRow As DataRow
      Dim vDetail As Boolean
      Dim vName As String
      Dim vHeading As String
      Dim vTop As Integer = DisplayTheme.DisplayPanelBorderY
      Dim vLeft As Integer = DisplayTheme.DisplayPanelBorderX
      Dim vMaxHeight As Integer
      Dim vRowNumber As Integer
      Dim vColumnNumber As Integer

      mvColumns = New List(Of DisplayColumnInfo)
      mvRows = New List(Of DisplayRowInfo)
      'First set up all the columns by adding a displaycolumninfo item for each one
      Dim vDCI As New DisplayColumnInfo
      mvColumns.Add(vDCI)
      Dim vDRI As New DisplayRowInfo
      mvRows.Add(vDRI)
      For Each vRow In vTable.Rows
        vName = DirectCast(vRow.Item("Name"), String)
        If vDetail = True Then
          Select Case vName
            Case "NewColumn", "NewColumn2", "NewColumn3"
              vDCI = New DisplayColumnInfo
              vDCI.PreferredWidth = MINIMUM_COLUMN_WIDTH
              mvColumns.Add(vDCI)
            Case Else
          End Select
        ElseIf vName = "DetailItems" Then
          vDetail = True
        End If
      Next
      ReDim mvDisplayItems(mvColumns.Count - 1, 0)

      'Now add the required controls 
      vDCI = mvColumns(0)
      vDetail = False
      For Each vRow In vTable.Rows
        vName = DirectCast(vRow.Item("Name"), String)
        If vDetail = True Then
          Select Case vName
            Case "RowCount"
              'Ignore this
            Case "NewColumn"
              vDCI = mvColumns(1)
              vTop = DisplayTheme.DisplayPanelBorderY
              vLeft += TEMP_COL_WIDTH
              vColumnNumber += 1
              vRowNumber = 0
            Case "NewColumn2"
              vDCI = mvColumns(2)
              vTop = DisplayTheme.DisplayPanelBorderY
              vLeft += TEMP_COL_WIDTH
              vColumnNumber += 1
              vRowNumber = 0
            Case "NewColumn3"
              vDCI = mvColumns(3)
              vTop = DisplayTheme.DisplayPanelBorderY
              vLeft += TEMP_COL_WIDTH
              vColumnNumber += 1
              vRowNumber = 0
            Case Else
              If vRowNumber >= mvDisplayItems.GetLength(1) Then ReDim Preserve mvDisplayItems(mvDisplayItems.GetUpperBound(0), vRowNumber)
              Dim vDisplayItem As New DisplayItem
              mvDisplayItems(vColumnNumber, vRowNumber) = vDisplayItem
              If vRowNumber >= mvRows.Count Then
                vDRI = New DisplayRowInfo
                mvRows.Add(vDRI)
              End If
              vHeading = DirectCast(vRow.Item("Heading"), String)
              Dim vLabelWidth As Integer = 0
              If vHeading.Length > 0 And vHeading <> "." Then
                Dim vLabel As New LabelEx
                vDisplayItem.Label = vLabel
                With vLabel
                  .Visible = False
                  .Name = vName & "_Label"
                  .Text = vHeading
                  .BackColor = DisplayTheme.DisplayPanelTheme.BackColor1
                  vLabelWidth = DisplayTheme.GetFontSize(Me, .Text).Width + (.CurvatureWidth * 2) + LABEL_GAP
                  If vLabelWidth > vDCI.LabelWidth Then
                    vDCI.LabelWidth = vLabelWidth
                    vDCI.MaxLabelText = vHeading
                  End If
                  .SetBounds(vLeft, vTop, vLabelWidth, mvDefaultHeight)
                End With
                vControls(vControlCount) = vLabel
                vControlCount += 1
                If vControlCount >= vControls.Length Then ReDim Preserve vControls(vControlCount)
              End If

              If vName.StartsWith("Spacer") Then
                If vColumnNumber > 0 Then
                  Dim vPrevColumn As Integer = vColumnNumber - 1
                  Dim vFound As Boolean = False
                  Do
                    If mvDisplayItems(vPrevColumn, vRowNumber) IsNot Nothing AndAlso mvDisplayItems(vPrevColumn, vRowNumber).TextBox IsNot Nothing Then
                      mvDisplayItems(vPrevColumn, vRowNumber).ColumnSpan += 1
                      vFound = True
                    End If
                    vPrevColumn -= 1
                  Loop While vPrevColumn >= 0 And vFound = False
                End If
              ElseIf vName.EndsWith("_") Then
                Dim vPB As New StandardPercentageBar
                vDisplayItem.PercentageBar = vPB
                With vPB
                  .Visible = False
                  .Name = vName
                  .BackColor = DisplayTheme.DisplayPanelTheme.BackColor1
                  .TabStop = False
                  .SetBounds(vLeft, vTop, LABEL_WIDTH, mvDefaultHeight)
                End With
                vControls(vControlCount) = vPB
                vControlCount += 1
                If vControlCount >= vControls.Length Then ReDim Preserve vControls(vControlCount)
              Else
                Dim vTextBox As New TextBox
                vDisplayItem.TextBox = vTextBox
                With vTextBox
                  .Visible = False
                  .Name = vName
                  If vName = "Notes" Or vName = "Precis" Or vName = "StatusReason" Then
                    .Multiline = True
                    .WordWrap = True
                  End If
                  .BorderStyle = DisplayTheme.DisplayDataTheme.BorderStyle
                  If DisplayTheme.DisplayDataTheme.BackColor1.A = 0 Then
                    .BackColor = DisplayTheme.DisplayPanelTheme.BackColor1
                  Else
                    .BackColor = DisplayTheme.DisplayDataTheme.BackColor1
                  End If
                  .ReadOnly = True
                  .TabStop = False
                  .SetBounds(vLeft, vTop, LABEL_WIDTH, mvDefaultHeight)
                End With
                vControls(vControlCount) = vTextBox
                vControlCount += 1
                If vControlCount >= vControls.Length Then ReDim Preserve vControls(vControlCount)
              End If
              vTop += mvHeightOffset
              vRowNumber += 1
              If vTop > vMaxHeight Then vMaxHeight = vTop
          End Select
        ElseIf vName = "DetailItems" Then
          vDetail = True
        End If
      Next
      'ResizeControls()
      If mvAutoSetHeight Then
        Me.Height = vMaxHeight
      End If
    End If
    Me.Controls.Clear()
    ReDim Preserve vControls(vControlCount - 1)
    Me.Controls.AddRange(vControls)
    Me.ResumeLayout()
  End Sub

  Public Sub Populate(ByVal pDataSet As DataSet, ByVal pRow As Integer)
    Dim vTextWidth As Integer
    Dim vRowNumber As Integer
    Dim vColumnNumber As Integer
    Dim vDCI As DisplayColumnInfo
    Dim vDRI As DisplayRowInfo

    If Me.Controls.Count = 0 Then
      Clear()
    Else
      For Each vDCI In mvColumns
        vDCI.PreferredWidth = MINIMUM_COLUMN_WIDTH
        vDCI.FillWidth = False
      Next
      For Each vDRI In mvRows
        vDRI.RowHeight = mvHeightOffset
        vDRI.Lines = 1
      Next
      If pDataSet.Tables.Contains("DataRow") Then
        Dim vRow As DataRow = pDataSet.Tables("DataRow").Rows(pRow)
        Dim vDisplayItem As DisplayItem
        For vRowNumber = 0 To mvDisplayItems.GetUpperBound(1)
          For vColumnNumber = 0 To mvDisplayItems.GetUpperBound(0)
            vDisplayItem = mvDisplayItems(vColumnNumber, vRowNumber)
            If Not vDisplayItem Is Nothing Then
              If vDisplayItem.TextBox IsNot Nothing Then
                Dim vTextBox As TextBox = vDisplayItem.TextBox
                vTextBox.ScrollBars = ScrollBars.None
                vTextBox.Text = MultiLine(vRow.Item(vTextBox.Name).ToString)
                If vTextBox.Text.Length > 0 Then
                  vDCI = mvColumns(vColumnNumber)
                  If mvTextHeight = 0 Then
                    mvTextHeight = DisplayTheme.GetFontSize(Me).Height
                  End If
                  If vTextBox.Multiline = True AndAlso vTextBox.Text.Length > 30 Then
                    vTextWidth = DisplayTheme.GetFontSize(Me, vTextBox.Text.Substring(1, 30)).Width
                    vDCI.FillWidth = True
                    vTextBox.ScrollBars = ScrollBars.Vertical
                    If vTextBox.Lines.Length > 1 Then
                      Dim vLines As Integer = vTextBox.Lines.Length
                      If vLines > 5 Then vLines = 5
                      vDRI = mvRows(vRowNumber)
                      vDRI.Lines = vLines
                      vDRI.RowHeight = (vLines * mvTextHeight) + (SystemInformation.Border3DSize.Height * 2) + (mvHeightOffset - mvDefaultHeight)
                    End If
                    If vTextWidth > vDCI.PreferredWidth Then
                      vDCI.PreferredWidth = vTextWidth
                      vDCI.MaxText = vTextBox.Text.Substring(1, 30)
                    End If
                  Else
                    vTextWidth = DisplayTheme.GetFontSize(Me, vTextBox.Text).Width
                    If vTextWidth > vDCI.PreferredWidth Then
                      vDCI.PreferredWidth = vTextWidth
                      vDCI.MaxText = vTextBox.Text
                    End If
                  End If
                End If
              ElseIf (vDisplayItem.PercentageBar IsNot Nothing) Then
                Dim vPB As StandardPercentageBar = vDisplayItem.PercentageBar
                vPB.Maximum = 200
                Dim vValue As Integer = CInt(DoubleValue(vRow.Item(vPB.Name).ToString) * 100)
                If vValue > vPB.Maximum Then vValue = vPB.Maximum
                vPB.Value = vValue
              End If
            End If
          Next
        Next
      End If
      ResizeControls()
    End If
  End Sub

  Private Sub ResizeControls()
    Dim vColumnLeft As Integer
    Dim vTop As Integer = DisplayTheme.DisplayPanelBorderY
    Dim vRowNumber As Integer
    Dim vColumnNumber As Integer
    Dim vDCI As DisplayColumnInfo
    Dim vDRI As DisplayRowInfo
    Dim vDisplayItem As DisplayItem

    'Debug.WriteLine("Resizing Display Panel " & Date.Now)

    Me.SuspendLayout()
    vColumnLeft = DisplayTheme.DisplayPanelBorderX
    For Each vDCI In mvColumns
      With vDCI
        .LabelLeft = vColumnLeft
        .ItemLeft = vColumnLeft + .LabelWidth
        .ItemWidth = .PreferredWidth
        vColumnLeft += (.PreferredWidth + .LabelWidth)
      End With
    Next
    If Me.Width > vColumnLeft + DisplayTheme.DisplayPanelBorderX Then
      Dim vExtra As Integer = Me.Width - (vColumnLeft + DisplayTheme.DisplayPanelBorderX)
      Dim vUsed As Integer
      For Each vDCI In mvColumns
        With vDCI
          If .FillWidth And vExtra > 0 Then
            .ItemWidth += vExtra
            vUsed = vExtra
            vExtra = 0
          Else
            .LabelLeft += vUsed
            .ItemLeft += vUsed
          End If
        End With
      Next
    End If
    For Each vDRI In mvRows
      vDRI.RowTop = vTop
      vTop += vDRI.RowHeight
    Next
    For vRowNumber = 0 To mvDisplayItems.GetUpperBound(1)
      vDRI = mvRows(vRowNumber)
      For vColumnNumber = 0 To mvDisplayItems.GetUpperBound(0)
        vDisplayItem = mvDisplayItems(vColumnNumber, vRowNumber)
        If Not vDisplayItem Is Nothing Then
          vDCI = mvColumns(vColumnNumber)
          If Not vDisplayItem.Label Is Nothing Then
            vDisplayItem.Label.SetBounds(vDCI.LabelLeft, vDRI.RowTop, vDCI.LabelWidth - LABEL_GAP, mvDefaultHeight)
            vDisplayItem.Label.Visible = True
          End If
          If vDisplayItem.TextBox IsNot Nothing Then
            Dim vWidth As Integer = 0
            Dim vColumn As Integer
            If vDisplayItem.ColumnSpan > 0 Then
              For vColumn = vColumnNumber To vColumnNumber + vDisplayItem.ColumnSpan
                vWidth += mvColumns(vColumn).ItemWidth
                If vColumn > vColumnNumber Then vWidth += mvColumns(vColumn).LabelWidth
              Next
            Else
              vWidth = vDCI.ItemWidth
            End If
            vDisplayItem.TextBox.SetBounds(vDCI.ItemLeft, vDRI.RowTop, vWidth, vDRI.RowHeight)
            vDisplayItem.TextBox.Visible = True
          ElseIf vDisplayItem.PercentageBar IsNot Nothing Then
            vDisplayItem.PercentageBar.Visible = True
            vDisplayItem.PercentageBar.SetBounds(vDCI.ItemLeft + LABEL_GAP, vDRI.RowTop, LABEL_WIDTH, mvDefaultHeight)
          End If
        End If
      Next
    Next
    vTop += DisplayTheme.DisplayPanelBorderY
    mvLastWidth = Me.Width        'Set last width here to stop antother resize event from the next line of code
    If mvAutoSetHeight AndAlso vTop > Me.Height Then Me.Height = vTop
    'If TypeOf (Me.Parent) Is PanelEx Then Me.Parent.Height = Me.Height
    Me.ResumeLayout()
  End Sub

  Public Sub Clear()
    Me.Controls.Clear()
    If mvAutoSetHeight Then Me.Height = 0
    mvColumns = New List(Of DisplayColumnInfo)
    mvRows = New List(Of DisplayRowInfo)
    mvDisplayItems = Nothing
  End Sub

  Public Property AutoSetHeight() As Boolean
    Get
      Return mvAutoSetHeight
    End Get
    Set(ByVal pValue As Boolean)
      mvAutoSetHeight = pValue
    End Set
  End Property

  Public Property ProcessResize() As Boolean
    Get
      Return mvProcessResize
    End Get
    Set(ByVal Value As Boolean)
      mvProcessResize = Value
    End Set
  End Property

  'Private Sub DisplayPanel_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
  'Debug.WriteLine("Got SizeChanged Event " & Date.Now)
  'If Me.Controls.Count > 0 And Me.Width <> mvLastWidth And mvProcessResize Then ResizeControls()
  'End Sub

  'Private Sub DisplayPanel_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
  '  'Debug.WriteLine("Got ReSize Event " & Date.Now)
  'End Sub

  Private Class DisplayItem
    Friend TextBox As TextBox
    Friend Label As Label
    Friend PercentageBar As StandardPercentageBar
    Friend ColumnSpan As Integer
  End Class

  Private Class DisplayRowInfo
    Friend RowTop As Integer
    Friend RowHeight As Integer
    Friend Lines As Integer
  End Class

  Private Class DisplayColumnInfo
    Friend LabelLeft As Integer
    Friend LabelWidth As Integer
    Friend MaxLabelText As String
    Friend ItemLeft As Integer
    Friend ItemWidth As Integer
    Friend PreferredWidth As Integer
    Friend MaxText As String
    Friend FillWidth As Boolean
  End Class

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    If e.ClipRectangle.Left = 0 OrElse e.ClipRectangle.Top = 0 OrElse e.ClipRectangle.Right = Me.Right OrElse e.ClipRectangle.Bottom = Me.Bottom Then
      Me.BackColor = System.Drawing.Color.Transparent
      MyBase.OnPaintBackground(e)
    End If
    'Me.BackColor = System.Drawing.Color.White
    'MyBase.OnPaintBackground(e)
    mvBackgroundExtender.PaintBackGround(Me.ClientRectangle, e)
  End Sub

  Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
    MyBase.OnResize(e)
    If Not mvBackgroundExtender Is Nothing Then mvBackgroundExtender.HandleResize()
  End Sub

  Private Sub ResizeHandler() Handles mvBackgroundExtender.ReSize
    If Me.Controls.Count > 0 And Me.Width <> mvLastWidth And mvProcessResize Then ResizeControls()
    Invalidate()
  End Sub

  Public Sub SetControlTheme() Implements CDBNETCL.IThemeSettable.SetControlTheme
    Dim vOldHeightOffset As Integer = mvHeightOffset
    mvDefaultHeight = DisplayTheme.GetFontSize(Me).Height + 4
    mvHeightOffset = mvDefaultHeight + 4
    If DisplayTheme.DisplayDataTheme.BorderStyle = Windows.Forms.BorderStyle.FixedSingle Then mvHeightOffset += 1
    If mvRows IsNot Nothing Then
      For Each vDRI As DisplayRowInfo In mvRows
        If vDRI.Lines = 1 Then
          vDRI.RowHeight = mvHeightOffset
        Else
          vDRI.RowHeight = (vDRI.Lines * mvDefaultHeight) + (SystemInformation.Border3DSize.Height * 2) + (mvHeightOffset - mvDefaultHeight)
        End If
      Next
    End If
    If mvColumns IsNot Nothing Then
      For Each vDCI As DisplayColumnInfo In mvColumns
        If vDCI.MaxLabelText IsNot Nothing Then
          vDCI.LabelWidth = DisplayTheme.GetFontSize(Me, vDCI.MaxLabelText).Width + (DisplayTheme.DisplayLabelTheme.Curvature * 2) + LABEL_GAP
        End If
        If vDCI.MaxText IsNot Nothing Then
          vDCI.PreferredWidth = DisplayTheme.GetFontSize(Me, vDCI.MaxText).Width
        End If
      Next
    End If
    For Each vControl As Control In Me.Controls
      If TypeOf (vControl) Is TextBox Then
        DirectCast(vControl, TextBox).BorderStyle = DisplayTheme.DisplayDataTheme.BorderStyle
        If DisplayTheme.DisplayDataTheme.BackColor1.A = 0 Then
          vControl.BackColor = DisplayTheme.DisplayPanelTheme.BackColor1
        Else
          vControl.BackColor = DisplayTheme.DisplayDataTheme.BackColor1
        End If
      End If
    Next
    If Me.Controls.Count > 0 AndAlso vOldHeightOffset <> mvHeightOffset Then ResizeControls()
    Me.Invalidate()
  End Sub
End Class