Public Class FormHeader
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
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents pnlCommands As System.Windows.Forms.Panel
  Friend WithEvents ttp As System.Windows.Forms.ToolTip
  Friend WithEvents cmdStickyNote As CDBNETCL.ImageButton
  Friend WithEvents ImageList As System.Windows.Forms.ImageList
  Friend WithEvents cmsContactPic As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents AddPictureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents cmdAction As CDBNETCL.ImageButton
  Friend WithEvents pnlContextMenu As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents pnlContextCustomise As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents pnlContextRevert As System.Windows.Forms.ToolStripMenuItem
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormHeader))
    Me.pnl = New System.Windows.Forms.Panel
    Me.pnlCommands = New System.Windows.Forms.Panel
    Me.cmdAction = New CDBNETCL.ImageButton
    Me.cmdStickyNote = New CDBNETCL.ImageButton
    Me.ttp = New System.Windows.Forms.ToolTip(Me.components)
    Me.cmsContactPic = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.AddPictureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
    Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
    Me.pnlContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.pnlContextCustomise = New System.Windows.Forms.ToolStripMenuItem()
    Me.pnlContextRevert = New System.Windows.Forms.ToolStripMenuItem()
    Me.pnlCommands.SuspendLayout()
    Me.cmsContactPic.SuspendLayout()
    Me.pnlContextMenu.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnl
    '
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(424, 64)
    Me.pnl.TabIndex = 0
    '
    'pnlCommands
    '
    Me.pnlCommands.Controls.Add(Me.cmdAction)
    Me.pnlCommands.Controls.Add(Me.cmdStickyNote)
    Me.pnlCommands.Dock = System.Windows.Forms.DockStyle.Right
    Me.pnlCommands.Location = New System.Drawing.Point(424, 0)
    Me.pnlCommands.Name = "pnlCommands"
    Me.pnlCommands.Size = New System.Drawing.Size(48, 64)
    Me.pnlCommands.TabIndex = 1
    '
    'cmdAction
    '
    Me.cmdAction.AccessibleName = "Actions"
    Me.cmdAction.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
    Me.cmdAction.BackColor = System.Drawing.Color.Transparent
    Me.cmdAction.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdAction.DrawBorder = False
    Me.cmdAction.Location = New System.Drawing.Point(4, 32)
    Me.cmdAction.Name = "cmdAction"
    Me.cmdAction.NormalImage = CType(resources.GetObject("cmdAction.NormalImage"), System.Drawing.Image)
    Me.cmdAction.ShowFocusRect = True
    Me.cmdAction.Size = New System.Drawing.Size(32, 32)
    Me.cmdAction.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.CenterImage
    Me.cmdAction.TabIndex = 2
    Me.cmdAction.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
    Me.ttp.SetToolTip(Me.cmdAction, "Actions")
    Me.cmdAction.TransparentColor = System.Drawing.Color.Transparent
    '
    'cmdStickyNote
    '
    Me.cmdStickyNote.AccessibleName = "Sticky Notes"
    Me.cmdStickyNote.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
    Me.cmdStickyNote.BackColor = System.Drawing.Color.Transparent
    Me.cmdStickyNote.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdStickyNote.DrawBorder = False
    Me.cmdStickyNote.Location = New System.Drawing.Point(4, 0)
    Me.cmdStickyNote.Name = "cmdStickyNote"
    Me.cmdStickyNote.NormalImage = CType(resources.GetObject("cmdStickyNote.NormalImage"), System.Drawing.Image)
    Me.cmdStickyNote.ShowFocusRect = True
    Me.cmdStickyNote.Size = New System.Drawing.Size(32, 32)
    Me.cmdStickyNote.SizeMode = CDBNETCL.ImageButton.ImageButtonSizeMode.CenterImage
    Me.cmdStickyNote.TabIndex = 1
    Me.cmdStickyNote.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    Me.ttp.SetToolTip(Me.cmdStickyNote, "Sticky Notes")
    Me.cmdStickyNote.TransparentColor = System.Drawing.Color.Transparent
    '
    'cmsContactPic
    '
    Me.cmsContactPic.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddPictureToolStripMenuItem})
    Me.cmsContactPic.Name = "cmsContactPic"
    Me.cmsContactPic.Size = New System.Drawing.Size(165, 26)
    '
    'AddPictureToolStripMenuItem
    '
    Me.AddPictureToolStripMenuItem.Name = "AddPictureToolStripMenuItem"
    Me.AddPictureToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
    Me.AddPictureToolStripMenuItem.Text = "Add Picture"
    '
    'ImageList
    '
    Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
    Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
    Me.ImageList.Images.SetKeyName(0, "businessman.png")
    '
    'pnlContextMenu
    '
    Me.pnlContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.pnlContextCustomise, pnlContextRevert})
    Me.pnlContextMenu.Name = "pnlContextMenu"
    Me.pnlContextMenu.Size = New System.Drawing.Size(159, 26)
    '
    'pnlContextCustomise
    '
    Me.pnlContextCustomise.Name = "pnlContextCustomise"
    Me.pnlContextCustomise.Size = New System.Drawing.Size(158, 22)
    Me.pnlContextCustomise.Text = "Customise"
    '
    'pnlContextRevert
    '
    Me.pnlContextRevert.Name = "pnlContextRevert"
    Me.pnlContextRevert.Size = New System.Drawing.Size(158, 22)
    Me.pnlContextRevert.Text = "Revert"
    '
    'FormHeader
    '
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.pnlCommands)
    Me.Name = "FormHeader"
    Me.Size = New System.Drawing.Size(472, 64)
    Me.pnlCommands.ResumeLayout(False)
    Me.cmsContactPic.ResumeLayout(False)
    Me.pnlContextMenu.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Const START_LEFT As Integer = 6           '80 Offset of items from start of column
  Private Const START_TOP As Integer = 4
  Private Const LABEL_WIDTH As Integer = 100
  Private Const LABEL_GAP As Integer = 4             'Gap from label to item
  Private Const MINIMUM_COLUMN_WIDTH As Integer = 20
  Private Const TEXT_EXTRA As Integer = 8
  Private Const COLUMN_GAP As Integer = 8

  Private Const HEIGHT_EXTRA As Integer = 6

  Private mvContactType As String
  Private mvColumns As Collection
  Private mvContactInfo As ContactInfo
  Private WithEvents mvBackgroundExtender As BackGroundExtender
  Private mvControls() As Control
  Private mvControlCount As Integer
  Private mvRequiredHeight As Integer
  Private mvBackColor As Color
  Private mvDefaultHeight As Integer = 18
  Private mvHeightOffset As Integer = 22          'Distance from one item to the next vertically
  Private mvDataSet As New DataSet
  Event PanelMouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
  Event PanelMouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)

  Public Sub InitialiseControls()
    'SetStyle(System.Windows.Forms.ControlStyles.DoubleBuffer, True)
    'SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, False)
    'SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw True)
    'SetStyle(System.Windows.Forms.ControlStyles.UserPaint, True)
    'SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    'Me.BackColor = System.Drawing.Color.Transparent
    'mvBackgroundExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betFormHeader)
    pnlCommands.Visible = False
    If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciDisplayListMaintenance) Then
      pnl.ContextMenuStrip = pnlContextMenu
    End If
  End Sub

  Public Sub Populate(ByVal pEventInfo As CareEventInfo)
    pnlCommands.Visible = False
    Dim vDataSet As DataSet = DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEventHeaderInfo, pEventInfo.EventNumber)
    mvDataSet = vDataSet
    DoPopulate(vDataSet)
  End Sub

  Public Sub Populate(ByVal pDataSet As DataSet)    'Used by the DashBoard
    pnlCommands.Visible = False
    DoPopulate(pDataSet)
  End Sub

  Public Sub Populate(ByVal pDataSet As DataSet, ByVal pContactInfo As ContactInfo)
    pnlCommands.Visible = True
    mvContactInfo = pContactInfo
    mvDataSet = pDataSet
    DoPopulate(pDataSet)
  End Sub

  Public Overrides Property AutoScroll() As Boolean
    Get
      Return pnl.AutoScroll
    End Get
    Set(ByVal value As Boolean)
      pnl.AutoScroll = value
    End Set
  End Property

  Private Sub DoPopulate(ByVal pDataSet As DataSet)
    Dim vTotalWidth As Integer

    If pDataSet.Tables.Contains("Column") And pDataSet.Tables.Contains("DataRow") Then
      Dim vColumnTable As DataTable = pDataSet.Tables("Column")
      Dim vDataRow As DataRow = pDataSet.Tables("DataRow").Rows(0)
      Dim vRow As DataRow
      Dim vName As String
      Dim vColumnWidth As Integer
      Dim vWidth As Integer
      Dim vSetHeight As Integer
      Dim vTopHeight As Integer
      Dim vSize As Integer
      Dim vTop As Integer = START_TOP
      Dim vLeft As Integer = START_LEFT
      Dim vMaxHeight As Integer
      Dim vParent As Control

      Me.SuspendLayout()
      vParent = pnl
      ReDim mvControls(50)
      mvControlCount = 0
      mvColumns = New Collection
      Dim vDCI As New HeaderColumnInfo
      mvColumns.Add(vDCI)
      For Each vRow In vColumnTable.Rows
        Select Case vRow.Item("Name").ToString
          Case "NewColumn", "NewColumn2", "NewColumn3"
            vDCI = New HeaderColumnInfo
            vDCI.PreferredWidth = MINIMUM_COLUMN_WIDTH
            mvColumns.Add(vDCI)
        End Select
      Next
      vTotalWidth = pnl.Width
      vColumnWidth = (vTotalWidth \ mvColumns.Count)
      If mvColumns.Count = 1 Then
        vWidth = vTotalWidth
      ElseIf mvColumns.Count = 2 Then
        vWidth = vColumnWidth * 2
      Else
        vWidth = vColumnWidth
      End If
      vWidth = vWidth - START_LEFT
      If vDataRow.Table.Columns.Contains("GroupRGBValue") Then
        Dim vValue As Integer = CInt(vDataRow.Item("GroupRGBValue"))
        mvBackColor = Color.FromArgb(vValue And 255, (vValue \ 256) And 255, (vValue \ 65536) And 255)
      Else
        mvBackColor = Color.BlanchedAlmond
      End If
      SetControlTheme()
      vParent.Controls.Clear()
      Dim vGraphics As Graphics = Me.CreateGraphics
      mvDefaultHeight = DisplayTheme.GetFontSize(Me).Height
      mvHeightOffset = mvDefaultHeight + HEIGHT_EXTRA
      vDCI = DirectCast(mvColumns(1), HeaderColumnInfo)
      For Each vRow In vColumnTable.Rows
        vName = vRow.Item("Name").ToString
        Select Case vName
          Case "StickyNoteCount"
            If IntegerValue(vDataRow.Item(vName).ToString) > 0 Then
              cmdStickyNote.NormalImage = AppHelper.ImageProvider.NewOtherImages32.Images("EditStickyNote")   'imgToolbar.Images(1)
              ttp.SetToolTip(cmdStickyNote, "Edit Sticky Notes")
            Else
              cmdStickyNote.NormalImage = AppHelper.ImageProvider.NewOtherImages32.Images("NewStickyNote")    'imgToolbar.Images(0)
              ttp.SetToolTip(cmdStickyNote, "Add Sticky Notes")
            End If
          Case "ActionCount"
            If IntegerValue(vDataRow.Item(vName).ToString) > 0 Then
              cmdAction.NormalImage = AppHelper.ImageProvider.NewOtherImages32.Images("EditAction")           'imgToolbar.Images(3)
              ttp.SetToolTip(cmdAction, "Edit Actions")
            Else
              cmdAction.NormalImage = AppHelper.ImageProvider.NewOtherImages32.Images("NewAction")            'imgToolbar.Images(2)
              ttp.SetToolTip(cmdAction, "Add Action")
            End If
          Case "ContactType"
            mvContactType = CType(vDataRow.Item(vName), String)
          Case "NewColumn"
            vDCI = DirectCast(mvColumns(2), HeaderColumnInfo)
            vTop = START_TOP
            vLeft += MINIMUM_COLUMN_WIDTH
          Case "NewColumn2"
            vDCI = DirectCast(mvColumns(3), HeaderColumnInfo)
            vTop = START_TOP
            vLeft += MINIMUM_COLUMN_WIDTH
          Case Else
            If vRow.Item("Visible").ToString = "Y" Then
              If vColumnTable.Columns.Contains("Size") Then
                vSize = CInt(vRow.Item("Size"))
              Else
                vSize = 300
              End If
              If vSize > 300 Then
                vSize = vSize \ 300
                vSetHeight = (mvDefaultHeight * vSize)
                vTopHeight = (mvDefaultHeight * (vSize - 1)) + mvHeightOffset
              Else
                vSetHeight = mvDefaultHeight
                vTopHeight = mvHeightOffset
              End If
              Select Case vName
                Case "Picture"
                  Dim vPicture As New PictureBox
                  vPicture.Tag = vDCI
                  vPicture.SetBounds(vLeft, vTop, vWidth, vSetHeight)
                  vPicture.ContextMenuStrip = cmsContactPic
                  vPicture.BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                  AddControl(vPicture)
                  Try
                    Dim vDataTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPictureDocuments, mvContactInfo.ContactNumber))
                    If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
                      Dim vDocumentNumber As Integer = IntegerValue(vDataTable.Rows(0)("DocumentNumber").ToString)
                      Dim vImageFile As String = DataHelper.GetDocumentFile(vDocumentNumber, ".jpg")
                      If vImageFile.Length > 0 Then
                        Dim vStream As New FileStream(vImageFile, IO.FileMode.Open, IO.FileAccess.Read)
                        vPicture.Image = Image.FromStream(vStream)
                        vStream.Dispose()
                        DataHelper.DeleteTempFile(vImageFile)
                        Dim vRequiredWidth As Integer = (vPicture.Image.Width * vSetHeight) \ vPicture.Image.Height
                        If vRequiredWidth > vDCI.PreferredWidth Then vDCI.PreferredWidth = vRequiredWidth
                        vPicture.SizeMode = PictureBoxSizeMode.Zoom
                      End If
                    Else
                      vPicture.Image = ImageList.Images(0)
                      Dim vRequiredWidth As Integer = (vPicture.Image.Width * vSetHeight) \ vPicture.Image.Height
                      If vRequiredWidth > vDCI.PreferredWidth Then vDCI.PreferredWidth = vRequiredWidth
                      vPicture.SizeMode = PictureBoxSizeMode.Zoom
                    End If
                  Catch vException As Exception
                    DataHelper.HandleException(vException)
                    vPicture.Image = Nothing
                  End Try
                Case "MarketingChart"
                  Dim vChart As New ChartControl
                  vChart.Tag = vDCI
                  vChart.SetBounds(vLeft, vTop, vWidth, vSetHeight)
                  vChart.PopulateFromScores(mvContactInfo.ContactNumber, False)
                  GetControlWidth(vGraphics, "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW", vDCI, vChart.Font)
                  AddControl(vChart)
                Case "CommunicationsList"
                  AddDisplayGrid(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderCommunications, mvContactInfo.ContactNumber, vDCI, vTop, vLeft, vWidth, vSetHeight)
                Case "HighProfileActivitiesList"
                  AddDisplayGrid(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderHighProfileCategories, mvContactInfo.ContactNumber, vDCI, vTop, vLeft, vWidth, vSetHeight)
                Case "DepartmentActivitiesList"
                  AddDisplayGrid(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderDepartmentCategories, mvContactInfo.ContactNumber, vDCI, vTop, vLeft, vWidth, vSetHeight)
                Case "HighProfileLinksList"
                  AddDisplayGrid(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderHighProfileRelationships, mvContactInfo.ContactNumber, vDCI, vTop, vLeft, vWidth, vSetHeight)
                Case Else
                  Dim vHeading As String = vRow.Item("Heading").ToString
                  Dim vLabelWidth As Integer = 0
                  Dim vMultiLine As Boolean
                  If vSize > 1 And vName = "AddressMultiLine" Or vName = "Notes" Then      'Deal with multiline fields that should not be
                    vMultiLine = True
                  Else
                    vSetHeight = mvDefaultHeight
                    vTopHeight = mvHeightOffset
                  End If
                  If vHeading.Length > 0 Then
                    Dim vLabel As New Label
                    With vLabel
                      .Tag = vDCI
                      .Name = vName & "_Label"
                      .Text = vHeading
                      vLabelWidth = DisplayTheme.GetFontSize(Me, vHeading).Width + LABEL_GAP
                      If vLabelWidth > vDCI.LabelWidth Then vDCI.LabelWidth = vLabelWidth
                      .SetBounds(vLeft, vTop, vLabelWidth, vSetHeight)
                    End With
                    AddControl(vLabel)
                  End If
                  Dim vValue As String = vDataRow.Item(vName).ToString
                  If (vValue.EndsWith("#") OrElse vValue.EndsWith("%")) AndAlso IsNumeric(vValue.TrimEnd("#%".ToCharArray)) Then
                    Dim vSTD As New StandardPercentageBar
                    With vSTD
                      .Tag = vDCI
                      .Name = vName
                      Dim vTextWidth As Integer = GetControlWidth(vGraphics, "WWWWWWWWWWWWWWWWW", vDCI, .Font)
                      .SetBounds(vLeft, vTop, vTextWidth, vSetHeight)
                      .TabStop = False
                      If vValue.EndsWith("%") Then
                        .Maximum = 100
                        .ShowStandardBar = False
                        .DrawGradient = False
                      Else
                        .Maximum = 200
                        .ShowStandardBar = True
                        .DrawGradient = True
                      End If
                      Dim vIntValue As Integer = IntegerValue(vValue.TrimEnd("#%".ToCharArray))
                      If vIntValue > .Maximum Then vIntValue = .Maximum
                      .Value = vIntValue
                      AddControl(vSTD)
                    End With

                  ElseIf vName = "WebAddress" Then
                    Dim vLink As New LinkLabel
                    With vLink
                      .Tag = vDCI
                      .Name = vName
                      .AutoSize = False
                      .BorderStyle = Windows.Forms.BorderStyle.None
                      .BackColor = Me.BackColor
                      .Text = MultiLine(CType(vDataRow.Item(vName), String))
                      Dim vTextWidth As Integer = GetControlWidth(vGraphics, .Text & "W", vDCI, .Font, True)
                      .SetBounds(vLeft, vTop, vTextWidth, vSetHeight)
                      AddHandler vLink.LinkClicked, AddressOf LinkClicked
                      .TabStop = False
                    End With
                    AddControl(vLink)
                  Else
                    Dim vTextBox As New TextBox
                    With vTextBox
                      .Tag = vDCI
                      .Name = vName
                      .AutoSize = False
                      If ((vName = "ContactName") Or (vName = "EventDesc")) Then .Font = New Font(Me.Font, FontStyle.Bold)
                      .Multiline = vMultiLine
                      .BorderStyle = Windows.Forms.BorderStyle.None
                      .BackColor = Me.BackColor
                      .ReadOnly = True
                      .TabStop = Settings.TabIntoHeaderPanel
                      .Text = MultiLine(CType(vDataRow.Item(vName), String))
                      Dim vTextWidth As Integer
                      If .Multiline = True AndAlso .Text.Length > 30 Then
                        vTextWidth = GetControlWidth(vGraphics, .Text.Substring(1, 30), vDCI, .Font)
                      Else
                        vTextWidth = GetControlWidth(vGraphics, .Text & "W", vDCI, .Font, True)
                      End If
                      .SetBounds(vLeft, vTop, vTextWidth, vSetHeight)
                    End With
                    AddControl(vTextBox)
                  End If
              End Select
              vTop += vTopHeight
              If vTop > vMaxHeight Then vMaxHeight = vTop
            End If
        End Select
      Next
      ReDim Preserve mvControls(mvControlCount - 1)
      vParent.Controls.AddRange(mvControls)
      Dim vCommandGap As Integer = 0
      If vMaxHeight > cmdStickyNote.Height + cmdAction.Height Then
        vCommandGap = (vMaxHeight - (cmdStickyNote.Height + cmdAction.Height)) \ 3
      End If
      cmdStickyNote.Top = vCommandGap
      cmdAction.Top = cmdStickyNote.Height + (vCommandGap * 2)

      vMaxHeight -= 2

      Me.Height = vMaxHeight
      mvRequiredHeight = vMaxHeight
      vGraphics.Dispose()
    End If
    If mvColumns IsNot Nothing Then ResizeControls()
    Me.ResumeLayout()
  End Sub

  Private Function GetControlWidth(ByVal pGraphics As Graphics, ByVal pString As String, ByVal pDCI As HeaderColumnInfo, ByVal pFont As Font, Optional ByVal pExtra As Boolean = False) As Integer
    Dim vTextWidth As Integer = CInt(pGraphics.MeasureString(pString, pFont).Width)
    If pExtra Then vTextWidth += TEXT_EXTRA
    If vTextWidth > pDCI.PreferredWidth Then pDCI.PreferredWidth = vTextWidth
    Return vTextWidth
  End Function

  Private Sub AddControl(ByVal pControl As Control)
    mvControls(mvControlCount) = pControl
    mvControlCount += 1
    If mvControlCount > mvControls.Length Then ReDim Preserve mvControls(mvControlCount)
  End Sub

  Private Sub FormHeader_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
    If Not mvColumns Is Nothing Then ResizeControls()
  End Sub

  Private Sub AddDisplayGrid(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Integer, ByVal pHCI As HeaderColumnInfo, ByVal pTop As Integer, ByVal pLeft As Integer, ByVal pWidth As Integer, ByVal pHeight As Integer)
    Dim vDGR As New DisplayGrid
    vDGR.Tag = pHCI
    vDGR.SetBounds(pLeft, pTop, pWidth, pHeight)
    vDGR.Populate(pType, pContactNumber)

    'If vDGR.RequiredWidth > pHCI.RequiredGridWidth Then pHCI.RequiredGridWidth = vDGR.RequiredWidth

    If vDGR.RequiredWidth > pHCI.RequiredGridWidth AndAlso vDGR.RequiredWidth <= pWidth Then
      pHCI.RequiredGridWidth = vDGR.RequiredWidth
    Else
      pHCI.RequiredGridWidth = pWidth
    End If
    If pType = CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderHighProfileRelationships Then AddHandler vDGR.ContactSelected, AddressOf dgr_ContactSelected
    'vDGR.SetBackgroundColour(Me.BackColor)
    AddControl(vDGR)
  End Sub

  Private Sub LinkClicked(ByVal pSender As Object, ByVal pArgs As LinkLabelLinkClickedEventArgs)
    Dim vLink As LinkLabel = DirectCast(pSender, LinkLabel)
    vLink.LinkVisited = True
    Dim vForm As New frmBrowser(vLink.Text, True)
    vForm.Show()
  End Sub

  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer)
    Try
      If pContactNumber <> mvContactInfo.ContactNumber Then
        FormHelper.ShowContactCardIndex(pContactNumber)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ResizeControls()
    Dim vDCI As HeaderColumnInfo
    Dim vColumnLeft As Integer
    Dim vControl As Control
    Dim vTotalWidth As Integer

    'First work out the total width that is required
    vTotalWidth = START_LEFT
    For Each vDCI In mvColumns
      vTotalWidth += vDCI.RequiredColWidth
      vTotalWidth += COLUMN_GAP
      vDCI.ExcessWidth = 0
    Next

    'Now if more than one column calculate the excess and apportion proportionally against each column
    If vTotalWidth > pnl.Width And mvColumns.Count > 1 Then
      Dim vExcessWidth As Integer = vTotalWidth - pnl.Width
      For Each vDCI In mvColumns
        vDCI.ExcessWidth = (vDCI.RequiredColWidth * vExcessWidth) \ vTotalWidth
      Next
    End If

    'Set up where each column should go
    vColumnLeft = START_LEFT
    For Each vDCI In mvColumns
      With vDCI
        .LabelLeft = vColumnLeft
        .ItemLeft = vColumnLeft + .LabelWidth
        .GridLeft = vColumnLeft
        .UseItemWidth = .PreferredWidth - .ExcessWidth
        If .RequiredGridWidth > 0 Then
          .UseGridWidth = .RequiredGridWidth - .ExcessWidth
          If .UseGridWidth > .UseItemWidth Then .UseItemWidth = .UseGridWidth
          .UseGridWidth += .ItemLeft - .GridLeft
        Else
          .UseGridWidth = 0
        End If
        vColumnLeft += (.RequiredColWidth - .ExcessWidth + .ItemLeft - .GridLeft)
        vColumnLeft += COLUMN_GAP
      End With
    Next

    'Handle any excess real estate
    If mvColumns.Count = 1 Then
      vDCI = DirectCast(mvColumns(1), HeaderColumnInfo)
      vDCI.UseItemWidth = pnl.Width - (vDCI.LabelWidth + COLUMN_GAP)
      'vDCI.UseGridWidth = vDCI.UseItemWidth                              'Leave grids at the required width
    ElseIf pnl.Width > vColumnLeft AndAlso mvColumns.Count = 2 Then
      vDCI = DirectCast(mvColumns(2), HeaderColumnInfo)
      vDCI.UseItemWidth = vDCI.PreferredWidth + (pnl.Width - vColumnLeft)
      'vDCI.UseGridWidth = vDCI.UseItemWidth                              'Leave grids at the required width
    ElseIf pnl.Width > vColumnLeft AndAlso mvColumns.Count > 2 Then
      Dim vExtra As Integer = 0
      For Each vDCI In mvColumns
        With vDCI
          .LabelLeft += vExtra
          .ItemLeft += vExtra
        End With
        vExtra += (pnl.Width - vColumnLeft) \ mvColumns.Count
      Next
    End If
    'Now position the controls
    For Each vControl In pnl.Controls
      vDCI = DirectCast(vControl.Tag, HeaderColumnInfo)
      If TypeOf vControl Is LinkLabel Then
        vControl.SetBounds(vDCI.ItemLeft, vControl.Top, vDCI.UseItemWidth, vControl.Height)
      ElseIf TypeOf vControl Is Label Then
        vControl.SetBounds(vDCI.LabelLeft, vControl.Top, vDCI.LabelWidth - LABEL_GAP, vControl.Height)
      ElseIf TypeOf vControl Is TextBox Then
        vControl.SetBounds(vDCI.ItemLeft, vControl.Top, vDCI.UseItemWidth, vControl.Height)
      ElseIf TypeOf vControl Is DisplayGrid Then
        vControl.SetBounds(vDCI.LabelLeft, vControl.Top, vDCI.UseGridWidth, vControl.Height)
      ElseIf TypeOf vControl Is ChartControl Then
        vControl.SetBounds(vDCI.LabelLeft, vControl.Top, vDCI.UseItemWidth + vDCI.LabelWidth, vControl.Height)
      ElseIf TypeOf vControl Is PictureBox Then
        vControl.SetBounds(vDCI.LabelLeft, vControl.Top, vDCI.UseItemWidth + vDCI.LabelWidth, vControl.Height)
      ElseIf TypeOf vControl Is StandardPercentageBar Then
        vControl.SetBounds(vDCI.ItemLeft, vControl.Top, vDCI.UseItemWidth, vControl.Height)
      End If
    Next
    pnl.HorizontalScroll.Visible = False
  End Sub

  'Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
  '  Me.BackColor = System.Drawing.Color.Transparent
  '  MyBase.OnPaintBackground(e)
  '  mvBackgroundExtender.PaintBackGround(Me.ClientRectangle, e)
  'End Sub

  Private Class HeaderColumnInfo
    Friend LabelLeft As Integer
    Friend LabelWidth As Integer
    Friend ItemLeft As Integer
    Friend GridLeft As Integer
    Friend PreferredWidth As Integer            'Preferred width of text item
    Friend RequiredGridWidth As Integer
    Friend UseItemWidth As Integer
    Friend UseGridWidth As Integer
    Friend ExcessWidth As Integer

    Public ReadOnly Property RequiredColWidth() As Integer
      Get
        Return Math.Max(RequiredGridWidth, PreferredWidth + LabelWidth)
      End Get
    End Property

  End Class

  Private Sub cmdStickyNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStickyNote.Click
    EditContactData(CType(Me.ParentForm, MaintenanceParentForm), mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactStickyNotes, False)
  End Sub

  Private Sub cmdAction_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAction.Click
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactActions, mvContactInfo.ContactNumber)
    Dim vForm As New frmCardMaintenance(CType(Me.ParentForm, MaintenanceParentForm), mvContactInfo, CareServices.XMLContactDataSelectionTypes.xcdtContactActions, vDataSet, False, 0)
    vForm.Show()
  End Sub

  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub

  Private Sub pnl_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnl.MouseEnter
    RaiseEvent PanelMouseEnter(sender, e)
  End Sub

  Private Sub pnl_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnl.MouseLeave
    RaiseEvent PanelMouseLeave(sender, e)
  End Sub

  Public Sub SetControlTheme() Implements CDBNETCL.IThemeSettable.SetControlTheme
    If DisplayTheme.HeaderBackgroundSameAsForm Then
      Me.BackColor = DisplayTheme.FormBackColor
    Else
      Me.BackColor = mvBackColor
    End If
    For Each vControl As Control In pnl.Controls
      If TypeOf (vControl) Is TextBox OrElse TypeOf (vControl) Is LinkLabel Then vControl.BackColor = Me.BackColor
    Next
    DoSetControlTheme(pnl)
    Me.Invalidate()
  End Sub

  Private Sub cmsContactPic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmsContactPic.Click
    Try
      Dim vList As New ParameterList(True)
      vList("DocumentDefault") = AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_contact_image_document_type)
      Dim vDocDefaultRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtDocumentDefaults, vList)
      If Not vDocDefaultRow Is Nothing Then
        Dim vDocClass As String = vDocDefaultRow("DocumentClass").ToString()
        Dim vDocType As String = vDocDefaultRow("DocumentType").ToString()
        Dim vTopic As String = vDocDefaultRow("Topic").ToString()
        Dim vSubTopic As String = vDocDefaultRow("SubTopic").ToString()
        Dim vPackage As String = vDocDefaultRow("Package").ToString()
        'add the picture only if the defaults have been set up
        If Len(vDocClass) > 0 AndAlso Len(vDocType) > 0 AndAlso Len(vTopic) > 0 _
           AndAlso Len(vSubTopic) > 0 AndAlso Len(vPackage) > 0 Then
          Dim vOFD As New OpenFileDialog
          With vOFD
            .InitialDirectory = AppValues.DefaultImportDirectory
            .Title = "Import contact picture"
            .CheckFileExists = True
            .CheckPathExists = True
            .FileName = ""
            .Filter = "JPEG (*.jpg)|*.jpg"
            .DefaultExt = "*.jpg"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
              vList.Remove("DocumentDefault")
              vList("Package") = vPackage
              vList("Direction") = "I"
              vList("Dated") = AppValues.TodaysDate
              vList("SenderContactNumber") = mvContactInfo.ContactNumber.ToString()
              vList("SenderAddressNumber") = mvContactInfo.AddressNumber.ToString()
              vList("DocumentClass") = vDocClass
              vList("DocumentSubject") = "Picture"
              vList("DocumentType") = vDocType
              vList("Precis") = "Picture"
              vList("AddresseeContactNumber") = DataHelper.UserContactInfo.ContactNumber.ToString()
              vList("AddresseeAddressNumber") = DataHelper.UserContactInfo.AddressNumber.ToString()
              vList("SubTopic") = vSubTopic
              vList("Topic") = vTopic
              Dim vReturnList As ParameterList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctDocument, vList)
              Dim vDocNo As Integer = IntegerValue(vReturnList("DocumentNumber"))
              DataHelper.UpdateDocumentFile(vDocNo, .FileName)
              'Refresh the header
              Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, mvContactInfo.ContactNumber)
              Me.Populate(vDataSet, mvContactInfo)
            End If
          End With
        Else
          ShowInformationMessage(InformationMessages.ImDocumentDefaultsNotConfigured)
        End If
      Else
        ShowInformationMessage(InformationMessages.ImDocumentDefaultsNotConfigured)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmsContactPic_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsContactPic.Opening
    'Get the configured contact image document type
    Dim vDocType As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_contact_image_document_type)
    'If contact image document type is not configured then don't display the 'Add Picture' option
    If Not Len(vDocType) > 0 Then e.Cancel = True
  End Sub

  Private Sub pnlContextCustomise_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlContextCustomise.Click
    Dim vParams As New ParameterList(True)
    If mvContactInfo IsNot Nothing Then
      vParams.Add("SelectionPages", "Y")
      vParams("DataSelectionType") = "128"
      vParams("ParameterName") = mvContactInfo.ContactGroupParameterName
      vParams("ParameterValue") = mvContactInfo.ContactGroup
    ElseIf mvDataSet.Tables("DataRow").Rows.Count > 0 AndAlso mvDataSet.Tables("DataRow").Columns.Contains("EventNumber") Then
      vParams.Add("SelectionPages", "Y")
      vParams("DataSelectionType") = mvDataSet.Tables("Column").Rows(0).Item("Value").ToString()
      vParams("ParameterName") = ""
      vParams("ParameterValue") = ""
    End If
    Dim vDisplayList As New frmDisplayList(frmDisplayList.ListUsages.CustomiseDisplayList, vParams)
    If vDisplayList.ShowDialog() = DialogResult.OK Then
    End If
  End Sub

  Private Sub pnlContextRevert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlContextRevert.Click
    If ShowQuestion(QuestionMessages.QmRevertModule, MessageBoxButtons.OKCancel) = DialogResult.OK Then
      Dim vParams As New ParameterList(True)
      If mvContactInfo IsNot Nothing Then
        vParams("DataSelectionType") = mvDataSet.Tables("Column").Rows(0).Item("Value").ToString()
        vParams.Add(mvContactInfo.ContactGroupParameterName, mvContactInfo.ContactGroup)
      ElseIf mvDataSet.Tables("DataRow").Rows.Count > 0 AndAlso mvDataSet.Tables("DataRow").Columns.Contains("EventNumber") Then
        vParams("DataSelectionType") = mvDataSet.Tables("Column").Rows(0).Item("Value").ToString()
      End If
      vParams.Add("AccessMethod", "S")
      vParams.Add("Logname", DataHelper.UserInfo.Logname.ToString)
      vParams.Add("Department", DataHelper.UserInfo.Department.ToString)
      vParams.Add("Client", DataHelper.GetClientCode())
      vParams.Add("WebPageItemNumber", "")
      DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctDisplayListItem, vParams)
    End If
  End Sub
End Class

