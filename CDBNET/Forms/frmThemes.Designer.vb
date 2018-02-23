<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmThemes
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmThemes))
    Me.tab = New System.Windows.Forms.TabControl()
    Me.tabAppearance = New System.Windows.Forms.TabPage()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.cboSchemes = New System.Windows.Forms.ComboBox()
    Me.gbpFonts = New System.Windows.Forms.GroupBox()
    Me.lblFont = New System.Windows.Forms.Label()
    Me.cboFonts = New System.Windows.Forms.ComboBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.lblDisplayItem = New System.Windows.Forms.Label()
    Me.lblDisplayLabel = New System.Windows.Forms.Label()
    Me.lblNavPanel = New System.Windows.Forms.Label()
    Me.lblGridFont = New System.Windows.Forms.Label()
    Me.lblSelectionPanel = New System.Windows.Forms.Label()
    Me.lblFormFont = New System.Windows.Forms.Label()
    Me.gboAppearance = New System.Windows.Forms.GroupBox()
    Me.lblAppearance = New System.Windows.Forms.Label()
    Me.cboAppearance = New System.Windows.Forms.ComboBox()
    Me.lblButtonPanelBackColor = New System.Windows.Forms.Label()
    Me.lblSplitterBackColor = New System.Windows.Forms.Label()
    Me.csSplitterBackColor = New CDBNETCL.ColorSelector()
    Me.chkUnderlineHyperlinks = New System.Windows.Forms.CheckBox()
    Me.chkHeaderBackgroundSameAsForm = New System.Windows.Forms.CheckBox()
    Me.csButtonPanelBackColor = New CDBNETCL.ColorSelector()
    Me.lblGridHyperlinkColor = New System.Windows.Forms.Label()
    Me.lblPanelBackColor = New System.Windows.Forms.Label()
    Me.lblGridBackColor = New System.Windows.Forms.Label()
    Me.lblFormBackColor = New System.Windows.Forms.Label()
    Me.csFormBackColor = New CDBNETCL.ColorSelector()
    Me.csGridHyperlinkColor = New CDBNETCL.ColorSelector()
    Me.csPanelBackColor = New CDBNETCL.ColorSelector()
    Me.csGridBackColor = New CDBNETCL.ColorSelector()
    Me.tabPanels = New System.Windows.Forms.TabControl()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
    Me.TabPage2 = New System.Windows.Forms.TabPage()
    Me.TabPage3 = New System.Windows.Forms.TabPage()
    Me.TabPage4 = New System.Windows.Forms.TabPage()
    Me.TabPage5 = New System.Windows.Forms.TabPage()
    Me.TabPage6 = New System.Windows.Forms.TabPage()
    Me.TabPage7 = New System.Windows.Forms.TabPage()
    Me.cmd = New System.Windows.Forms.OpenFileDialog()
    Me.cdlg = New System.Windows.Forms.ColorDialog()
    Me.cmdOK = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.bpl = New CDBNETCL.ButtonPanel()
    Me.cmdSaveAs = New System.Windows.Forms.Button()
    Me.cmdApply = New System.Windows.Forms.Button()
    Me.cmdDefaults = New System.Windows.Forms.Button()
    Me.tim = New System.Windows.Forms.Timer(Me.components)
    Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
    Me.fsDashboardHeading = New CDBNET.FontSelector()
    Me.fsDisplayItem = New CDBNET.FontSelector()
    Me.fsDisplayLabel = New CDBNET.FontSelector()
    Me.fsNavigationPanel = New CDBNET.FontSelector()
    Me.fsSelectionPanel = New CDBNET.FontSelector()
    Me.fsGrid = New CDBNET.FontSelector()
    Me.fsForm = New CDBNET.FontSelector()
    Me.pteDisplayPanel = New CDBNET.PanelThemeEditor()
    Me.pteEditPanel = New CDBNET.PanelThemeEditor()
    Me.pteSelectionPanel = New CDBNET.PanelThemeEditor()
    Me.pteDisplayLabel = New CDBNET.PanelThemeEditor()
    Me.pteDisplayData = New CDBNET.PanelThemeEditor()
    Me.pteDashboardHeading = New CDBNET.PanelThemeEditor()
    Me.pteToolbar = New CDBNET.PanelThemeEditor()
    Me.tab.SuspendLayout()
    Me.tabAppearance.SuspendLayout()
    Me.gbpFonts.SuspendLayout()
    Me.gboAppearance.SuspendLayout()
    Me.tabPanels.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.TabPage2.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.TabPage4.SuspendLayout()
    Me.TabPage5.SuspendLayout()
    Me.TabPage6.SuspendLayout()
    Me.TabPage7.SuspendLayout()
    Me.bpl.SuspendLayout()
    CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'tab
    '
    Me.tab.Controls.Add(Me.tabAppearance)
    Me.tab.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tab.Location = New System.Drawing.Point(0, 0)
    Me.tab.Name = "tab"
    Me.tab.SelectedIndex = 0
    Me.tab.Size = New System.Drawing.Size(849, 559)
    Me.tab.TabIndex = 0
    '
    'tabAppearance
    '
    Me.tabAppearance.Controls.Add(Me.Label2)
    Me.tabAppearance.Controls.Add(Me.cboSchemes)
    Me.tabAppearance.Controls.Add(Me.gbpFonts)
    Me.tabAppearance.Controls.Add(Me.gboAppearance)
    Me.tabAppearance.Controls.Add(Me.tabPanels)
    Me.tabAppearance.Location = New System.Drawing.Point(4, 22)
    Me.tabAppearance.Name = "tabAppearance"
    Me.tabAppearance.Padding = New System.Windows.Forms.Padding(3)
    Me.tabAppearance.Size = New System.Drawing.Size(841, 533)
    Me.tabAppearance.TabIndex = 3
    Me.tabAppearance.Text = "Appearance"
    Me.tabAppearance.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(16, 9)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(79, 13)
    Me.Label2.TabIndex = 20
    Me.Label2.Text = "Select Scheme"
    '
    'cboSchemes
    '
    Me.cboSchemes.FormattingEnabled = True
    Me.cboSchemes.Location = New System.Drawing.Point(105, 6)
    Me.cboSchemes.Name = "cboSchemes"
    Me.cboSchemes.Size = New System.Drawing.Size(252, 21)
    Me.cboSchemes.TabIndex = 19
    '
    'gbpFonts
    '
    Me.gbpFonts.Controls.Add(Me.lblFont)
    Me.gbpFonts.Controls.Add(Me.cboFonts)
    Me.gbpFonts.Controls.Add(Me.Label1)
    Me.gbpFonts.Controls.Add(Me.lblDisplayItem)
    Me.gbpFonts.Controls.Add(Me.lblDisplayLabel)
    Me.gbpFonts.Controls.Add(Me.lblNavPanel)
    Me.gbpFonts.Controls.Add(Me.lblGridFont)
    Me.gbpFonts.Controls.Add(Me.lblSelectionPanel)
    Me.gbpFonts.Controls.Add(Me.lblFormFont)
    Me.gbpFonts.Controls.Add(Me.fsDashboardHeading)
    Me.gbpFonts.Controls.Add(Me.fsDisplayItem)
    Me.gbpFonts.Controls.Add(Me.fsDisplayLabel)
    Me.gbpFonts.Controls.Add(Me.fsNavigationPanel)
    Me.gbpFonts.Controls.Add(Me.fsSelectionPanel)
    Me.gbpFonts.Controls.Add(Me.fsGrid)
    Me.gbpFonts.Controls.Add(Me.fsForm)
    Me.gbpFonts.Location = New System.Drawing.Point(400, 46)
    Me.gbpFonts.Name = "gbpFonts"
    Me.gbpFonts.Size = New System.Drawing.Size(435, 312)
    Me.gbpFonts.TabIndex = 18
    Me.gbpFonts.TabStop = False
    Me.gbpFonts.Text = "Fonts"
    '
    'lblFont
    '
    Me.lblFont.AutoSize = True
    Me.lblFont.Location = New System.Drawing.Point(3, 23)
    Me.lblFont.Name = "lblFont"
    Me.lblFont.Size = New System.Drawing.Size(28, 13)
    Me.lblFont.TabIndex = 29
    Me.lblFont.Text = "Font"
    '
    'cboFonts
    '
    Me.cboFonts.FormattingEnabled = True
    Me.cboFonts.Location = New System.Drawing.Point(92, 20)
    Me.cboFonts.Name = "cboFonts"
    Me.cboFonts.Size = New System.Drawing.Size(252, 21)
    Me.cboFonts.TabIndex = 28
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(0, 269)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(102, 13)
    Me.Label1.TabIndex = 27
    Me.Label1.Text = "Dashboard Heading"
    '
    'lblDisplayItem
    '
    Me.lblDisplayItem.AutoSize = True
    Me.lblDisplayItem.Location = New System.Drawing.Point(0, 235)
    Me.lblDisplayItem.Name = "lblDisplayItem"
    Me.lblDisplayItem.Size = New System.Drawing.Size(64, 13)
    Me.lblDisplayItem.TabIndex = 24
    Me.lblDisplayItem.Text = "Display Item"
    '
    'lblDisplayLabel
    '
    Me.lblDisplayLabel.AutoSize = True
    Me.lblDisplayLabel.Location = New System.Drawing.Point(0, 200)
    Me.lblDisplayLabel.Name = "lblDisplayLabel"
    Me.lblDisplayLabel.Size = New System.Drawing.Size(70, 13)
    Me.lblDisplayLabel.TabIndex = 22
    Me.lblDisplayLabel.Text = "Display Label"
    '
    'lblNavPanel
    '
    Me.lblNavPanel.AutoSize = True
    Me.lblNavPanel.Location = New System.Drawing.Point(0, 165)
    Me.lblNavPanel.Name = "lblNavPanel"
    Me.lblNavPanel.Size = New System.Drawing.Size(88, 13)
    Me.lblNavPanel.TabIndex = 20
    Me.lblNavPanel.Text = "Navigation Panel"
    '
    'lblGridFont
    '
    Me.lblGridFont.AutoSize = True
    Me.lblGridFont.Location = New System.Drawing.Point(0, 91)
    Me.lblGridFont.Name = "lblGridFont"
    Me.lblGridFont.Size = New System.Drawing.Size(31, 13)
    Me.lblGridFont.TabIndex = 16
    Me.lblGridFont.Text = "Grids"
    '
    'lblSelectionPanel
    '
    Me.lblSelectionPanel.AutoSize = True
    Me.lblSelectionPanel.Location = New System.Drawing.Point(0, 128)
    Me.lblSelectionPanel.Name = "lblSelectionPanel"
    Me.lblSelectionPanel.Size = New System.Drawing.Size(81, 13)
    Me.lblSelectionPanel.TabIndex = 18
    Me.lblSelectionPanel.Text = "Selection Panel"
    '
    'lblFormFont
    '
    Me.lblFormFont.AutoSize = True
    Me.lblFormFont.Location = New System.Drawing.Point(0, 53)
    Me.lblFormFont.Name = "lblFormFont"
    Me.lblFormFont.Size = New System.Drawing.Size(35, 13)
    Me.lblFormFont.TabIndex = 14
    Me.lblFormFont.Text = "Forms"
    '
    'gboAppearance
    '
    Me.gboAppearance.Controls.Add(Me.lblAppearance)
    Me.gboAppearance.Controls.Add(Me.cboAppearance)
    Me.gboAppearance.Controls.Add(Me.lblButtonPanelBackColor)
    Me.gboAppearance.Controls.Add(Me.lblSplitterBackColor)
    Me.gboAppearance.Controls.Add(Me.csSplitterBackColor)
    Me.gboAppearance.Controls.Add(Me.chkUnderlineHyperlinks)
    Me.gboAppearance.Controls.Add(Me.chkHeaderBackgroundSameAsForm)
    Me.gboAppearance.Controls.Add(Me.csButtonPanelBackColor)
    Me.gboAppearance.Controls.Add(Me.lblGridHyperlinkColor)
    Me.gboAppearance.Controls.Add(Me.lblPanelBackColor)
    Me.gboAppearance.Controls.Add(Me.lblGridBackColor)
    Me.gboAppearance.Controls.Add(Me.lblFormBackColor)
    Me.gboAppearance.Controls.Add(Me.csFormBackColor)
    Me.gboAppearance.Controls.Add(Me.csGridHyperlinkColor)
    Me.gboAppearance.Controls.Add(Me.csPanelBackColor)
    Me.gboAppearance.Controls.Add(Me.csGridBackColor)
    Me.gboAppearance.Location = New System.Drawing.Point(9, 46)
    Me.gboAppearance.Name = "gboAppearance"
    Me.gboAppearance.Size = New System.Drawing.Size(380, 313)
    Me.gboAppearance.TabIndex = 17
    Me.gboAppearance.TabStop = False
    Me.gboAppearance.Text = "Appearance"
    '
    'lblAppearance
    '
    Me.lblAppearance.AutoSize = True
    Me.lblAppearance.Location = New System.Drawing.Point(7, 22)
    Me.lblAppearance.Name = "lblAppearance"
    Me.lblAppearance.Size = New System.Drawing.Size(65, 13)
    Me.lblAppearance.TabIndex = 32
    Me.lblAppearance.Text = "Appearance"
    '
    'cboAppearance
    '
    Me.cboAppearance.FormattingEnabled = True
    Me.cboAppearance.Location = New System.Drawing.Point(96, 19)
    Me.cboAppearance.Name = "cboAppearance"
    Me.cboAppearance.Size = New System.Drawing.Size(252, 21)
    Me.cboAppearance.TabIndex = 31
    '
    'lblButtonPanelBackColor
    '
    Me.lblButtonPanelBackColor.AutoSize = True
    Me.lblButtonPanelBackColor.Location = New System.Drawing.Point(7, 215)
    Me.lblButtonPanelBackColor.Name = "lblButtonPanelBackColor"
    Me.lblButtonPanelBackColor.Size = New System.Drawing.Size(120, 13)
    Me.lblButtonPanelBackColor.TabIndex = 30
    Me.lblButtonPanelBackColor.Text = "Button Panel BackColor"
    '
    'lblSplitterBackColor
    '
    Me.lblSplitterBackColor.AutoSize = True
    Me.lblSplitterBackColor.Location = New System.Drawing.Point(7, 182)
    Me.lblSplitterBackColor.Name = "lblSplitterBackColor"
    Me.lblSplitterBackColor.Size = New System.Drawing.Size(100, 13)
    Me.lblSplitterBackColor.TabIndex = 29
    Me.lblSplitterBackColor.Text = "Splitter Background"
    '
    'csSplitterBackColor
    '
    Me.csSplitterBackColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csSplitterBackColor.ColorDialog = Nothing
    Me.csSplitterBackColor.Location = New System.Drawing.Point(168, 182)
    Me.csSplitterBackColor.Margin = New System.Windows.Forms.Padding(2)
    Me.csSplitterBackColor.Name = "csSplitterBackColor"
    Me.csSplitterBackColor.RGBValue = "16777215"
    Me.csSplitterBackColor.Size = New System.Drawing.Size(203, 31)
    Me.csSplitterBackColor.SupportsTransparentColor = False
    Me.csSplitterBackColor.TabIndex = 25
    '
    'chkUnderlineHyperlinks
    '
    Me.chkUnderlineHyperlinks.AutoSize = True
    Me.chkUnderlineHyperlinks.Location = New System.Drawing.Point(7, 281)
    Me.chkUnderlineHyperlinks.Name = "chkUnderlineHyperlinks"
    Me.chkUnderlineHyperlinks.Size = New System.Drawing.Size(121, 17)
    Me.chkUnderlineHyperlinks.TabIndex = 28
    Me.chkUnderlineHyperlinks.Text = "Underline Grid Links"
    Me.chkUnderlineHyperlinks.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.chkUnderlineHyperlinks.UseVisualStyleBackColor = True
    '
    'chkHeaderBackgroundSameAsForm
    '
    Me.chkHeaderBackgroundSameAsForm.AutoSize = True
    Me.chkHeaderBackgroundSameAsForm.Location = New System.Drawing.Point(7, 251)
    Me.chkHeaderBackgroundSameAsForm.Name = "chkHeaderBackgroundSameAsForm"
    Me.chkHeaderBackgroundSameAsForm.Size = New System.Drawing.Size(187, 17)
    Me.chkHeaderBackgroundSameAsForm.TabIndex = 27
    Me.chkHeaderBackgroundSameAsForm.Text = "Header Background same as form"
    Me.chkHeaderBackgroundSameAsForm.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.chkHeaderBackgroundSameAsForm.UseVisualStyleBackColor = True
    '
    'csButtonPanelBackColor
    '
    Me.csButtonPanelBackColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csButtonPanelBackColor.ColorDialog = Nothing
    Me.csButtonPanelBackColor.Location = New System.Drawing.Point(168, 215)
    Me.csButtonPanelBackColor.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.csButtonPanelBackColor.Name = "csButtonPanelBackColor"
    Me.csButtonPanelBackColor.RGBValue = "16777215"
    Me.csButtonPanelBackColor.Size = New System.Drawing.Size(203, 31)
    Me.csButtonPanelBackColor.SupportsTransparentColor = False
    Me.csButtonPanelBackColor.TabIndex = 26
    '
    'lblGridHyperlinkColor
    '
    Me.lblGridHyperlinkColor.AutoSize = True
    Me.lblGridHyperlinkColor.Location = New System.Drawing.Point(7, 152)
    Me.lblGridHyperlinkColor.Name = "lblGridHyperlinkColor"
    Me.lblGridHyperlinkColor.Size = New System.Drawing.Size(100, 13)
    Me.lblGridHyperlinkColor.TabIndex = 24
    Me.lblGridHyperlinkColor.Text = "Grid Hyperlink Color"
    '
    'lblPanelBackColor
    '
    Me.lblPanelBackColor.AutoSize = True
    Me.lblPanelBackColor.Location = New System.Drawing.Point(7, 119)
    Me.lblPanelBackColor.Name = "lblPanelBackColor"
    Me.lblPanelBackColor.Size = New System.Drawing.Size(142, 13)
    Me.lblPanelBackColor.TabIndex = 23
    Me.lblPanelBackColor.Text = "Selection Panel Background"
    '
    'lblGridBackColor
    '
    Me.lblGridBackColor.AutoSize = True
    Me.lblGridBackColor.Location = New System.Drawing.Point(7, 86)
    Me.lblGridBackColor.Name = "lblGridBackColor"
    Me.lblGridBackColor.Size = New System.Drawing.Size(87, 13)
    Me.lblGridBackColor.TabIndex = 22
    Me.lblGridBackColor.Text = "Grid Background"
    '
    'lblFormBackColor
    '
    Me.lblFormBackColor.AutoSize = True
    Me.lblFormBackColor.Location = New System.Drawing.Point(7, 53)
    Me.lblFormBackColor.Name = "lblFormBackColor"
    Me.lblFormBackColor.Size = New System.Drawing.Size(91, 13)
    Me.lblFormBackColor.TabIndex = 21
    Me.lblFormBackColor.Text = "Form Background"
    '
    'csFormBackColor
    '
    Me.csFormBackColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csFormBackColor.ColorDialog = Nothing
    Me.csFormBackColor.Location = New System.Drawing.Point(168, 53)
    Me.csFormBackColor.Margin = New System.Windows.Forms.Padding(2)
    Me.csFormBackColor.Name = "csFormBackColor"
    Me.csFormBackColor.RGBValue = "16777215"
    Me.csFormBackColor.Size = New System.Drawing.Size(203, 31)
    Me.csFormBackColor.SupportsTransparentColor = False
    Me.csFormBackColor.TabIndex = 17
    '
    'csGridHyperlinkColor
    '
    Me.csGridHyperlinkColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csGridHyperlinkColor.ColorDialog = Nothing
    Me.csGridHyperlinkColor.Location = New System.Drawing.Point(168, 152)
    Me.csGridHyperlinkColor.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.csGridHyperlinkColor.Name = "csGridHyperlinkColor"
    Me.csGridHyperlinkColor.RGBValue = "16777215"
    Me.csGridHyperlinkColor.Size = New System.Drawing.Size(203, 26)
    Me.csGridHyperlinkColor.SupportsTransparentColor = False
    Me.csGridHyperlinkColor.TabIndex = 20
    '
    'csPanelBackColor
    '
    Me.csPanelBackColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csPanelBackColor.ColorDialog = Nothing
    Me.csPanelBackColor.Location = New System.Drawing.Point(168, 119)
    Me.csPanelBackColor.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.csPanelBackColor.Name = "csPanelBackColor"
    Me.csPanelBackColor.RGBValue = "16777215"
    Me.csPanelBackColor.Size = New System.Drawing.Size(203, 31)
    Me.csPanelBackColor.SupportsTransparentColor = False
    Me.csPanelBackColor.TabIndex = 19
    '
    'csGridBackColor
    '
    Me.csGridBackColor.Color = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
    Me.csGridBackColor.ColorDialog = Nothing
    Me.csGridBackColor.Location = New System.Drawing.Point(168, 86)
    Me.csGridBackColor.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.csGridBackColor.Name = "csGridBackColor"
    Me.csGridBackColor.RGBValue = "16777215"
    Me.csGridBackColor.Size = New System.Drawing.Size(203, 31)
    Me.csGridBackColor.SupportsTransparentColor = False
    Me.csGridBackColor.TabIndex = 18
    '
    'tabPanels
    '
    Me.tabPanels.Controls.Add(Me.TabPage1)
    Me.tabPanels.Controls.Add(Me.TabPage2)
    Me.tabPanels.Controls.Add(Me.TabPage3)
    Me.tabPanels.Controls.Add(Me.TabPage4)
    Me.tabPanels.Controls.Add(Me.TabPage5)
    Me.tabPanels.Controls.Add(Me.TabPage6)
    Me.tabPanels.Controls.Add(Me.TabPage7)
    Me.tabPanels.Location = New System.Drawing.Point(9, 365)
    Me.tabPanels.Name = "tabPanels"
    Me.tabPanels.SelectedIndex = 0
    Me.tabPanels.Size = New System.Drawing.Size(825, 151)
    Me.tabPanels.TabIndex = 7
    '
    'TabPage1
    '
    Me.TabPage1.Controls.Add(Me.pteDisplayPanel)
    Me.TabPage1.Location = New System.Drawing.Point(4, 22)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(817, 125)
    Me.TabPage1.TabIndex = 0
    Me.TabPage1.Text = "Display Panel"
    Me.TabPage1.UseVisualStyleBackColor = True
    '
    'TabPage2
    '
    Me.TabPage2.Controls.Add(Me.pteEditPanel)
    Me.TabPage2.Location = New System.Drawing.Point(4, 22)
    Me.TabPage2.Name = "TabPage2"
    Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage2.Size = New System.Drawing.Size(817, 125)
    Me.TabPage2.TabIndex = 1
    Me.TabPage2.Text = "Edit Panel"
    Me.TabPage2.UseVisualStyleBackColor = True
    '
    'TabPage3
    '
    Me.TabPage3.Controls.Add(Me.pteSelectionPanel)
    Me.TabPage3.Location = New System.Drawing.Point(4, 22)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage3.Size = New System.Drawing.Size(817, 125)
    Me.TabPage3.TabIndex = 2
    Me.TabPage3.Text = "Selection Panel"
    Me.TabPage3.UseVisualStyleBackColor = True
    '
    'TabPage4
    '
    Me.TabPage4.Controls.Add(Me.pteDisplayLabel)
    Me.TabPage4.Location = New System.Drawing.Point(4, 22)
    Me.TabPage4.Name = "TabPage4"
    Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage4.Size = New System.Drawing.Size(817, 125)
    Me.TabPage4.TabIndex = 3
    Me.TabPage4.Text = "Display Label"
    Me.TabPage4.UseVisualStyleBackColor = True
    '
    'TabPage5
    '
    Me.TabPage5.Controls.Add(Me.pteDisplayData)
    Me.TabPage5.Location = New System.Drawing.Point(4, 22)
    Me.TabPage5.Name = "TabPage5"
    Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage5.Size = New System.Drawing.Size(817, 125)
    Me.TabPage5.TabIndex = 4
    Me.TabPage5.Text = "Display Item"
    Me.TabPage5.UseVisualStyleBackColor = True
    '
    'TabPage6
    '
    Me.TabPage6.Controls.Add(Me.pteDashboardHeading)
    Me.TabPage6.Location = New System.Drawing.Point(4, 22)
    Me.TabPage6.Name = "TabPage6"
    Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage6.Size = New System.Drawing.Size(817, 125)
    Me.TabPage6.TabIndex = 5
    Me.TabPage6.Text = "Dashboard Heading"
    Me.TabPage6.UseVisualStyleBackColor = True
    '
    'TabPage7
    '
    Me.TabPage7.Controls.Add(Me.pteToolbar)
    Me.TabPage7.Location = New System.Drawing.Point(4, 22)
    Me.TabPage7.Name = "TabPage7"
    Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage7.Size = New System.Drawing.Size(817, 125)
    Me.TabPage7.TabIndex = 6
    Me.TabPage7.Text = "Toolbar"
    Me.TabPage7.UseVisualStyleBackColor = True
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(154, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(96, 27)
    Me.cmdOK.TabIndex = 0
    Me.cmdOK.Text = "OK"
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(598, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 4
    Me.cmdCancel.Text = "Cancel"
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdOK)
    Me.bpl.Controls.Add(Me.cmdSaveAs)
    Me.bpl.Controls.Add(Me.cmdApply)
    Me.bpl.Controls.Add(Me.cmdDefaults)
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 559)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(849, 39)
    Me.bpl.TabIndex = 2
    '
    'cmdSaveAs
    '
    Me.cmdSaveAs.Location = New System.Drawing.Point(265, 6)
    Me.cmdSaveAs.Name = "cmdSaveAs"
    Me.cmdSaveAs.Size = New System.Drawing.Size(96, 27)
    Me.cmdSaveAs.TabIndex = 1
    Me.cmdSaveAs.Text = "&Save"
    '
    'cmdApply
    '
    Me.cmdApply.Location = New System.Drawing.Point(376, 6)
    Me.cmdApply.Name = "cmdApply"
    Me.cmdApply.Size = New System.Drawing.Size(96, 27)
    Me.cmdApply.TabIndex = 2
    Me.cmdApply.Text = "&Apply"
    '
    'cmdDefaults
    '
    Me.cmdDefaults.Location = New System.Drawing.Point(487, 6)
    Me.cmdDefaults.Name = "cmdDefaults"
    Me.cmdDefaults.Size = New System.Drawing.Size(96, 27)
    Me.cmdDefaults.TabIndex = 3
    Me.cmdDefaults.Text = "Defaults"
    '
    'tim
    '
    '
    'ErrorProvider
    '
    Me.ErrorProvider.ContainerControl = Me
    '
    'fsDashboardHeading
    '
    Me.fsDashboardHeading.Location = New System.Drawing.Point(146, 262)
    Me.fsDashboardHeading.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsDashboardHeading.Name = "fsDashboardHeading"
    Me.fsDashboardHeading.SelectedFont = Nothing
    Me.fsDashboardHeading.Size = New System.Drawing.Size(288, 31)
    Me.fsDashboardHeading.TabIndex = 26
    '
    'fsDisplayItem
    '
    Me.fsDisplayItem.Location = New System.Drawing.Point(146, 227)
    Me.fsDisplayItem.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsDisplayItem.Name = "fsDisplayItem"
    Me.fsDisplayItem.SelectedFont = Nothing
    Me.fsDisplayItem.Size = New System.Drawing.Size(288, 31)
    Me.fsDisplayItem.TabIndex = 25
    '
    'fsDisplayLabel
    '
    Me.fsDisplayLabel.Location = New System.Drawing.Point(146, 192)
    Me.fsDisplayLabel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsDisplayLabel.Name = "fsDisplayLabel"
    Me.fsDisplayLabel.SelectedFont = Nothing
    Me.fsDisplayLabel.Size = New System.Drawing.Size(288, 31)
    Me.fsDisplayLabel.TabIndex = 23
    '
    'fsNavigationPanel
    '
    Me.fsNavigationPanel.Location = New System.Drawing.Point(146, 157)
    Me.fsNavigationPanel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsNavigationPanel.Name = "fsNavigationPanel"
    Me.fsNavigationPanel.SelectedFont = Nothing
    Me.fsNavigationPanel.Size = New System.Drawing.Size(288, 31)
    Me.fsNavigationPanel.TabIndex = 21
    '
    'fsSelectionPanel
    '
    Me.fsSelectionPanel.Location = New System.Drawing.Point(146, 120)
    Me.fsSelectionPanel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsSelectionPanel.Name = "fsSelectionPanel"
    Me.fsSelectionPanel.SelectedFont = Nothing
    Me.fsSelectionPanel.Size = New System.Drawing.Size(288, 31)
    Me.fsSelectionPanel.TabIndex = 19
    '
    'fsGrid
    '
    Me.fsGrid.Location = New System.Drawing.Point(146, 83)
    Me.fsGrid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsGrid.Name = "fsGrid"
    Me.fsGrid.SelectedFont = Nothing
    Me.fsGrid.Size = New System.Drawing.Size(288, 31)
    Me.fsGrid.TabIndex = 17
    '
    'fsForm
    '
    Me.fsForm.Location = New System.Drawing.Point(146, 46)
    Me.fsForm.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.fsForm.Name = "fsForm"
    Me.fsForm.SelectedFont = Nothing
    Me.fsForm.Size = New System.Drawing.Size(288, 31)
    Me.fsForm.TabIndex = 15
    '
    'pteDisplayPanel
    '
    Me.pteDisplayPanel.ColorDialog = Nothing
    Me.pteDisplayPanel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteDisplayPanel.Location = New System.Drawing.Point(3, 3)
    Me.pteDisplayPanel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteDisplayPanel.Name = "pteDisplayPanel"
    Me.pteDisplayPanel.Size = New System.Drawing.Size(811, 119)
    Me.pteDisplayPanel.TabIndex = 0
    Me.pteDisplayPanel.TextBox = False
    '
    'pteEditPanel
    '
    Me.pteEditPanel.ColorDialog = Nothing
    Me.pteEditPanel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteEditPanel.Location = New System.Drawing.Point(3, 3)
    Me.pteEditPanel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteEditPanel.Name = "pteEditPanel"
    Me.pteEditPanel.Size = New System.Drawing.Size(811, 119)
    Me.pteEditPanel.TabIndex = 0
    Me.pteEditPanel.TextBox = False
    '
    'pteSelectionPanel
    '
    Me.pteSelectionPanel.ColorDialog = Nothing
    Me.pteSelectionPanel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteSelectionPanel.Location = New System.Drawing.Point(3, 3)
    Me.pteSelectionPanel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteSelectionPanel.Name = "pteSelectionPanel"
    Me.pteSelectionPanel.Size = New System.Drawing.Size(811, 119)
    Me.pteSelectionPanel.TabIndex = 0
    Me.pteSelectionPanel.TextBox = False
    '
    'pteDisplayLabel
    '
    Me.pteDisplayLabel.ColorDialog = Nothing
    Me.pteDisplayLabel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteDisplayLabel.Location = New System.Drawing.Point(3, 3)
    Me.pteDisplayLabel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteDisplayLabel.Name = "pteDisplayLabel"
    Me.pteDisplayLabel.Size = New System.Drawing.Size(811, 119)
    Me.pteDisplayLabel.TabIndex = 0
    Me.pteDisplayLabel.TextBox = False
    '
    'pteDisplayData
    '
    Me.pteDisplayData.ColorDialog = Nothing
    Me.pteDisplayData.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteDisplayData.Location = New System.Drawing.Point(3, 3)
    Me.pteDisplayData.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteDisplayData.Name = "pteDisplayData"
    Me.pteDisplayData.Size = New System.Drawing.Size(811, 119)
    Me.pteDisplayData.TabIndex = 1
    Me.pteDisplayData.TextBox = True
    '
    'pteDashboardHeading
    '
    Me.pteDashboardHeading.ColorDialog = Nothing
    Me.pteDashboardHeading.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteDashboardHeading.Location = New System.Drawing.Point(3, 3)
    Me.pteDashboardHeading.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteDashboardHeading.Name = "pteDashboardHeading"
    Me.pteDashboardHeading.Size = New System.Drawing.Size(811, 119)
    Me.pteDashboardHeading.TabIndex = 1
    Me.pteDashboardHeading.TextBox = False
    '
    'pteToolbar
    '
    Me.pteToolbar.ColorDialog = Nothing
    Me.pteToolbar.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pteToolbar.Location = New System.Drawing.Point(3, 3)
    Me.pteToolbar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
    Me.pteToolbar.Name = "pteToolbar"
    Me.pteToolbar.Size = New System.Drawing.Size(811, 119)
    Me.pteToolbar.TabIndex = 2
    Me.pteToolbar.TextBox = False
    '
    'frmThemes
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(849, 598)
    Me.Controls.Add(Me.tab)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmThemes"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = ""
    Me.tab.ResumeLayout(False)
    Me.tabAppearance.ResumeLayout(False)
    Me.tabAppearance.PerformLayout()
    Me.gbpFonts.ResumeLayout(False)
    Me.gbpFonts.PerformLayout()
    Me.gboAppearance.ResumeLayout(False)
    Me.gboAppearance.PerformLayout()
    Me.tabPanels.ResumeLayout(False)
    Me.TabPage1.ResumeLayout(False)
    Me.TabPage2.ResumeLayout(False)
    Me.TabPage3.ResumeLayout(False)
    Me.TabPage4.ResumeLayout(False)
    Me.TabPage5.ResumeLayout(False)
    Me.TabPage6.ResumeLayout(False)
    Me.TabPage7.ResumeLayout(False)
    Me.bpl.ResumeLayout(False)
    CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdDefaults As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tab As System.Windows.Forms.TabControl
  Friend WithEvents cmd As System.Windows.Forms.OpenFileDialog
  Friend WithEvents cdlg As System.Windows.Forms.ColorDialog
  Friend WithEvents cmdApply As System.Windows.Forms.Button
  Friend WithEvents cmdSaveAs As System.Windows.Forms.Button
  Friend WithEvents tim As System.Windows.Forms.Timer
  Friend WithEvents tabAppearance As System.Windows.Forms.TabPage
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents cboSchemes As System.Windows.Forms.ComboBox
  Friend WithEvents gbpFonts As System.Windows.Forms.GroupBox
  Friend WithEvents lblFont As System.Windows.Forms.Label
  Friend WithEvents cboFonts As System.Windows.Forms.ComboBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents lblDisplayItem As System.Windows.Forms.Label
  Friend WithEvents lblDisplayLabel As System.Windows.Forms.Label
  Friend WithEvents lblNavPanel As System.Windows.Forms.Label
  Friend WithEvents lblGridFont As System.Windows.Forms.Label
  Friend WithEvents lblSelectionPanel As System.Windows.Forms.Label
  Friend WithEvents lblFormFont As System.Windows.Forms.Label
  Friend WithEvents fsDashboardHeading As CDBNET.FontSelector
  Friend WithEvents fsDisplayItem As CDBNET.FontSelector
  Friend WithEvents fsDisplayLabel As CDBNET.FontSelector
  Friend WithEvents fsNavigationPanel As CDBNET.FontSelector
  Friend WithEvents fsSelectionPanel As CDBNET.FontSelector
  Friend WithEvents fsGrid As CDBNET.FontSelector
  Friend WithEvents fsForm As CDBNET.FontSelector
  Friend WithEvents gboAppearance As System.Windows.Forms.GroupBox
  Friend WithEvents lblAppearance As System.Windows.Forms.Label
  Friend WithEvents cboAppearance As System.Windows.Forms.ComboBox
  Friend WithEvents lblButtonPanelBackColor As System.Windows.Forms.Label
  Friend WithEvents lblSplitterBackColor As System.Windows.Forms.Label
  Friend WithEvents csSplitterBackColor As CDBNETCL.ColorSelector
  Friend WithEvents chkUnderlineHyperlinks As System.Windows.Forms.CheckBox
  Friend WithEvents chkHeaderBackgroundSameAsForm As System.Windows.Forms.CheckBox
  Friend WithEvents csButtonPanelBackColor As CDBNETCL.ColorSelector
  Friend WithEvents lblGridHyperlinkColor As System.Windows.Forms.Label
  Friend WithEvents lblPanelBackColor As System.Windows.Forms.Label
  Friend WithEvents lblGridBackColor As System.Windows.Forms.Label
  Friend WithEvents lblFormBackColor As System.Windows.Forms.Label
  Friend WithEvents csFormBackColor As CDBNETCL.ColorSelector
  Friend WithEvents csGridHyperlinkColor As CDBNETCL.ColorSelector
  Friend WithEvents csPanelBackColor As CDBNETCL.ColorSelector
  Friend WithEvents csGridBackColor As CDBNETCL.ColorSelector
  Friend WithEvents tabPanels As System.Windows.Forms.TabControl
  Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
  Friend WithEvents pteDisplayPanel As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
  Friend WithEvents pteEditPanel As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
  Friend WithEvents pteSelectionPanel As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
  Friend WithEvents pteDisplayLabel As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
  Friend WithEvents pteDisplayData As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
  Friend WithEvents pteDashboardHeading As CDBNET.PanelThemeEditor
  Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
  Friend WithEvents pteToolbar As CDBNET.PanelThemeEditor
  Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
End Class
