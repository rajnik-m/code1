Public Class PanelThemeEditor

  Private mvTextBox As Boolean

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  <System.ComponentModel.Browsable(True)> _
   <System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)> _
   Public Property TextBox() As Boolean
    Get
      Return mvTextBox
    End Get
    Set(ByVal pValue As Boolean)
      mvTextBox = pValue
      SetControlsVisible()
    End Set
  End Property

  Public Property ColorDialog() As ColorDialog
    Get
      Return cs1.ColorDialog
    End Get
    Set(ByVal value As ColorDialog)
      cs1.ColorDialog = value
      cs2.ColorDialog = value
    End Set
  End Property

  Private Sub SetControlsVisible()
    cboCurvature.Visible = Not mvTextBox
    lblCurvature.Visible = Not mvTextBox
    cboCurveMode.Visible = Not mvTextBox
    lblCurveMode.Visible = Not mvTextBox
    cboGradientMode.Visible = Not mvTextBox
    lblGradientMode.Visible = Not mvTextBox
    cs2.Visible = Not mvTextBox
    lblColor2.Visible = Not mvTextBox
  End Sub

  Private Sub InitialiseControls()
    Dim vGM() As LookupItem = { _
    New LookupItem(CStr(BackGroundExtender.LinearGradientMode.None), [Enum].GetName(GetType(BackGroundExtender.LinearGradientMode), BackGroundExtender.LinearGradientMode.None)), _
    New LookupItem(CStr(BackGroundExtender.LinearGradientMode.Horizontal), [Enum].GetName(GetType(BackGroundExtender.LinearGradientMode), BackGroundExtender.LinearGradientMode.Horizontal)), _
    New LookupItem(CStr(BackGroundExtender.LinearGradientMode.Vertical), [Enum].GetName(GetType(BackGroundExtender.LinearGradientMode), BackGroundExtender.LinearGradientMode.Vertical)), _
    New LookupItem(CStr(BackGroundExtender.LinearGradientMode.ForwardDiagonal), [Enum].GetName(GetType(BackGroundExtender.LinearGradientMode), BackGroundExtender.LinearGradientMode.ForwardDiagonal)), _
    New LookupItem(CStr(BackGroundExtender.LinearGradientMode.BackwardDiagonal), [Enum].GetName(GetType(BackGroundExtender.LinearGradientMode), BackGroundExtender.LinearGradientMode.BackwardDiagonal))}
    cboGradientMode.DisplayMember = "LookupDesc"
    cboGradientMode.ValueMember = "LookupCode"
    cboGradientMode.DataSource = vGM

    Dim vCM() As LookupItem = { _
    New LookupItem(CStr(BackGroundExtender.CornerCurveMode.None), [Enum].GetName(GetType(BackGroundExtender.CornerCurveMode), BackGroundExtender.CornerCurveMode.None)), _
    New LookupItem(CStr(BackGroundExtender.CornerCurveMode.TopLeft_TopRight), [Enum].GetName(GetType(BackGroundExtender.CornerCurveMode), BackGroundExtender.CornerCurveMode.TopLeft_TopRight)), _
    New LookupItem(CStr(BackGroundExtender.CornerCurveMode.BottomRight_BottomLeft), [Enum].GetName(GetType(BackGroundExtender.CornerCurveMode), BackGroundExtender.CornerCurveMode.BottomRight_BottomLeft)), _
    New LookupItem(CStr(BackGroundExtender.CornerCurveMode.All), [Enum].GetName(GetType(BackGroundExtender.CornerCurveMode), BackGroundExtender.CornerCurveMode.All))}
    cboCurveMode.DisplayMember = "LookupDesc"
    cboCurveMode.ValueMember = "LookupCode"
    cboCurveMode.DataSource = vCM

    Dim vBS() As LookupItem = { _
    New LookupItem(CStr(System.Windows.Forms.BorderStyle.None), [Enum].GetName(GetType(System.Windows.Forms.BorderStyle), System.Windows.Forms.BorderStyle.None)), _
    New LookupItem(CStr(System.Windows.Forms.BorderStyle.FixedSingle), [Enum].GetName(GetType(System.Windows.Forms.BorderStyle), System.Windows.Forms.BorderStyle.FixedSingle))}
    cboBorderStyle.DisplayMember = "LookupDesc"
    cboBorderStyle.ValueMember = "LookupCode"
    cboBorderStyle.DataSource = vBS

    Dim vCurve() As LookupItem = { _
        New LookupItem("0", "0"), _
        New LookupItem("1", "1"), _
        New LookupItem("2", "2"), _
        New LookupItem("3", "3"), _
        New LookupItem("4", "4"), _
        New LookupItem("5", "5"), _
        New LookupItem("6", "6"), _
        New LookupItem("7", "7"), _
        New LookupItem("8", "8"), _
        New LookupItem("9", "9"), _
        New LookupItem("10", "10"), _
        New LookupItem("11", "11"), _
        New LookupItem("12", "12"), _
        New LookupItem("13", "13"), _
        New LookupItem("14", "14"), _
        New LookupItem("15", "15"), _
        New LookupItem("16", "16"), _
        New LookupItem("17", "17"), _
        New LookupItem("18", "18"), _
        New LookupItem("19", "19"), _
        New LookupItem("20", "20"), _
        New LookupItem("21", "21"), _
        New LookupItem("22", "22"), _
        New LookupItem("23", "23"), _
        New LookupItem("24", "24"), _
        New LookupItem("25", "25"), _
        New LookupItem("26", "26"), _
        New LookupItem("27", "27"), _
        New LookupItem("28", "28"), _
        New LookupItem("29", "29"), _
        New LookupItem("30", "30")}
    cboCurvature.DisplayMember = "LookupDesc"
    cboCurvature.ValueMember = "LookupCode"
    cboCurvature.DataSource = vCurve
  End Sub

  Public Sub InitFromPanelTheme(ByVal pPanelTheme As PanelTheme)
    cs1.Color = pPanelTheme.BackColor1
    cs2.Color = pPanelTheme.BackColor2
    SelectComboBoxItem(cboBorderStyle, CStr(pPanelTheme.BorderStyle))
    SelectComboBoxItem(cboGradientMode, CStr(pPanelTheme.GradientMode))
    SelectComboBoxItem(cboCurveMode, CStr(pPanelTheme.CurveMode))
    SelectComboBoxItem(cboCurvature, pPanelTheme.Curvature.ToString)
  End Sub

  Public Sub InitFromPanelTheme(ByVal pPanelTheme As PanelThemeSettings)
    cs1.Color = pPanelTheme.BackColor1
    cs2.Color = pPanelTheme.BackColor2
    SelectComboBoxItem(cboBorderStyle, CStr(pPanelTheme.BorderStyle))
    SelectComboBoxItem(cboGradientMode, CStr(pPanelTheme.GradientMode))
    SelectComboBoxItem(cboCurveMode, CStr(pPanelTheme.CurveMode))
    SelectComboBoxItem(cboCurvature, pPanelTheme.Curvature.ToString)
  End Sub

  Public Sub UpdatePanelTheme(ByVal pPanelTheme As PanelTheme)
    pPanelTheme.BackColor1 = cs1.Color
    pPanelTheme.BackColor2 = cs2.Color
    pPanelTheme.GradientMode = CType(cboGradientMode.SelectedValue, BackGroundExtender.LinearGradientMode)
    pPanelTheme.BorderStyle = CType(cboBorderStyle.SelectedValue, System.Windows.Forms.BorderStyle)
    pPanelTheme.CurveMode = CType(cboCurveMode.SelectedValue, BackGroundExtender.CornerCurveMode)
    pPanelTheme.Curvature = CInt(cboCurvature.SelectedValue)
    If pPanelTheme Is DisplayTheme.EditPanelTheme Then
      DisplayTheme.FillPanelTheme.BackColor1 = cs1.Color
      DisplayTheme.FillPanelTheme.BackColor2 = cs1.Color
      DisplayTheme.FillPanelTheme.GradientMode = CType(cboGradientMode.SelectedValue, BackGroundExtender.LinearGradientMode)
    End If
  End Sub

  Public Sub UpdatePanelTheme(ByVal pPanelTheme As PanelThemeSettings)
    pPanelTheme.BackColor1 = cs1.Color
    pPanelTheme.BackColor2 = cs2.Color
    pPanelTheme.GradientMode = CType(cboGradientMode.SelectedValue, BackGroundExtender.LinearGradientMode)
    pPanelTheme.BorderStyle = CType(cboBorderStyle.SelectedValue, System.Windows.Forms.BorderStyle)
    pPanelTheme.CurveMode = CType(cboCurveMode.SelectedValue, BackGroundExtender.CornerCurveMode)
    pPanelTheme.Curvature = CInt(cboCurvature.SelectedValue)
  End Sub
End Class
