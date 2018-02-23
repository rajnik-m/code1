Public Class frmFilter
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New(ByVal pList As ParameterList, ByVal pFieldType As DBField.FieldTypes, ByVal pRestrictions As DBFields, ByVal pLastRestrictions As DBFields)
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()
    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pList, pFieldType, pRestrictions, pLastRestrictions)
  End Sub

  'Form overrides dispose to clean up the component list.
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
  Friend WithEvents ButtonPanel1 As CDBNETCL.ButtonPanel
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents epl As CDBNETCL.EditPanel
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFilter))
    Me.ButtonPanel1 = New CDBNETCL.ButtonPanel
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.epl = New CDBNETCL.EditPanel
    Me.ButtonPanel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'ButtonPanel1
    '
    Me.ButtonPanel1.Controls.Add(Me.cmdOK)
    Me.ButtonPanel1.Controls.Add(Me.cmdCancel)
    Me.ButtonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.ButtonPanel1.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.ButtonPanel1.Location = New System.Drawing.Point(0, 178)
    Me.ButtonPanel1.Name = "ButtonPanel1"
    Me.ButtonPanel1.Size = New System.Drawing.Size(673, 39)
    Me.ButtonPanel1.TabIndex = 2
    '
    'cmdOK
    '
    Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOK.Location = New System.Drawing.Point(235, 6)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.Size = New System.Drawing.Size(94, 27)
    Me.cmdOK.TabIndex = 5
    Me.cmdOK.Text = "OK"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(344, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(94, 27)
    Me.cmdCancel.TabIndex = 6
    Me.cmdCancel.Text = "Cancel"
    '
    'epl
    '
    Me.epl.AddressChanged = False
    Me.epl.BackColor = System.Drawing.Color.Transparent
    Me.epl.DataChanged = False
    Me.epl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.epl.Location = New System.Drawing.Point(0, 0)
    Me.epl.Name = "epl"
    Me.epl.Recipients = Nothing
    Me.epl.Size = New System.Drawing.Size(673, 178)
    Me.epl.SuppressDrawing = False
    Me.epl.TabIndex = 3
    Me.epl.TabSelectedIndex = 0
    '
    'frmFilter
    '
    Me.AcceptButton = Me.cmdOK
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(673, 217)
    Me.Controls.Add(Me.epl)
    Me.Controls.Add(Me.ButtonPanel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmFilter"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.ButtonPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private mvList As ParameterList
  Private mvFields As DBFields
  Private mvLastFieldList As DBFields
  Private mvFieldName As String
  Private mvFieldType As DBField.FieldTypes

  Private Enum FilterSelectionOptions
    fsoNone
    fsoEquals
    fsoNotEquals
    fsoLike
    fsoNotLike
    fsoGreaterThan
    fsoGreaterThanOrEquals
    fsoLessThan
    fsoLessThanOrEquals
    fsoBeginsWith
    fsoNotBeginsWith
    fsoEndsWith
    fsoNotEndsWith
    fsoContains
    fsoNotContains
  End Enum

  Private Sub InitialiseControls(ByVal pList As ParameterList, ByVal pFieldType As DBField.FieldTypes, ByVal pRestrictions As DBFields, ByVal pLastRestrictions As DBFields)
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    SetControlColors(Me)
    Me.Text = ControlText.FrmEnterFilter
    mvList = pList
    mvFields = pRestrictions
    mvLastFieldList = pLastRestrictions
    mvFieldName = mvList("AttributeName")
    mvFieldType = pFieldType
    Dim vPanelInfo As New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optLMFilter)
    vPanelInfo.InitFilter(pList, pRestrictions)
    epl.Init(vPanelInfo)
    If mvList.ContainsKey("RestrictionValue") Then
      epl.FindTextLookupBox("Field1").FillComboWithRestriction(mvList("RestrictionValue"))
      epl.FindTextLookupBox("Field2").FillComboWithRestriction(mvList("RestrictionValue"))
    End If
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vList As New ParameterList
    If epl.AddValuesToList(vList, False, EditPanel.AddNullValueTypes.anvtAll) Then
      mvLastFieldList.Clear()
      mvLastFieldList = mvFields.Clone
      Dim vIndex As Integer
      Dim vCount As Integer
      Dim vExpression As String
      Do
        vExpression = "Expression" & CStr(vIndex + 1)
        If vList.Contains(vExpression) AndAlso vList(vExpression) <> "0" Then
          vIndex += 1
        Else
          Exit Do
        End If
      Loop
      If mvFields.ContainsKey(mvFieldName) Then mvFields.Remove(mvFieldName)
      Dim vIndex2 As Integer = 2
      While mvFields.ContainsKey(mvFieldName & "#" & vIndex2)
        mvFields.Remove(mvFieldName & "#" & vIndex2)
        vIndex2 += 1
      End While
      vCount = vIndex
      For vIndex = 1 To vCount
        SetWhereField(vList, vIndex, vCount)
      Next
      Me.Close()
    End If
  End Sub

  Private Sub GetWhereField(ByVal pField As DBField, ByVal pIndex As Integer)
    Dim vIndex As FilterSelectionOptions
    Dim vValue As String
    Dim vOperator As DBField.FieldWhereOperators

    vIndex = FilterSelectionOptions.fsoNone
    vValue = pField.Value
    vOperator = pField.WhereOperator And DBField.FieldWhereOperators.fwoOperatorOnly
    If pIndex > 1 Then
      If (pField.WhereOperator And DBField.FieldWhereOperators.fwoOR) > 0 Then
        epl.SetValue("AndOr" & pIndex.ToString, "OR")
      Else
        epl.SetValue("AndOr" & pIndex.ToString, "AND")
      End If
    End If
    Select Case vOperator
      Case DBField.FieldWhereOperators.fwoEqual
        vIndex = FilterSelectionOptions.fsoEquals
      Case DBField.FieldWhereOperators.fwoNotEqual
        vIndex = FilterSelectionOptions.fsoNotEquals
      Case DBField.FieldWhereOperators.fwoGreaterThan
        vIndex = FilterSelectionOptions.fsoGreaterThan
      Case DBField.FieldWhereOperators.fwoGreaterThanEqual
        vIndex = FilterSelectionOptions.fsoGreaterThanOrEquals
      Case DBField.FieldWhereOperators.fwoLessThan
        vIndex = FilterSelectionOptions.fsoLessThan
      Case DBField.FieldWhereOperators.fwoLessThanEqual
        vIndex = FilterSelectionOptions.fsoLessThanOrEquals
      Case DBField.FieldWhereOperators.fwoLike
        vIndex = FilterSelectionOptions.fsoLike
        If vValue.StartsWith("*") Then
          If vValue.EndsWith("*") Then
            vIndex = FilterSelectionOptions.fsoContains
            vValue = vValue.Substring(0, vValue.Length - 1)
          Else
            vIndex = FilterSelectionOptions.fsoEndsWith
          End If
          vValue = vValue.Substring(1)
        ElseIf vValue.EndsWith("*") Then
          vIndex = FilterSelectionOptions.fsoBeginsWith
          vValue = vValue.Substring(0, vValue.Length - 1)
        End If
      Case DBField.FieldWhereOperators.fwoNotLike
        vIndex = FilterSelectionOptions.fsoNotLike
        If vValue.StartsWith("*") Then
          If vValue.EndsWith("*") Then
            vIndex = FilterSelectionOptions.fsoNotContains
            vValue = vValue.Substring(0, vValue.Length - 1)
          Else
            vIndex = FilterSelectionOptions.fsoNotEndsWith
          End If
          vValue = vValue.Substring(1)
        ElseIf vValue.EndsWith("*") Then
          vIndex = FilterSelectionOptions.fsoNotBeginsWith
          vValue = vValue.Substring(0, vValue.Length - 1)
        End If
    End Select
    epl.SetValue("Expression" & pIndex, CInt(vIndex).ToString)
    epl.SetValue("Field" & pIndex, vValue)
  End Sub

  Private Sub SetWhereField(ByVal pList As ParameterList, ByVal pIndex As Integer, ByVal pCount As Integer)
    Dim vValue As String
    Dim vOperator As DBField.FieldWhereOperators
    Dim vOption As FilterSelectionOptions

    vOption = CType(pList("Expression" & pIndex), FilterSelectionOptions)
    vValue = pList("Field" & pIndex)
    If vValue.Length = 0 Then
      Select Case vOption
        Case FilterSelectionOptions.fsoLike, FilterSelectionOptions.fsoBeginsWith, FilterSelectionOptions.fsoEndsWith, FilterSelectionOptions.fsoContains, FilterSelectionOptions.fsoLessThanOrEquals, FilterSelectionOptions.fsoLessThan
          vOption = FilterSelectionOptions.fsoEquals
        Case FilterSelectionOptions.fsoNotLike, FilterSelectionOptions.fsoNotBeginsWith, FilterSelectionOptions.fsoNotEndsWith, FilterSelectionOptions.fsoNotContains, FilterSelectionOptions.fsoGreaterThanOrEquals, FilterSelectionOptions.fsoGreaterThan
          vOption = FilterSelectionOptions.fsoNotEquals
      End Select
    End If
    Select Case vOption
      Case FilterSelectionOptions.fsoEquals
        vOperator = DBField.FieldWhereOperators.fwoEqual
      Case FilterSelectionOptions.fsoNotEquals
        vOperator = DBField.FieldWhereOperators.fwoNotEqual
      Case FilterSelectionOptions.fsoLike
        vOperator = DBField.FieldWhereOperators.fwoLike
      Case FilterSelectionOptions.fsoNotLike
        vOperator = DBField.FieldWhereOperators.fwoNotLike
      Case FilterSelectionOptions.fsoGreaterThan
        vOperator = DBField.FieldWhereOperators.fwoGreaterThan
      Case FilterSelectionOptions.fsoGreaterThanOrEquals
        vOperator = DBField.FieldWhereOperators.fwoGreaterThanEqual
      Case FilterSelectionOptions.fsoLessThan
        vOperator = DBField.FieldWhereOperators.fwoLessThan
      Case FilterSelectionOptions.fsoLessThanOrEquals
        vOperator = DBField.FieldWhereOperators.fwoLessThanEqual
      Case FilterSelectionOptions.fsoBeginsWith
        vOperator = DBField.FieldWhereOperators.fwoLike
        If Not vValue.EndsWith("*") Then vValue &= "*"
      Case FilterSelectionOptions.fsoNotBeginsWith
        vOperator = DBField.FieldWhereOperators.fwoNotLike
        If Not vValue.EndsWith("*") Then vValue &= "*"
      Case FilterSelectionOptions.fsoEndsWith
        vOperator = DBField.FieldWhereOperators.fwoLike
        If Not vValue.StartsWith("*") Then vValue = "*" & vValue
      Case FilterSelectionOptions.fsoNotEndsWith
        vOperator = DBField.FieldWhereOperators.fwoNotLike
        If Not vValue.StartsWith("*") Then vValue = "*" & vValue
      Case FilterSelectionOptions.fsoContains
        vOperator = DBField.FieldWhereOperators.fwoLike
        If Not vValue.StartsWith("*") Then vValue = "*" & vValue
        If Not vValue.EndsWith("*") Then vValue &= "*"
      Case FilterSelectionOptions.fsoNotContains
        vOperator = DBField.FieldWhereOperators.fwoNotLike
        If Not vValue.StartsWith("*") Then vValue = "*" & vValue
        If Not vValue.EndsWith("*") Then vValue &= "*"
    End Select

    If pIndex > 1 Then
      If pList("AndOr" & pIndex) = "OR" Then vOperator = vOperator Or DBField.FieldWhereOperators.fwoOR
      If pIndex = pCount Then vOperator = vOperator Or DBField.FieldWhereOperators.fwoCloseBracket
      mvFields.Add(mvFieldName & "#" & pIndex, mvFieldType, vValue, vOperator)
    Else
      If pCount > 1 Then vOperator = vOperator Or DBField.FieldWhereOperators.fwoOpenBracket
      If mvFields.ContainsKey(mvFieldName) Then
        mvFields(mvFieldName).Value = vValue
        mvFields(mvFieldName).WhereOperator = vOperator
      Else
        mvFields.Add(mvFieldName, mvFieldType, vValue, vOperator)
      End If
    End If
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub frmFilter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
    epl.SetValue("FieldName_Label", mvList("AttributeNameDesc"))
    Dim vIndex As Integer
    If mvFields.ContainsKey(mvFieldName) Then
      GetWhereField(mvFields(mvFieldName), 1)
      vIndex = 2
      While mvFields.ContainsKey(mvFieldName & "#" & vIndex)
        'If vIndex >= mvMaxField Then AddFields(vIndex + 1)
        GetWhereField(mvFields(mvFieldName & "#" & vIndex), vIndex)
        vIndex = vIndex + 1
      End While
    End If
    Me.Width = 710
    Me.Height += (epl.RequiredHeight - epl.Height)
  End Sub

  Private Sub epl_ChangeHeight(ByVal pSender As Object, ByVal pNewHeight As Integer) Handles epl.ChangeHeight
    Me.Height += (pNewHeight - epl.Height)
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    'This means a new combo box will have been added
    If mvList.ContainsKey("RestrictionValue") Then
      epl.FindTextLookupBox(pParameterName).FillComboWithRestriction(mvList("RestrictionValue"))
    End If
  End Sub
End Class
