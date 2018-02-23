Public Class frmCriteriaLists

#Region "Private variables"
  Private mvMailingInfo As MailingInfo
  Private mvListManager As Boolean
  Private mvDataSet As DataSet
  Private mvResult As Boolean
  Private mvCriteriaSet As Integer
  Private mvCriteriaSetDesc As String
  Private mvRecordCount As Integer
  Friend WithEvents erp As System.Windows.Forms.ErrorProvider
#End Region

  Public Sub New(ByVal pMailingSelection As MailingInfo, Optional ByVal pListManager As Boolean = False)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pMailingSelection, pListManager)
  End Sub

  Private Sub InitialiseControls(ByVal pMailingSelection As MailingInfo, Optional ByVal pListManager As Boolean = False)
    Try
      SetControlTheme()
      mvMailingInfo = pMailingSelection
      mvListManager = pListManager
      dgrCriteriaSets.MaxGridRows = DisplayTheme.DefaultMaxGridRows

      Dim vParamList As New ParameterList(True)
      vParamList("TableName") = "departments"
      vParamList("FieldName") = "department"
      vParamList("FieldType") = "C"  ' Character FieldType
      Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vParamList)
      vParams("AttributeName") = "department"
      vParams("ValidationAttribute") = "department"
      vParams("ValidationTable") = "departments"

      txtLookupDepartment.BackColor = Me.BackColor
      Dim vPanelItem As PanelItem = New PanelItem(txtLookupDepartment, "department")
      vPanelItem.InitFromMaintenanceData(vParams)
      txtLookupDepartment.Init(vPanelItem, False, False)
      txtLookupDepartment.TotalWidth = txtLookupDepartment.Width
      txtLookupDepartment.SetBounds(txtLookupDepartment.Location.X, txtLookupDepartment.Location.Y, 80, txtLookupDepartment.TextBox.Size.Height)

      If mvListManager Then
        Me.Text = ControlText.FrmListManagerSteps
      Else
        Me.Text = mvMailingInfo.Caption & ControlText.FrmListManagerCriteriaSets
      End If

      Me.erp = New System.Windows.Forms.ErrorProvider(Me.components)

      GetCriteriaSets()

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


#Region "Private Methods"
  Private Sub GetCriteriaSets()
    'Get criteria sets and store the attributes in the specified grid
    'All sets will be retrieved
    'Similarly the sets can be limited by owner, description or department
    'Select the current row if there is one
    'and set the label for the number of selected records

    Dim vList As New ParameterList(True)
    vList("ApplicationName") = mvMailingInfo.MailingTypeCode
    If txtOwner.Text.Length > 0 Then vList("Owner") = txtOwner.Text
    If txtDescription.Text.Length > 0 Then vList("CriteriaSetDesc") = txtDescription.Text
    If txtLookupDepartment.Text.Length > 0 Then vList("Department") = txtLookupDepartment.Text

    If mvListManager Then
      vList("ListManager") = "Y"
    Else
      vList("ListManager") = "N"
    End If

    mvDataSet = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstCriteriaSets, CareServices.XMLTableDataSelectionTypes), vList)

    If mvDataSet IsNot Nothing AndAlso mvDataSet.Tables.Contains("Column") Then
      If Not mvDataSet.Tables.Contains("DataRow") Then AddDataRowTableToDataSet(mvDataSet)
      dgrCriteriaSets.Populate(mvDataSet)

      For vRowIndex As Integer = 0 To dgrCriteriaSets.RowCount - 1
        If dgrCriteriaSets.GetValue(vRowIndex, "CriteriaSetNumber") = mvMailingInfo.NewCriteriaSet.ToString Then
          dgrCriteriaSets.SelectRow(vRowIndex, True)
        End If
      Next
    End If

    mvRecordCount = dgrCriteriaSets.RowCount
    ChangeCurrentRow(dgrCriteriaSets.CurrentRow)
    lblMessage.Text = GetCountString()
  End Sub

  Private Sub ChangeCurrentRow(ByVal pRow As Integer)
    Dim vUpdatable As Boolean

    If pRow < 0 Then
      txtDescription.Text = ""
      txtLookupDepartment.Text = ""
      txtOwner.Text = ""
    Else
      If GridRowExists(dgrCriteriaSets) Then
        With dgrCriteriaSets
          txtDescription.Text = .GetValue(pRow, "CriteriaSetDesc")
          txtOwner.Text = .GetValue(pRow, "Owner")
          txtLookupDepartment.Text = .GetValue(pRow, "Department")
        End With
      Else
        cmdClear_Click(Me, Nothing)
      End If
    End If
    'Set the state of ok dependant on whether current row is blank
    cmdOK.Enabled = txtDescription.Text.Length > 0
    'Set the ability to delete or update Dependant on the owner
    vUpdatable = (txtOwner.Text = AppValues.Logname)
    cmdDelete.Enabled = vUpdatable
    cmdUpdate.Enabled = vUpdatable
  End Sub

  Private Function GetCountString() As String
    'Return a string specifying the number of records selected
    If mvRecordCount = 0 Then
      Return InformationMessages.ImNoRecordsSelected                      'No Records Selected
    Else
      Return String.Format(InformationMessages.ImRecordsSelected, mvRecordCount.ToString) 'LoadStringP1(27403, Format$(mvRecordCount))  '%s Records Selected
    End If
  End Function
#End Region

#Region "Other Events"
  Private Sub txtOwner_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwner.Leave
    Dim vCount As Integer
    Dim vList As New ParameterList(True)
    Try
      If txtOwner.Text.Length > 0 Then
        vList("Logname") = txtOwner.Text
        vCount = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctUsers, vList)
        If vCount = 0 Then
          ShowInformationMessage(InformationMessages.ImUnknownUser)
          txtOwner.Focus()
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrCriteriaSets_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgrCriteriaSets.RowDoubleClicked
    Try
      If cmdOK.Enabled = True Then
        cmdOK_Click(Me, Nothing)
        Me.Close() 'cmdOK.Value = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrCriteriaSets_RowSelected(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pDataRow As System.Int32) Handles dgrCriteriaSets.RowSelected
    Try
      If pRow >= 0 Then ChangeCurrentRow(pRow)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "Button Events"
  Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
    Try
      GetCriteriaSets()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Dim vSelectedRow As DataRow = dgrCriteriaSets.DataSourceDataRow(dgrCriteriaSets.CurrentDataRow)

      mvMailingInfo.NewCriteriaSet = IntegerValue(vSelectedRow("CriteriaSetNumber"))
      mvCriteriaSet = IntegerValue(vSelectedRow("CriteriaSetNumber"))
      mvCriteriaSetDesc = vSelectedRow("CriteriaSetDesc").ToString()
      mvResult = True
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
    Dim vList As New ParameterList(True)
    Dim vUpdate As Boolean
    Try
      Dim vSelectedRow As DataRow = dgrCriteriaSets.DataSourceDataRow(dgrCriteriaSets.CurrentDataRow)

      vList("CriteriaSet") = vSelectedRow("CriteriaSetNumber").ToString()
      If txtDescription.Text.Length > 0 Then
        vSelectedRow("CriteriaSetDesc") = txtDescription.Text
        vList("CriteriaSetDesc") = txtDescription.Text
        vUpdate = True
      End If
      If txtOwner.Text.Length > 0 Then
        vSelectedRow("Owner") = txtOwner.Text
        vList("Owner") = txtOwner.Text
        vUpdate = True
      End If
      If txtLookupDepartment.Text.Length > 0 Then
        vSelectedRow("Department") = txtLookupDepartment.Text
        vList("Department") = txtLookupDepartment.Text
        erp.SetError(txtLookupDepartment, "")
        vUpdate = True
      Else
        erp.SetError(txtLookupDepartment, InformationMessages.ImFieldMandatory)
        vUpdate = False
      End If
      vSelectedRow.AcceptChanges()

      If vUpdate Then
        DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctCriterialSet, vList)
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

    If dgrCriteriaSets.CurrentRow = -1 Then
      ShowInformationMessage(InformationMessages.ImSelectRow)
    Else
      Dim vList As New ParameterList(True)
      Dim vSelectedRow As DataRow = dgrCriteriaSets.DataSourceDataRow(dgrCriteriaSets.CurrentDataRow)
      vList.IntegerValue("CriteriaSet") = CInt(vSelectedRow("CriteriaSetNumber"))

      ' There is a possibility that criteria_set_details table might not have records & it will throw exception that unknown parameter criteria set 
      ' This must be ignored here while criteria_sets table must have record before deleting
      Try
        'Delete the criteria set details & criteria sets from the database
        DataHelper.DeleteCriteriaSetDetails(vList)
      Catch
      End Try
      Try
        Try
          'Delete the criteria set from the database
          DataHelper.DeleteCriteriaSet(vList)
        Catch vException As Exception
          'DataHelper.HandleException(vException) 'Concurrency error, the record was there when the form openned
        End Try

        mvRecordCount = mvRecordCount - 1

        lblMessage.Text = GetCountString()
        dgrCriteriaSets.DeleteRow(dgrCriteriaSets.CurrentRow)
        vSelectedRow.AcceptChanges() 'The grid gets very confused if you don't do this

        If dgrCriteriaSets.CurrentRow > 0 Then dgrCriteriaSets.SelectRow(dgrCriteriaSets.CurrentRow - 1, False)
        ChangeCurrentRow(dgrCriteriaSets.CurrentRow)

      Catch vException As Exception
        DataHelper.HandleException(vException)
      End Try
    End If
  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      If dgrCriteriaSets.DataRowCount > 0 Then
        dgrCriteriaSets.ClearDataRows()
      End If
      dgrCriteriaSets.MaxGridRows = 0
      ChangeCurrentRow(-1)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      mvMailingInfo.NewCriteriaSet = 0
      mvResult = False
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function GridRowExists(ByVal pDgr As DisplayGrid) As Boolean
    Return pDgr.RowCount > 0
  End Function
#End Region

#Region "Public Properties"
  Public ReadOnly Property Result() As Boolean
    Get
      Return mvResult
    End Get
  End Property

  Public ReadOnly Property CriteriaSet() As Integer
    Get
      Return mvCriteriaSet
    End Get
  End Property

  Public ReadOnly Property CriteriaSetDesc() As String
    Get
      Return mvCriteriaSetDesc
    End Get
  End Property
#End Region

End Class