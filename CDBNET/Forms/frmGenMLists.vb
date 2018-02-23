Public Class frmGenMLists

  Private mvMailingType As String = ""
  Private mvDepartment As String = ""
  Private mvDataset As DataSet = Nothing
  Private mvMailingInfo As MailingInfo
  Public Sub New(ByVal pMailingTypeCode As String, ByVal pMailingInfo As MailingInfo)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    mvMailingType = pMailingTypeCode
    mvDepartment = DataHelper.UserInfo.Department
    mvMailingInfo = pMailingInfo
    'mvCurrentSelectSet = pCurrentSelectionSet
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()

  End Sub
  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      If mvDataset IsNot Nothing Then mvMailingInfo.NewSelectionSet = CInt(dgr.DataSourceDataRow(dgr.CurrentRow).Item("SelectionSet")) 'BR17264 - bug fix and tidy up
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()
    dgr.MaxGridRows = DisplayTheme.DefaultMaxGridRows
    Dim vPanelItems As New PanelItems("epl")
    Dim vTmpDataSet As New DataSet
    If mvMailingInfo Is Nothing Then mvMailingInfo = New MailingInfo()
    'mvMailingInfo.Init(mvMailingType, 0)
    Dim vTmpDataTable As DataTable = DataHelper.NewColumnTable
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optGenMLists))
    GetSelectionSets()
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
  End Sub

  Private Sub GetSelectionSets(Optional ByVal pDescription As String = "", Optional ByVal pOwner As String = "", Optional ByVal pDepartment As String = "", Optional ByVal pSource As String = "")
    Dim vParams As New ParameterList(True)
    vParams("MailingType") = mvMailingType
    vParams("Department") = pDepartment
    vParams("UserName") = pOwner
    vParams("SelectionSetDesc") = pDescription
    vParams("Source") = pSource

    mvDataset = DataHelper.GetTableData(CType(CareNetServices.XMLTableDataSelectionTypes.xtdstSelectionSetData, CareServices.XMLTableDataSelectionTypes), vParams)
    If mvDataset IsNot Nothing And Not mvDataset.Tables.Contains("Column") Then
      Dim vTable As DataTable = DataHelper.NewColumnTable
      mvDataset.Tables.Add(vTable)
      DataHelper.AddDataColumn(vTable, "SelectionSet", "SelectionSet", , "N")
      DataHelper.AddDataColumn(vTable, "SelectionSetDesc", "Description")
      DataHelper.AddDataColumn(vTable, "UserName", "Owner")
      DataHelper.AddDataColumn(vTable, "department", "Department")
      DataHelper.AddDataColumn(vTable, "NumberInSet", "Records")
      DataHelper.AddDataColumn(vTable, "Source", "Source")
    End If
    If mvDataset IsNot Nothing And Not mvDataset.Tables.Contains("DataRow") Then
      AddDataRowTableToDataSet(mvDataset)
    End If
    If mvDataset IsNot Nothing AndAlso mvDataset.Tables.Contains("Column") AndAlso mvDataset.Tables.Contains("DataRow") Then
      dgr.Populate(mvDataset)
      If dgr.RowCount > 0 Then
        dgr.SelectRow(0)
      Else
        dgr.SelectRow(-1)
        cmdOK.Enabled = False
      End If
    End If
    DisplayGridTag()
  End Sub

  Private Sub DisplayGridTag()
    If dgr.RowCount = 0 Then
      lblContents.Text = InformationMessages.ImNoRecordsSelected
    Else
      lblContents.Text = String.Format(InformationMessages.ImRecordsSelected, dgr.RowCount.ToString)
    End If
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As System.Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    Try
      Dim vSource As String = ""
      If Not mvDataset Is Nothing Then
        If mvDataset.Tables("DataRow").Columns.Contains("Source") Then vSource = mvDataset.Tables("DataRow").Rows(pDataRow).Item("Source").ToString
        SetSelectionSetDetails(mvDataset.Tables("DataRow").Rows(pDataRow), vSource)
      End If
      cmdOK.Enabled = CBool(IIf(dgr.CurrentRow() <> -1, True, False))
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="vRow"></param>
  ''' <remarks></remarks>
  Private Sub SetSelectionSetDetails(ByVal vRow As DataRow, ByVal vSource As String)
    Dim vUpdatable As Boolean
    epl.Populate(vRow)
    If vSource <> String.Empty Then
      Dim vParams As New ParameterList(True)
      vParams("Source") = vRow("Source").ToString
      Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSources, vParams)
      If vDataTable IsNot Nothing Then epl.FindTextLookupBox("Source").Text = vDataTable.Rows(0).Item("Source").ToString
    End If
    vUpdatable = CBool(IIf(epl.FindTextBox("UserName").Text = DataHelper.UserInfo.Logname, True, False))
    cmdDelete.Enabled = vUpdatable
    cmdUpdate.Enabled = vUpdatable
  End Sub
  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
    Try
      Dim vParam As New ParameterList(True)
      If mvDataset IsNot Nothing Then
        'If mvDataset.Tables.Count > 1 Then
        Dim vParams As New ParameterList(True)
        If epl.AddValuesToList(vParams, True, EditPanel.AddNullValueTypes.anvtAll, True) Then
          Dim vDataRow As DataRow = mvDataset.Tables("DataRow").Rows(dgr.CurrentRow)
          vParams("SelectionSetNumber") = vDataRow("SelectionSet").ToString
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vParams)
          GetSelectionSets()
        End If
        'End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
    Try
      GetSelectionSets(epl.FindTextBox("SelectionSetDesc").Text, epl.FindTextBox("UserName").Text, epl.FindTextLookupBox("Department").Text, epl.FindTextLookupBox("Source").Text)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      dgr.ClearDataRows()
      epl.Clear()
      DisplayGridTag()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      If mvDataset IsNot Nothing Then
        Dim vParams As New ParameterList(True)
        Dim vDataRow As DataRow = mvDataset.Tables("DataRow").Rows(dgr.CurrentRow)
        vParams("SelectionSetNumber") = vDataRow("SelectionSet").ToString
        mvMailingInfo.DeleteSelection(CInt(vDataRow("SelectionSet")), 0)
        DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet, vParams)
        GetSelectionSets()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub drg_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgr.RowDoubleClicked
    Try
      If cmdOK.Enabled = True Then
        cmdOK_Click(Me, Nothing)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class