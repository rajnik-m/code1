Public Class frmConfigurationMaintenance

#Region "Private Members"

  Private mvNotes As CollectionList(Of String)
  Private mvDefaultValues As CollectionList(Of String)
  Private mvInProcess As Boolean

#End Region

  Public Sub New()

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

#Region "Private Methods"

  Private Sub InitialiseControls()
    SetControlTheme()
    mvNotes = New CollectionList(Of String)
    mvDefaultValues = New CollectionList(Of String)
  End Sub

  Private Sub GetConfigTree()
    Dim vKey As String
    Dim vName As String
    Dim vNode As TreeNode
    Dim vChildRows As DataRow()

    'Get all the config groups
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigGroups)

    If Not vDataTable Is Nothing AndAlso vDataTable.Rows.Count > 0 Then
      For Each vRow As DataRow In vDataTable.Rows
        If vRow.Item("ConfigGroup").ToString.Length = 2 Then 'Main node
          vNode = tvw.Nodes.Add(vRow.Item("ConfigGroup").ToString, vRow.Item("ConfigGroupDesc").ToString)
          vChildRows = vDataTable.Select("ConfigGroup like '" & vRow.Item("ConfigGroup").ToString & "%'", "ConfigGroupDesc")
          'Add child nodes
          For Each vChildRow As DataRow In vChildRows
            If Not vChildRow.Item("ConfigGroup").ToString = vRow.Item("ConfigGroup").ToString Then
              vNode.Nodes.Add(vChildRow.Item("ConfigGroup").ToString, vChildRow.Item("ConfigGroupDesc").ToString)
            End If
          Next
        End If
      Next

      tvw.Nodes.Add("UD", "Undefined Group").Tag = "UD"

      'Now get all the config names
      Dim vConfigNames As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigNames)
      If Not vConfigNames Is Nothing Then
        Dim vConfigNameNodes As TreeNode()
        For Each vRow As DataRow In vConfigNames.Rows
          vName = vRow.Item("ConfigName").ToString
          If Not Strings.Left(vName, 5) = "tabs_" AndAlso (Not Strings.Left(vName, 19) = "trader_applications") Then
            vKey = vRow.Item("ConfigGroup").ToString
            If vKey = String.Empty Then vKey = "UD"
            vConfigNameNodes = tvw.Nodes.Find(vKey, True)
            If vConfigNameNodes.Length > 0 Then
              vNode = vConfigNameNodes(0).Nodes.Add(vName, vRow.Item("ConfigNameDesc").ToString)
              vNode.Tag = vRow.Item("ConfigName").ToString
              If vRow.Item("ConfigDefaultValue") IsNot Nothing Then
                mvDefaultValues.Add(vRow.Item("ConfigName").ToString, vRow.Item("ConfigDefaultValue").ToString)
              End If
              mvNotes.Add(vRow.Item("ConfigName").ToString, vRow.Item("Notes").ToString)
            End If
          End If
        Next
      End If
    End If
  End Sub

#End Region

#Region "Event Handling"

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Me.Close()
  End Sub

  Private Sub cmdAmend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAmend.Click
    Try
      EditConfig(CareNetServices.XMLTableMaintenanceMode.xtmmAmend)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      mvInProcess = False
    End Try
  End Sub

  Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      EditConfig(CareNetServices.XMLTableMaintenanceMode.xtmmNew)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      mvInProcess = False
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      Dim vList As New ParameterList(True)
      Dim vNode As TreeNode = tvw.SelectedNode

      If vNode.Nodes.Count = 0 AndAlso mvInProcess = False Then
        mvInProcess = True
        vList("ConfigName") = dgr.GetValue(dgr.CurrentRow, "ConfigName")
        vList("ConfigValue") = dgr.GetValue(dgr.CurrentRow, "ConfigValue")
        vList("Client") = dgr.GetValue(dgr.CurrentRow, "Client")
        vList("Department") = dgr.GetValue(dgr.CurrentRow, "Department")
        vList("Logname") = dgr.GetValue(dgr.CurrentRow, "Logname")
        vList("MaintenanceTableName") = "config"
        If ConfirmDelete() Then
          DataHelper.DeleteTableMaintenanceData(vList)
          DisplaySelectedNode(vNode)
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      mvInProcess = False
    End Try
  End Sub

  Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
    Try
      Try
        'Passing in the dashboard item as the owner opens the search form just above the dashboard item
        Dim vFrmSearchtreeView As New frmSearchTreeView(tvw, Nothing)
        vFrmSearchtreeView.StartPosition = FormStartPosition.CenterParent
        vFrmSearchtreeView.ShowDialog()
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub frmConfigurationMaintenance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      GetConfigTree()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
#End Region

  ''' <summary>
  ''' Displays all the cofiguration options under the current node
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub DisplayAllConfigs(ByVal pNode As TreeNode)
    Dim vDataSet As New DataSet
    Dim vRow As DataRow
    Dim vTable As DataTable = DataHelper.NewColumnTable
    vDataSet.Tables.Add(vTable)
    DataHelper.AddDataColumn(vTable, "Key", "Key", , "N")
    DataHelper.AddDataColumn(vTable, "Description", "Configuration Items")
    DataHelper.AddDataColumn(vTable, "Value", "Value in use")

    Dim vDataTable As DataTable = New DataTable("DataRow")
    vDataTable.Columns.Add(New DataColumn("Key"))
    vDataTable.Columns.Add(New DataColumn("Description"))
    vDataTable.Columns.Add(New DataColumn("Value"))
    vDataSet.Tables.Add(vDataTable)

    Dim vConfigTable As DataTable = DataHelper.GetCachedLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigs)
    For vIndex As Integer = 0 To pNode.Nodes.Count - 1
      vRow = vDataTable.NewRow
      If Not pNode.Nodes(vIndex).Tag Is Nothing Then
        vRow("Key") = pNode.Nodes(vIndex).Tag.ToString
        Dim vRows() As DataRow = vConfigTable.Select(String.Format("ConfigName = '{0}'", pNode.Nodes(vIndex).Tag.ToString))
        If vRows.Length > 0 Then
          vRow("Value") = vRows(0).Item("ConfigValue").ToString
        Else
          vRow("Value") = String.Empty
        End If
      End If
      vRow("Description") = pNode.Nodes(vIndex).Text
      vDataTable.Rows.Add(vRow)
    Next
    dgr.Populate(vDataSet)

    If dgr.RowCount > 0 Then dgr.SelectRow(0)
  End Sub

  ''' <summary>
  ''' Displays the name of the config and the notes
  ''' </summary>
  ''' <param name="pKey"></param>
  ''' <remarks></remarks>
  Private Sub DisplayConfigAndNotes(ByVal pKey As String)
    If pKey.Length > 0 AndAlso Not pKey = "UD" Then
      txtConfigName.Text = pKey
      If mvDefaultValues.ContainsKey(pKey) Then
        Me.txtDefaultValue.Text = mvDefaultValues(pKey)
      End If
      If mvNotes.ContainsKey(pKey) Then
        txtNotes.Text = mvNotes(pKey).Replace(Chr(10).ToString, Environment.NewLine)
      End If
    Else
      txtConfigName.Text = String.Empty
      txtDefaultValue.Text = String.Empty
      txtNotes.Text = String.Empty
    End If
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pDataRow As System.Int32) Handles dgr.RowSelected
    Try
      DisplayConfigAndNotes(dgr.GetValue(pRow, 0))
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub tvw_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterSelect
    Try
      DisplaySelectedNode(e.Node)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub DisplaySelectedNode(ByVal pNode As TreeNode)
    If Not pNode.Tag Is Nothing Then
      DisplayConfigAndNotes(pNode.Tag.ToString)
    Else
      DisplayConfigAndNotes(String.Empty)
    End If

    cmdAmend.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False

    'Check what kind of node we are sitting on
    If pNode.Nodes.Count = 0 AndAlso pNode.Parent IsNot Nothing Then
      'We are at the bottom level and therefore looking at a config - Get all the entries
      Dim vList As New ParameterList(True)
      vList("ConfigName") = pNode.Tag.ToString
      dgr.Populate(DataHelper.GetConfigValue(vList))
      If dgr.RowCount = 0 Then
        vList.Clear()
        vList("ConfigName") = pNode.Tag.ToString
        vList("ConfigValue") = "(Value Not Set)"
        dgr.AddDataRow(vList)
        cmdNew.Enabled = True
      Else
        cmdAmend.Enabled = True
        cmdDelete.Enabled = True
        cmdNew.Enabled = True
      End If
    Else
      'We are at an upper level so list all the configs
      DisplayAllConfigs(pNode)
    End If
  End Sub

  Private Sub EditConfig(ByVal pMode As CareNetServices.XMLTableMaintenanceMode)
    Dim vNode As TreeNode
    Dim vCriteriaList As New ParameterList(True)
    Dim vValList As New ParameterList(True)

    vNode = tvw.SelectedNode
    If vNode.Nodes.Count = 0 And mvInProcess = False Then
      mvInProcess = True
      vCriteriaList("ConfigName") = dgr.GetValue(dgr.CurrentRow, "ConfigName")
      If pMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
        vValList("ConfigValue") = dgr.GetValue(dgr.CurrentRow, "ConfigValue")
        vCriteriaList("Client") = dgr.GetValue(dgr.CurrentRow, "Client")
        vCriteriaList("Department") = dgr.GetValue(dgr.CurrentRow, "Department")
        vCriteriaList("Logname") = dgr.GetValue(dgr.CurrentRow, "Logname")
      End If
      If DisplayTableEntry(pMode, "config", "Configuration Value", vValList, vCriteriaList) = System.Windows.Forms.DialogResult.OK Then
        DisplaySelectedNode(vNode)
      End If
    End If
  End Sub

  Private Function DisplayTableEntry(ByVal pEditMode As CareNetServices.XMLTableMaintenanceMode, ByVal pTable As String, ByVal pDesc As String, ByVal pParams As ParameterList, ByVal pCriteria As ParameterList) As DialogResult
    Dim vform As New frmTableEntry(pEditMode, pTable, pParams, pCriteria)
    Select Case pEditMode
      Case CareNetServices.XMLTableMaintenanceMode.xtmmNew
        vform.Text = ControlText.FrmAddTo & pDesc
      Case CareNetServices.XMLTableMaintenanceMode.xtmmAmend
        vform.Text = ControlText.FrmAmend & pDesc
    End Select
    Return vform.ShowDialog()
  End Function

  Private Sub txtConfigName_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConfigName.DoubleClick
    Dim vParams As New ParameterList(True)
    Dim vValList As New ParameterList(True)
    Dim vName As String
    Try
      If DataHelper.GetClientCode = "CARE" AndAlso txtConfigName.Text.Length > 0 Then
        vName = txtConfigName.Text
        vParams("ConfigName") = vName
        Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtConfigNames, vParams)
        If Not vTable Is Nothing AndAlso vTable.Rows.Count > 0 Then
          With vTable.Rows(0)
            For Each vCol As DataColumn In vTable.Columns
              vValList(vCol.ColumnName) = .Item(vCol.ColumnName).ToString
            Next
          End With
        End If
        If DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, "config_names", "Config Names", vValList, vParams) = System.Windows.Forms.DialogResult.OK Then
          tvw.SelectedNode.Text = vValList("ConfigNameDesc")
          txtNotes.Text = vValList("Notes")
          mvDefaultValues.Remove(txtConfigName.Text)
          mvNotes.Remove(txtConfigName.Text)
          mvDefaultValues.Add(txtConfigName.Text, txtDefaultValue.Text)
          mvNotes.Add(txtConfigName.Text, txtNotes.Text)
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub dgr_RowDoubleClicked(ByVal sender As System.Object, ByVal pRow As System.Int32) Handles dgr.RowDoubleClicked
    Try
      If cmdAmend.Enabled Then
        EditConfig(CareNetServices.XMLTableMaintenanceMode.xtmmAmend)
      ElseIf cmdNew.Enabled Then
        EditConfig(CareNetServices.XMLTableMaintenanceMode.xtmmNew)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      mvInProcess = False
    End Try
  End Sub
End Class