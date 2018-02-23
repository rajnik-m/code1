Imports System.Xml.Linq
Imports Advanced.Extensibility

Public Class frmReportMaintenance
  Private mvReportSet As DataSet
  Private mvReportSectionSet As DataSet
  Private WithEvents mvReportMenu As ReportMenu
  Private mvDataChanged As Boolean
  Private mvPrevReportNumber As Integer
  Private mvNewReportNumber As Integer
  Private mvDataSet As New DataSet
  Private Property ExtensionDictionary As Dictionary(Of Control, XElement)

  Public Sub New()

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()

    'InitialiseExtensions()
    ExtensionDictionary = New Dictionary(Of Control, XElement)
  End Sub

  Private Sub InitialiseControls()
    mvReportMenu = New ReportMenu()
    Dim vNodeInfo As TreeViewNodeInfo
    LoadReport()
    LoadReportSection()
    tvw.Indent = AppValues.TreeViewIndent
    tvw.Nodes.Clear()

    'Creating System Report Node
    Dim vNewNode As TreeNode = New TreeNode("System Reports")
    vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.ReportType, 1)
    vNewNode.Tag = vNodeInfo
    BuildNode(vNewNode, 1)
    tvw.Nodes.Add(vNewNode)

    'Creating System User Reports Node
    vNewNode = New TreeNode("System User Reports")
    vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.ReportType, 2)
    vNewNode.Tag = vNodeInfo
    BuildNode(vNewNode, 2)
    tvw.Nodes.Add(vNewNode)

    'Creating Custom Report Node
    vNewNode = New TreeNode("Custom Reports")
    vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.ReportType, 3)
    vNewNode.Tag = vNodeInfo
    BuildNode(vNewNode, 3)
    tvw.Nodes.Add(vNewNode)

    'Creating Custom User Report Node
    vNewNode = New TreeNode("Custom User Reports")
    vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.ReportType, 4)
    vNewNode.Tag = vNodeInfo
    BuildNode(vNewNode, 4)
    tvw.Nodes.Add(vNewNode)

    'Loading Default View 
    mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber < 10000 AND ReportCode <> 'USER'"
    dgr.Populate(mvReportSet)
    tvw.SelectedNode = tvw.Nodes(0)
    dgr.AllowSorting = False
  End Sub
  Private Sub LoadReport()
    Dim vList As New ParameterList(True)
    mvReportSet = DataHelper.GetReportData(vList)
  End Sub
  Private Sub LoadReportSection()
    Dim vList As New ParameterList(True)
    mvReportSectionSet = DataHelper.GetReportSectionData(vList)
  End Sub
  Private Sub BuildSectionNode(ByVal pNode As TreeNode, ByVal pReportNumber As Integer)
    pNode.Nodes.Clear()
    Dim vNodeInfo As TreeViewNodeInfo = CType(pNode.Tag, TreeViewNodeInfo)
    Dim vNewNodeInfo As TreeViewNodeInfo
    mvReportSectionSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber = '" + pReportNumber.ToString() + "'"
    For Each vDataRowRView As DataRowView In mvReportSectionSet.Tables("DataRow").DefaultView
      'pNode.Nodes(pReportNumber.ToString()).Nodes.Add(vDataRowRView("SectionNumber").ToString(), vDataRowRView("SectionName").ToString())
      pNode.Nodes.Add(vDataRowRView("SectionNumber").ToString(), vDataRowRView("SectionName").ToString())
      vNewNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.SubSection, vNodeInfo.ReportType, IntegerValue(vDataRowRView("ReportNumber")), IntegerValue(vDataRowRView("SectionNumber")), vNodeInfo.ReportCode)
      pNode.Nodes(vDataRowRView("SectionNumber").ToString()).Tag = vNewNodeInfo
    Next
  End Sub
  ''' <summary>
  ''' This function will build child node for each report type
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <param name="pReportType"></param>
  ''' <param name="pReportNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildNode(ByVal pNode As TreeNode, ByVal pReportType As Integer, Optional ByVal pReportNumber As Integer = 0) As TreeNode
    Dim vNodeInfo As TreeViewNodeInfo
    Select Case pReportType
      Case 1
        mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber < 10000 AND ReportCode <> 'USER'"
      Case 2
        mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber < 10000 AND ReportCode = 'USER'"
      Case 3
        mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber >= 10000 AND ReportCode <> 'USER'"
      Case 4
        mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber >= 10000 AND ReportCode = 'USER'"
    End Select
    If pReportNumber > 0 Then
      mvReportSet.Tables("DataRow").DefaultView.RowFilter = mvReportSet.Tables("DataRow").DefaultView.RowFilter + " AND ReportNumber =" + pReportNumber.ToString()
    End If

    For Each vDataRowView As DataRowView In mvReportSet.Tables("DataRow").DefaultView
      'Create Node : Report 
      pNode.Nodes.Add(vDataRowView("ReportNumber").ToString(), vDataRowView("ReportName").ToString())
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Report, pReportType, IntegerValue(vDataRowView("ReportNumber")), , vDataRowView("ReportCode").ToString())
      pNode.Nodes(vDataRowView("ReportNumber").ToString).Tag = vNodeInfo

      'Create Node : Parameters 
      pNode.Nodes(vDataRowView("ReportNumber").ToString).Nodes.Add("Parameters", "Parameters")
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Parameter, pReportType, IntegerValue(vDataRowView("ReportNumber")), , vDataRowView("ReportCode").ToString())
      pNode.Nodes(vDataRowView("ReportNumber").ToString).Nodes("Parameters").Tag = vNodeInfo

      'Create Node : Sections
      pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes.Add("Sections", "Sections")
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Section, pReportType, IntegerValue(vDataRowView("ReportNumber")), , vDataRowView("ReportCode").ToString())
      pNode.Nodes(vDataRowView("ReportNumber").ToString).Nodes("Sections").Tag = vNodeInfo

      'Create Node : Sub Sections
      mvReportSectionSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber = '" + vDataRowView("ReportNumber").ToString() + "'"
      For Each vDataRowRView As DataRowView In mvReportSectionSet.Tables("DataRow").DefaultView
        pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes("Sections").Nodes.Add(vDataRowRView("SectionNumber").ToString(), vDataRowRView("SectionName").ToString())

        vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.SubSection, pReportType, IntegerValue(vDataRowRView("ReportNumber")), IntegerValue(vDataRowRView("SectionNumber")), vDataRowView("ReportCode").ToString())
        pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes("Sections").Nodes(vDataRowRView("SectionNumber").ToString()).Tag = vNodeInfo
      Next

      'Create Node : Version
      pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes.Add("Versions", "Versions")
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Version, pReportType, IntegerValue(vDataRowView("ReportNumber")), , vDataRowView("ReportCode").ToString())
      pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes("Versions").Tag = vNodeInfo

      'Create Node : Control
      pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes.Add("Controls", "Controls")
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Control, pReportType, IntegerValue(vDataRowView("ReportNumber")), , vDataRowView("ReportCode").ToString())
      pNode.Nodes(vDataRowView("ReportNumber").ToString()).Nodes("Controls").Tag = vNodeInfo
    Next
    Return pNode
  End Function

  ''' <summary>
  ''' This event will be fired when menu process will get completed
  ''' This will refresh current node
  ''' </summary>
  ''' <param name="pMenuItem"></param>
  ''' <remarks></remarks>
  Private Sub mvReportMenu_ItemSelected(ByVal pMenuItem As ReportMenu.ReportMenuItems) Handles mvReportMenu.MenuActionCompleted
    Try
      Dim vNodeInfo As TreeViewNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
      Select Case pMenuItem
        Case ReportMenu.ReportMenuItems.rmiRenumberParameter
          mvDataChanged = True
          ShowParameterData(vNodeInfo.ReportNumber)
          tvw.SelectedNode = tvw.Nodes(vNodeInfo.ReportType - 1).Nodes(vNodeInfo.ReportNumber.ToString()).Nodes("Parameters")
        Case ReportMenu.ReportMenuItems.rmiRenumberSection
          mvDataChanged = True
          LoadReportSection()
          ShowIndividualReportData(vNodeInfo.ReportNumber)
          tvw.SelectedNode = tvw.Nodes(vNodeInfo.ReportType - 1).Nodes(vNodeInfo.ReportNumber.ToString())
        Case ReportMenu.ReportMenuItems.rmiDuplicateReport
          If mvReportMenu.ReportNumber > 0 Then
            LoadReport()
            LoadReportSection()
            Dim vReportType As Integer = 1
            If mvReportMenu.ReportNumber < 10000 And vNodeInfo.ReportCode <> "USER" Then
              vReportType = 1
            ElseIf mvReportMenu.ReportNumber < 10000 And vNodeInfo.ReportCode = "USER" Then
              vReportType = 2
            ElseIf mvReportMenu.ReportNumber >= 10000 And vNodeInfo.ReportCode <> "USER" Then
              vReportType = 3
            ElseIf mvReportMenu.ReportNumber >= 10000 And vNodeInfo.ReportCode = "USER" Then
              vReportType = 4
            End If

            BuildNode(tvw.Nodes(vReportType - 1), vReportType, mvReportMenu.ReportNumber)
            tvw.SelectedNode = tvw.Nodes(vReportType - 1).Nodes(mvReportMenu.ReportNumber.ToString())
            ShowIndividualReportData(mvReportMenu.ReportNumber)
          End If
        Case ReportMenu.ReportMenuItems.rmiDuplicateSection
          mvDataChanged = True
          LoadReportSection()
          BuildSectionNode(tvw.Nodes(vNodeInfo.ReportType - 1).Nodes(vNodeInfo.ReportNumber.ToString()).Nodes("Sections"), vNodeInfo.ReportNumber)
          ShowSectionData(vNodeInfo.ReportNumber)
        Case ReportMenu.ReportMenuItems.rmiRenumberItem
          ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)
          mvDataChanged = True
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub tvw_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterSelect
    Try
      Dim vNodeInfo As TreeViewNodeInfo = CType(e.Node.Tag, TreeViewNodeInfo)
      If mvPrevReportNumber <> vNodeInfo.ReportNumber Then
        CheckForVersionHistory()
      End If
      tvw.SelectedNode = e.Node
      ShowReportData(e.Node)
      InitialiseExtensions(vNodeInfo.ReportNumber)
      mvPrevReportNumber = vNodeInfo.ReportNumber
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub tvw_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvw.MouseDown
    Try
      If e.Button = System.Windows.Forms.MouseButtons.Right Then
        Dim vNode As TreeNode = tvw.GetNodeAt(e.X, e.Y)
        tvw.SelectedNode = vNode
        If Not vNode Is Nothing AndAlso Not vNode.Tag Is Nothing Then
          tvw.ContextMenuStrip = mvReportMenu
          If tvw.ContextMenuStrip IsNot Nothing Then
            Dim vNodeInfo As TreeViewNodeInfo = CType(vNode.Tag, TreeViewNodeInfo)
            mvReportMenu.ReportNodeInfo = vNodeInfo
            If mvReportMenu.ReportNodeInfo.NodeType = TreeViewNodeType.ReportType Then
              tvw.ContextMenuStrip = Nothing
            End If
          End If
        Else
          tvw.ContextMenuStrip = Nothing
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub CheckForVersionHistory()
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    Dim vSequenceNumber As Integer = 1
    Dim vAdd As Boolean
    Dim vRowPos As Integer = 0

    If mvDataChanged Then
      vList("ReportNumber") = mvPrevReportNumber.ToString()
      vDataSet = DataHelper.GetReportVersion(vList)

      If vDataSet.Tables("DataRow") IsNot Nothing Then
        vDataSet.Tables("DataRow").DefaultView.Sort = "VersionNumber DESC"
        If vDataSet.Tables("DataRow").Rows.Count > 0 Then
          vRowPos = vDataSet.Tables("DataRow").Rows.Count - 1 'last row
          If vDataSet.Tables("DataRow").Rows(vRowPos).Item("LogName").ToString() = DataHelper.UserInfo.Logname And vDataSet.Tables("DataRow").Rows(vRowPos).Item("ChangeDate").ToString() = AppValues.TodaysDate() Then
            vSequenceNumber = IntegerValue(vDataSet.Tables("DataRow").Rows(vRowPos).Item("VersionNumber"))
            vAdd = False
            vList("ChangeDescription") = vDataSet.Tables("DataRow").Rows(vRowPos).Item("ChangeDescription").ToString()
          Else
            vSequenceNumber = IntegerValue(vDataSet.Tables("DataRow").Rows(vDataSet.Tables("DataRow").Rows.Count - 1).Item("VersionNumber").ToString()) + 1
            vAdd = True
          End If
        Else
          vSequenceNumber = vSequenceNumber + 1
          vAdd = True
        End If
      Else
        vAdd = True
      End If
      vList("VersionNumber") = vSequenceNumber.ToString()
      vList("Logname") = DataHelper.UserInfo.Logname
      vList("ChangeDate") = AppValues.TodaysDate
      If vAdd Then
        Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, "report_version_history", vList)
      Else
        Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, "report_version_history", vList)
      End If
    End If
    mvDataChanged = False
  End Sub

  Private Sub ShowReportData(ByVal pNode As TreeNode)
    Dim vNodeInfo As TreeViewNodeInfo = CType(pNode.Tag, TreeViewNodeInfo)
    dgr.AutoSetRowHeight = False
    Select Case vNodeInfo.NodeType
      Case TreeViewNodeType.ReportType
        Select Case vNodeInfo.ReportType
          Case 1
            mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber < 10000 AND ReportCode <> 'USER'"
            dgr.Populate(mvReportSet)
          Case 2
            mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber < 10000 AND ReportCode = 'USER'"
            dgr.Populate(mvReportSet)
          Case 3
            mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber >= 10000 AND ReportCode <> 'USER'"
            dgr.Populate(mvReportSet)
          Case 4
            mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber >= 10000 AND ReportCode = 'USER'"
            dgr.Populate(mvReportSet)
        End Select
        dgr.SetColumnHeaderVisible(True)

      Case TreeViewNodeType.Report
        ShowIndividualReportData(vNodeInfo.ReportNumber)
        dgr.AllowRowMove(False)
        dgr.SetRowHeaderVisible()
        dgr.SetPreferredRowHeaderWidth(0)
      Case TreeViewNodeType.Section
        dgr.AutoSetRowHeight = True
        mvDataSet = mvReportSectionSet.Copy
        ShowSectionData(vNodeInfo.ReportNumber)
        dgr.AllowRowMove()
      Case TreeViewNodeType.Parameter
        ShowParameterData(vNodeInfo.ReportNumber)
        dgr.AllowRowMove()
      Case TreeViewNodeType.SubSection
        ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)
        dgr.AllowRowMove()
      Case TreeViewNodeType.Version
        ShowVersion(vNodeInfo.ReportNumber)
        ShowButton(True)
        cmdDelete.Enabled = False
        If dgr.RowCount = 0 Then cmdAmend.Enabled = False
      Case TreeViewNodeType.Control
        ShowControl(vNodeInfo.ReportNumber, vNodeInfo.ReportCode)
    End Select
  End Sub

  Private Sub ShowSectionData(ByVal pReportNumber As Integer)
    Dim vTable As New DataTable("DataRow")
    mvReportSectionSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber=" + pReportNumber.ToString()
    mvDataSet.Tables.Clear()
    mvDataSet.Tables.Add(mvReportSectionSet.Tables("Column").Copy)
    For Each vCol As DataColumn In mvReportSectionSet.Tables("DataRow").Columns
      vTable.Columns.Add(vCol.ColumnName)
    Next
    For Each vRow As DataRowView In mvReportSectionSet.Tables("DataRow").DefaultView
      vTable.ImportRow(vRow.Row)
    Next
    mvDataSet.Tables.Add(vTable)
    dgr.Populate(mvDataSet)
    dgr.SetColumnHeaderVisible(True)
    If mvReportSectionSet.Tables("DataRow") Is Nothing OrElse mvReportSectionSet.Tables("DataRow").DefaultView.Count = 0 Then
      ShowButton(False)
    Else
      ShowButton(True)
    End If
  End Sub
  Private Sub ShowButton(ByVal pEnable As Boolean)
    cmdAmend.Enabled = pEnable
    cmdDelete.Enabled = pEnable
    cmdRun.Enabled = pEnable
  End Sub
  Private Sub ShowControl(ByVal pReportNumber As Integer, ByVal pReportCode As String)
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    vList("ApplicationNumber") = pReportNumber.ToString()
    vList("PageType") = pReportCode
    vDataSet = DataHelper.GetReportControl(vList)
    dgr.Populate(vDataSet)
    dgr.SetColumnHeaderVisible(True)
    If vDataSet.Tables("DataRow") Is Nothing OrElse vDataSet.Tables("DataRow").DefaultView.Count = 0 Then
      ShowButton(False)
    Else
      ShowButton(True)
    End If
  End Sub
  Private Sub ShowVersion(ByVal pReportNumber As Integer)
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    vList("ReportNumber") = pReportNumber.ToString()
    vDataSet = DataHelper.GetReportVersion(vList)
    dgr.Populate(vDataSet)
    dgr.SetColumnHeaderVisible(True)
  End Sub
  Private Sub ShowSubSectionDetail(ByVal pReportNumber As Integer, ByVal pSectionNumber As Integer)
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    vList("ReportNumber") = pReportNumber.ToString()
    vList("SectionNumber") = pSectionNumber.ToString()
    vDataSet = DataHelper.GetReportSectionDetail(vList)
    dgr.Populate(vDataSet)
    dgr.SetColumnHeaderVisible(True)
    If vDataSet.Tables("DataRow") Is Nothing OrElse vDataSet.Tables("DataRow").Rows.Count = 0 Then
      ShowButton(False)
    Else
      ShowButton(True)
    End If

  End Sub
  Private Sub ShowParameterData(ByVal pReportNumber As Integer)
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    vList("ReportNumber") = pReportNumber.ToString()
    vDataSet = DataHelper.GetReportParameters(vList)
    dgr.Populate(vDataSet)
    dgr.SetColumnHeaderVisible(True)
    If vDataSet.Tables("DataRow") Is Nothing OrElse vDataSet.Tables("DataRow").DefaultView.Count = 0 Then
      ShowButton(False)
    Else
      ShowButton(True)
    End If
  End Sub

  Private Sub ShowIndividualReportData(ByVal pReportNumber As Integer)
    Dim vDataSet As New DataSet
    Dim vNewRow As DataRow
    Dim vDataTable As New DataTable("DataRow")
    vDataTable.Columns.Add()
    Dim vRow As DataRow()
    vRow = mvReportSet.Tables("DataRow").Select("ReportNumber = '" + pReportNumber.ToString() + "'")
    If vRow.Length > 0 Then
      For intCtr As Integer = 0 To vRow(0).ItemArray.Length - 1
        vNewRow = vDataTable.NewRow()
        vNewRow(0) = vRow(0).Item(intCtr).ToString()
        vDataTable.Rows.Add(vNewRow)
      Next
      vDataSet.Tables.Add(vDataTable)
      dgr.Populate(vDataSet)
      dgr.SetColumnHeaderVisible(False)
      dgr.SetRowHeaderVisible()
      For intIncr As Integer = 0 To mvReportSet.Tables("Column").Rows.Count - 1
        dgr.SetRowHeaderValue(intIncr, 0, mvReportSet.Tables("Column").Rows(intIncr).Item("Heading").ToString())
      Next
      dgr.SetPreferredRowHeaderWidth(0)
    End If
    ShowButton(True)
  End Sub

  Private Sub cmdRun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRun.Click
    Try
      Dim vReportList As New ParameterList(True)
      Dim vNodeInfo As TreeViewNodeInfo
      vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
      vReportList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
      vReportList("ReportCode") = "USER"
      Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFind.Click
    Try
      Dim vFinder As New frmSearchTreeView(tvw, Nothing)
      vFinder.ShowDialog()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Try
      Me.Close()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  Private Function DisplayTableEntry(ByVal pEditMode As CareNetServices.XMLTableMaintenanceMode, ByVal pTable As String, ByVal pParams As ParameterList, Optional ByVal pCriteria As ParameterList = Nothing) As DialogResult
    Dim vCriteria As New ParameterList
    Dim vResult As New DialogResult
    If pCriteria IsNot Nothing Then vCriteria = pCriteria
    Dim vform As New frmTableEntry(pEditMode, pTable, pParams, vCriteria, , True)
    'set the form caption and whether the Add More menu is visible
    Select Case pEditMode
      Case CareNetServices.XMLTableMaintenanceMode.xtmmNew
        vform.Text = ControlText.FrmAddTo + pTable
      Case CareNetServices.XMLTableMaintenanceMode.xtmmAmend
        vform.Text = ControlText.FrmAmend + pTable
      Case CareNetServices.XMLTableMaintenanceMode.xtmmSelect
        vform.Text = ControlText.FrmSelectForm + pTable
    End Select
    vResult = vform.ShowDialog()
    If pParams.ContainsKey("ReportNumber") Then mvNewReportNumber = IntegerValue(pParams.Item("ReportNumber").ToString)
    Return vResult
  End Function

  Private Sub UpdateReportData()
    Dim vNodeInfo As TreeViewNodeInfo
    Dim vCurrentTable As String = ""
    Dim vParams As New ParameterList(True)
    Dim vCriteria As New ParameterList(True)

    vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
    Select Case vNodeInfo.NodeType
      Case TreeViewNodeType.ReportType
        vCurrentTable = "reports"
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
            End If
          Next
        End If
      Case TreeViewNodeType.Report
        Dim vRow As DataRow
        vCurrentTable = "reports"
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          vRow = mvReportSet.Tables("DataRow").Select("ReportNumber = '" + vNodeInfo.ReportNumber.ToString + "'")(0)
          For vCtr As Integer = 0 To mvReportSet.Tables("DataRow").Columns.Count - 1
            If vRow.Item(vCtr).ToString.Length > 0 Then
              vParams(mvReportSet.Tables("DataRow").Columns(vCtr).Caption) = vRow.Item(vCtr).ToString()
            End If
          Next
        End If
      Case TreeViewNodeType.Parameter
        vCurrentTable = "report_parameters"
        vParams("MaintenanceTableName") = vCurrentTable
        vParams("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
            End If
          Next
        End If
      Case TreeViewNodeType.Section
        vCurrentTable = "report_sections"
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
            End If
          Next
          vParams.Remove("SectionTypeDesc")
        End If
      Case TreeViewNodeType.SubSection
        vCurrentTable = "report_items"
        vParams("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        vParams("SectionNumber") = vNodeInfo.SectionNumber.ToString()
        vParams("ReportItemType") = dgr.GetValue(dgr.ActiveRow, "ReportItemType")
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
            End If
          Next
        End If
        vParams.Remove("ReportItemTypeDesc")
      Case TreeViewNodeType.Version
        vCurrentTable = "report_version_history"
        vCriteria("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              If dgr.ColumnName(vCtr) = "VersionNumber" Then
                vCriteria(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
              Else
                vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
              End If
            End If
          Next
        End If
      Case TreeViewNodeType.Control
        vCurrentTable = "fp_controls"
        vCriteria("FpApplication") = vNodeInfo.ReportNumber.ToString()
        vCriteria("FpPageType") = vNodeInfo.ReportCode.ToString()
        vParams("MaintenanceTableName") = vCurrentTable
        If dgr.RowCount > 0 Then
          For vCtr As Integer = 0 To dgr.ColumnCount - 1
            If dgr.GetValue(dgr.ActiveRow, vCtr).Length > 0 Then
              If dgr.ColumnName(vCtr) = "SequenceNumber" Then
                vCriteria(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
              Else
                vParams(dgr.ColumnName(vCtr)) = dgr.GetValue(dgr.ActiveRow, vCtr)
              End If
            End If
          Next
        End If
    End Select
    Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, vCurrentTable, vParams, vCriteria)
    If vResult = System.Windows.Forms.DialogResult.OK Then
      mvDataChanged = True
      Select Case vNodeInfo.NodeType
        Case TreeViewNodeType.ReportType
          LoadReport()
          ShowReportData(tvw.Nodes(vNodeInfo.ReportType - 1))
        Case TreeViewNodeType.Report
          LoadReport()
          ShowIndividualReportData(vNodeInfo.ReportNumber)
        Case TreeViewNodeType.Parameter
          ShowParameterData(vNodeInfo.ReportNumber)
        Case TreeViewNodeType.Section
          LoadReportSection()
          ShowSectionData(vNodeInfo.ReportNumber)
        Case TreeViewNodeType.SubSection
          ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)
        Case TreeViewNodeType.Version
          ShowVersion(vNodeInfo.ReportNumber)
        Case TreeViewNodeType.Control
          ShowControl(vNodeInfo.ReportNumber, vNodeInfo.ReportCode)
      End Select
      UpdateCurrentNode(tvw.SelectedNode)
    End If
  End Sub
  Private Sub cmdAmend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAmend.Click
    Try
      UpdateReportData()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      Dim vNodeInfo As TreeViewNodeInfo
      Dim vParamaterNumber As Integer
      Dim vList As New ParameterList(True)
      vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
      Select Case vNodeInfo.NodeType
        Case TreeViewNodeType.ReportType
          If ShowQuestion(QuestionMessages.QmDeleteReportAndAllComponents, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            vList("ReportNumber") = dgr.GetValue(dgr.ActiveRow, 0)
            vList("ReportCode") = dgr.GetValue(dgr.ActiveRow, 2)
            DataHelper.DeleteReportAndContent(vList)
            tvw.Nodes(vNodeInfo.ReportType - 1).Nodes(dgr.GetValue(dgr.ActiveRow, 0).ToString()).Remove()
            mvReportSet.Tables("DataRow").DefaultView.Delete(dgr.ActiveRow)
            ShowReportData(tvw.Nodes(vNodeInfo.ReportType - 1))
          End If
        Case TreeViewNodeType.Report
          If ShowQuestion(QuestionMessages.QmDeleteReportAndAllComponents, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
            vList("ReportCode") = vNodeInfo.ReportCode
            DataHelper.DeleteReportAndContent(vList)
            tvw.SelectedNode.Remove()
            LoadReport()
            ShowReportData(tvw.Nodes(vNodeInfo.ReportType - 1))
            mvDataChanged = True
          End If
        Case TreeViewNodeType.Parameter
          If dgr.RowCount > 0 Then
            If ShowQuestion(QuestionMessages.QmDeleteReportParameter, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vParamaterNumber = IntegerValue(dgr.GetValue(dgr.ActiveRow, 0))
              vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
              vList("ParameterNumber") = vParamaterNumber.ToString()
              DataHelper.DeleteReportParameter(vList)
              ShowParameterData(vNodeInfo.ReportNumber)
              mvDataChanged = True
            End If
          End If
        Case TreeViewNodeType.Section
          Dim vSectionNumber As Integer
          If dgr.RowCount > 0 Then
            If ShowQuestion(QuestionMessages.QmDeleteReportSectionAndAssociatedItems, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vSectionNumber = IntegerValue(dgr.GetValue(dgr.ActiveRow, "SectionNumber"))
              vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
              vList("SectionNumber") = vSectionNumber.ToString()
              DataHelper.DeleteReportSection(vList)

              tvw.SelectedNode.Nodes(vSectionNumber.ToString()).Remove()
              mvReportSectionSet.Tables("DataRow").DefaultView.Delete(dgr.ActiveRow)
              ShowSectionData(vNodeInfo.ReportNumber)
              mvDataChanged = True
            End If
          End If
        Case TreeViewNodeType.SubSection
          Dim vSectionNumber As Integer
          Dim vItemNumber As Integer
          If dgr.RowCount > 0 Then
            If ShowQuestion(QuestionMessages.QmDeleteReportItem, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vSectionNumber = vNodeInfo.SectionNumber
              vItemNumber = IntegerValue(dgr.GetValue(dgr.ActiveRow, "ItemNumber"))
              vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
              vList("SectionNumber") = vSectionNumber.ToString()
              vList("ItemNumber") = vItemNumber.ToString()
              DataHelper.DeleteReportItem(vList)

              ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)
              mvDataChanged = True
            End If
          End If

        Case TreeViewNodeType.Control
          Dim vSequenceNumber As Integer
          If dgr.RowCount > 0 Then
            If ShowQuestion(QuestionMessages.QmDeleteReportControl, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vSequenceNumber = IntegerValue(dgr.GetValue(dgr.ActiveRow, "SequenceNumber"))
              vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
              vList("ReportCode") = vNodeInfo.ReportCode.ToString()
              vList("SequenceNumber") = vSequenceNumber.ToString()
              DataHelper.DeleteReportControl(vList)

              ShowControl(vNodeInfo.ReportNumber, vNodeInfo.ReportCode)
            End If
          End If
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      Dim vNodeInfo As TreeViewNodeInfo
      Dim vCurrentTable As String = ""
      Dim vParams As New ParameterList(True)
      Dim vCriteria As New ParameterList(True)

      vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
      Select Case vNodeInfo.NodeType
        Case TreeViewNodeType.ReportType
          vCurrentTable = "Reports"
          vParams("MaintenanceTableName") = vCurrentTable
        Case TreeViewNodeType.Report
          vCurrentTable = "Reports"
          vParams("MaintenanceTableName") = vCurrentTable
        Case TreeViewNodeType.Parameter
          vCurrentTable = "report_parameters"
          vParams("MaintenanceTableName") = vCurrentTable
          vCriteria("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        Case TreeViewNodeType.Section
          vCurrentTable = "report_sections"
          vParams("MaintenanceTableName") = vCurrentTable
          vCriteria("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        Case TreeViewNodeType.SubSection
          vCurrentTable = "report_items"
          vParams("MaintenanceTableName") = vCurrentTable
          vCriteria("ReportNumber") = vNodeInfo.ReportNumber.ToString()
          vCriteria("SectionNumber") = vNodeInfo.SectionNumber.ToString()
        Case TreeViewNodeType.Version
          Dim vLastVersion As Integer = 1
          vCurrentTable = "report_version_history"
          vParams("MaintenanceTableName") = vCurrentTable
          vParams("Logname") = DataHelper.UserInfo.Logname
          vCriteria("ChangeDate") = Today.ToString
          vCriteria("ReportNumber") = vNodeInfo.ReportNumber.ToString()
          For vRowCounter As Integer = 0 To dgr.RowCount - 1
            If IntegerValue(dgr.GetValue(vRowCounter, 0)) > vLastVersion Then vLastVersion = IntegerValue(dgr.GetValue(vRowCounter, 0))
          Next
          vCriteria("VersionNumber") = (vLastVersion + 1).ToString
        Case TreeViewNodeType.Control
          vCurrentTable = "fp_controls"
          vParams("MaintenanceTableName") = vCurrentTable
          vCriteria("FpApplication") = vNodeInfo.ReportNumber.ToString()
          vCriteria("FpPageType") = vNodeInfo.ReportCode.ToString()
      End Select
      Dim vResult As DialogResult = DisplayTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTable, vParams, vCriteria)
      If vResult = System.Windows.Forms.DialogResult.OK Then
        mvDataChanged = True
        Select Case vNodeInfo.NodeType
          Case TreeViewNodeType.ReportType
            LoadReport()
            BuildNode(tvw.Nodes(vNodeInfo.ReportType - 1), vNodeInfo.ReportType, mvNewReportNumber)
            ShowReportData(tvw.SelectedNode)
          Case TreeViewNodeType.Report
            LoadReport()
            BuildNode(tvw.Nodes(vNodeInfo.ReportType - 1), vNodeInfo.ReportType, mvNewReportNumber)
            ShowIndividualReportData(vNodeInfo.ReportNumber)
          Case TreeViewNodeType.Parameter
            ShowParameterData(vNodeInfo.ReportNumber)
          Case TreeViewNodeType.Section
            LoadReportSection()
            ShowSectionData(vNodeInfo.ReportNumber)
            BuildSectionNode(tvw.SelectedNode, vNodeInfo.ReportNumber)
          Case TreeViewNodeType.SubSection
            ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)
          Case TreeViewNodeType.Version
            ShowVersion(vNodeInfo.ReportNumber)
            If dgr.RowCount > 0 Then cmdAmend.Enabled = True
          Case TreeViewNodeType.Control
            ShowControl(vNodeInfo.ReportNumber, vNodeInfo.ReportCode)
        End Select
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub



  Private Sub dgr_RowDoubleClicked(ByVal sender As Object, ByVal pRow As Integer) Handles dgr.RowDoubleClicked
    Try
      UpdateReportData()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub frmReportMaintenance_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      CheckForVersionHistory()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub UpdateCurrentNode(ByVal pCurrentNode As TreeNode)
    Dim vDataRowView As DataRowView
    Dim vNodeInfo As TreeViewNodeInfo = CType(pCurrentNode.Tag, TreeViewNodeInfo)
    mvReportSet.Tables("DataRow").DefaultView.RowFilter = "ReportNumber = '" + vNodeInfo.ReportNumber.ToString + "'"
    vDataRowView = mvReportSet.Tables("DataRow").DefaultView(0)
    If vNodeInfo.NodeType = TreeViewNodeType.Report Then
      tvw.SelectedNode.Text = vDataRowView("ReportName").ToString
      vNodeInfo = New TreeViewNodeInfo(TreeViewNodeType.Report, vNodeInfo.ReportType, vNodeInfo.ReportNumber, 0, vDataRowView("ReportCode").ToString)
      tvw.SelectedNode.Tag = vNodeInfo
    ElseIf vNodeInfo.NodeType = TreeViewNodeType.Section Then
      BuildSectionNode(tvw.SelectedNode, vNodeInfo.ReportNumber)
    End If
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    RemoveHandler tvw.AfterSelect, AddressOf tvw_AfterSelect
    Dim vNodeInfo As TreeViewNodeInfo
    vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
    If vNodeInfo.NodeType = 0 Then
      tvw.SelectedNode = tvw.SelectedNode.Nodes(dgr.GetValue(pRow, "ReportNumber"))
      InitialiseExtensions(CInt(dgr.GetValue(pRow, "ReportNumber")))
    ElseIf vNodeInfo.NodeType = 1 Then
      If dgr.GetValue(pRow, "ReportNumber").Length > 0 Then
        tvw.SelectedNode = tvw.SelectedNode.Parent.Nodes(dgr.GetValue(pRow, "ReportNumber"))
        InitialiseExtensions(vNodeInfo.ReportNumber)
      End If
    End If
    AddHandler tvw.AfterSelect, AddressOf tvw_AfterSelect
  End Sub

  Private Sub dgr_RowMoved(ByVal sender As Object, ByVal pOldRowNumber As Integer, ByVal pNewRowNumber As Integer) Handles dgr.RowMoved
    Dim vNodeInfo As TreeViewNodeInfo
    vNodeInfo = CType(tvw.SelectedNode.Tag, TreeViewNodeInfo)
    Dim vList As ParameterList = New ParameterList(True)

    Select Case vNodeInfo.NodeType
      Case TreeViewNodeType.SubSection
        vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        vList("SectionNumber") = vNodeInfo.SectionNumber.ToString()
        vList("SourceItem") = dgr.GetValue(pNewRowNumber, "ItemNumber")
        If pNewRowNumber = 0 Then
          vList("AfterItem") = "0"
        Else
          vList("AfterItem") = dgr.GetValue(pNewRowNumber - 1, "ItemNumber")
        End If
        DataHelper.RenumberItems(vList)
        ShowSubSectionDetail(vNodeInfo.ReportNumber, vNodeInfo.SectionNumber)

      Case TreeViewNodeType.Parameter
        vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        vList("SourceItem") = dgr.GetValue(pNewRowNumber, "ParameterNumber")
        If pNewRowNumber = 0 Then
          vList("AfterItem") = "0"
        Else
          vList("AfterItem") = dgr.GetValue(pNewRowNumber - 1, "ParameterNumber")
        End If
        DataHelper.RenumberParameters(vList)
        ShowParameterData(vNodeInfo.ReportNumber)

      Case TreeViewNodeType.Section
        vList("ReportNumber") = vNodeInfo.ReportNumber.ToString()
        vList("SourceItem") = dgr.GetValue(pNewRowNumber, "SectionNumber")
        If pNewRowNumber = 0 Then
          vList("AfterItem") = "0"
        Else
          vList("AfterItem") = dgr.GetValue(pNewRowNumber - 1, "SectionNumber")
        End If
        DataHelper.RenumberSections(vList)
        LoadReportSection()
        ShowSectionData(vNodeInfo.ReportNumber)
    End Select
  End Sub

  Private Sub InitialiseExtensions(pReportNumber As Integer)
    Dim vParams As New ParameterList(True, True)
    vParams.Add("ReportNumber", pReportNumber)
    Dim vExtensionList As XDocument = DataHelper.GetReportExtensions(vParams)
    bpl.SuspendLayout()
    ClearExtensionConfigButtons()
    If vExtensionList IsNot Nothing AndAlso vExtensionList.Elements("Extensions").Any() Then
      For Each vExtension As XElement In vExtensionList.Descendants("Extensions").Elements
        AddExtensionConfigurationButton(vExtension)
      Next
    End If
    bpl.ResumeLayout()
  End Sub

  Private Sub AddExtensionConfigurationButton(pExtension As XElement)
    If pExtension.Elements("PropertyPage").Any() AndAlso pExtension.Elements("Description").Any() Then
      Dim vButton As New Button
      Dim vLabel As String = String.Empty
      If pExtension.Descendants("SettingsLabel").Any() Then
        vLabel = pExtension.Element("SettingsLabel").Value
      ElseIf pExtension.Descendants("Description").Any() Then
        vLabel = pExtension.Element("Description").Value
      End If
      vButton.Text = vLabel
      AddHandler vButton.Click, AddressOf OnExtensionConfig_Clicked
      bpl.Controls.Add(vButton)
      ExtensionDictionary.Add(vButton, pExtension)
    End If
  End Sub

  Private Sub ClearExtensionConfigButtons()
    If Me.ExtensionDictionary IsNot Nothing AndAlso Me.ExtensionDictionary.Count > 0 Then
      For Each vEntry As KeyValuePair(Of Control, XElement) In ExtensionDictionary
        Dim vControl As Control = vEntry.Key
        If vControl IsNot Nothing AndAlso vControl.Parent IsNot Nothing Then
          vControl.Parent.Controls.Remove(vControl)
        End If
      Next
      Me.ExtensionDictionary.Clear()
    End If
  End Sub


  Private Sub OnExtensionConfig_Clicked(sender As Object, e As EventArgs)
    If TypeOf sender Is Control Then
      Dim vButton As Button = DirectCast(sender, Button)
      If Me.ExtensionDictionary.ContainsKey(vButton) Then
        Dim vExtension As XElement = Me.ExtensionDictionary(vButton)
        If vExtension IsNot Nothing Then
          If vExtension.Descendants("PropertyPage").Any() AndAlso vExtension.Descendants("SettingsKey").Any Then
            Dim vHtmlPage As String = vExtension.Element("PropertyPage").Value
            Dim vSettings As XDocument = XDocument.Parse(vExtension.Elements("Settings").DefaultIfEmpty(New XElement("Settings")).Value)
            Dim vOptions As New frmHtmlSettings.HtmlSettingsOptions
            vOptions.Html = vHtmlPage
            vOptions.SaveSettingsCommand = "saveSettings"
            vOptions.LoadSettingsCommand = "loadSettings"
            vOptions.Settings = vSettings
            Dim vSettingsDialog As New frmHtmlSettings(vOptions)
            Dim vResult As DialogResult = System.Windows.Forms.DialogResult.OK
            Do While vResult = System.Windows.Forms.DialogResult.OK
              vResult = vSettingsDialog.ShowDialog()
              If vResult = System.Windows.Forms.DialogResult.OK Then
                Try
                  If vOptions.Settings.ToString() <> vSettings.ToString() Then
                    DataHelper.ValidateExtensionSettings(vExtension.ToString(), vOptions.Settings.ToString())
                    DataHelper.SaveExtensionSettings(vExtension.Element("SettingsKey").Value, vOptions.Settings)
                    mvDataChanged = True
                    InitialiseExtensions(mvPrevReportNumber)
                    Exit Do
                  End If
                Catch ex As Exception
                  DataHelper.HandleException(ex)
                End Try
              End If
            Loop
          End If
        End If
      End If
    End If
  End Sub

End Class
Public Enum TreeViewNodeType
  ReportType
  Report
  Parameter
  Section
  SubSection
  Version
  Control
End Enum
Public Class TreeViewNodeInfo
  Private mvNodeType As TreeViewNodeType
  Private mvReportTypeId As Integer = 0
  Private mvReportNumber As Integer = 0
  Private mvSectionNumber As Integer = 0
  Private mvReportCode As String = ""

  Public Sub New(ByVal pNodeType As TreeViewNodeType, Optional ByVal pReportTypeId As Integer = 0, Optional ByVal pReportNumber As Integer = 0, Optional ByVal pSectionNumber As Integer = 0, Optional ByVal pReportCode As String = "")
    mvNodeType = pNodeType
    mvReportTypeId = pReportTypeId
    mvReportNumber = pReportNumber
    mvSectionNumber = pSectionNumber
    mvReportCode = pReportCode
  End Sub
  Public ReadOnly Property NodeType() As Integer
    Get
      Return mvNodeType
    End Get
  End Property
  Public ReadOnly Property ReportType() As Integer
    Get
      Return mvReportTypeId
    End Get
  End Property
  Public ReadOnly Property ReportNumber() As Integer
    Get
      Return mvReportNumber
    End Get
  End Property
  Public ReadOnly Property SectionNumber() As Integer
    Get
      Return mvSectionNumber
    End Get
  End Property
  Public ReadOnly Property ReportCode() As String
    Get
      Return mvReportCode
    End Get
  End Property

End Class