Public Class frmOwnershipMaintenance

  Private mvOwnershipSet As DataSet
  Private WithEvents mvOwnershipGroupMenu As OwnershipMenu

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Private Sub InitialiseControls()
    Try
      mvOwnershipGroupMenu = New OwnershipMenu()
      SetControlTheme()
      LoadOwnershipData()
      GetOwnershipTree()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub
#Region "Private Methods"
  ''' <summary>
  ''' Function to build initial ownership nodes
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetOwnershipTree()
    Dim vNodeInfo As OwnershipNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.OwnershipGroupType, "", "", "", "")
    Dim vNode As TreeNode = New TreeNode("OwnershipGroup")
    tvw.Nodes.Clear()
    Dim vMasterNode As TreeNode = tvw.Nodes.Add("OwnershipGroups", "Ownership Groups")
    vMasterNode.Tag = vNodeInfo
    ' Get all the ownership groups
    Dim vDataTable As DataTable = mvOwnershipSet.Tables("DataRow")
    If Not vDataTable Is Nothing AndAlso vDataTable.Rows.Count > 0 Then
      For Each vRow As DataRow In vDataTable.Rows
        'Adding OwnershipGroup Node
        vNode = vMasterNode.Nodes.Add(vRow("OwnershipGroup").ToString(), vRow.Item("OwnershipGroupDesc").ToString)
        vNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.OwnershipGroup, vRow("OwnershipGroup").ToString())
        vNode.Tag = vNodeInfo
        'Adding Users in OwnershipGroup Node
        vNode.Nodes.Add("Users", "Users")
        vNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.UserType, vRow("OwnershipGroup").ToString())
        vNode.Nodes("Users").Tag = vNodeInfo
        vNode.Nodes("Users").Nodes.Add("_DUMMY", "_DUMMY")
        'Adding Department in OwnershipGroup Node
        vNode.Nodes.Add("Departments", "Department Defaults")
        vNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.DepartmentType, vRow("OwnershipGroup").ToString())
        vNode.Nodes("Departments").Tag = vNodeInfo
        vNode.Nodes("Departments").Nodes.Add("_DUMMY", "_DUMMY")
      Next
    End If
    vMasterNode.Expand()
    dgr.Populate(mvOwnershipSet)
    DataHelper.SetConfigValue("ownership_method", "G", True)
  End Sub

  Private Sub LoadOwnershipData()
    Dim vList As New ParameterList(True)
    mvOwnershipSet = DataHelper.GetOwnershipGroupInformation(vList)
  End Sub
  ''' <summary>
  ''' Displaying child node of currently selected node
  ''' </summary>
  ''' <param name="pNode"></param>
  ''' <remarks></remarks>
  Private Sub DisplayChildNode(ByVal pNode As TreeNode)
    Dim vNodeInfo As OwnershipNodeInfo
    Dim vNewNodeInfo As OwnershipNodeInfo
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True)
    vNodeInfo = CType(pNode.Tag, OwnershipNodeInfo)

    Select Case vNodeInfo.NodeType
      Case OwnershipNodeType.UserType
        If pNode.Nodes.Find("_DUMMY", False).Length <> 0 Then
          pNode.Nodes.Clear()
          vList("OwnershipGroup") = vNodeInfo.OwnershipGroup
          vDataSet = DataHelper.GetOwnershipUsers(vList)
          If Not vDataSet.Tables("DataRow") Is Nothing Then
            For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
              Dim vNode As TreeNode = New TreeNode("User")
              vNewNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.User, vRow("OwnershipGroup").ToString(), vRow("Logname").ToString())
              pNode.Nodes.Add(vRow("Logname").ToString(), vRow("Logname").ToString() + "(" + vRow("FullName").ToString() + ")")
              pNode.Nodes(vRow("Logname").ToString()).Tag = vNewNodeInfo
            Next
          End If
          If pNode.Nodes.Count > 0 Then
            If vNodeInfo.NodeType = OwnershipNodeType.UserType Then
              pNode.Expand()
            End If
          End If
        End If
      Case OwnershipNodeType.DepartmentType
        If pNode.Nodes.Find("_DUMMY", False).Length <> 0 Then
          pNode.Nodes.Clear()
          vList("OwnershipGroup") = pNode.Parent.Name
          vDataSet = DataHelper.GetOwnershipDepartmentInformation(vList)
          If Not vDataSet.Tables("DataRow") Is Nothing Then
            For Each vRow As DataRow In vDataSet.Tables("DataRow").Rows
              Dim vNode As TreeNode = New TreeNode("Department")
              vNewNodeInfo = New OwnershipNodeInfo(OwnershipNodeType.Department, vNodeInfo.OwnershipGroup, "", vRow("Department").ToString(), vRow("OwnershipAccessLevel").ToString())
              pNode.Nodes.Add(vRow("Department").ToString(), vRow("DepartmentDesc").ToString())
              pNode.Nodes(vRow("Department").ToString()).Tag = vNewNodeInfo
            Next
          End If
          If pNode.Nodes.Count > 0 Then
            If vNodeInfo.NodeType = OwnershipNodeType.DepartmentType Then
              pNode.Expand()
            End If
          End If
        End If
    End Select

  End Sub
  ''' <summary>
  ''' Show corrosponding grid data
  ''' </summary>
  ''' <param name="pnode"></param>
  ''' <remarks></remarks>
  Private Sub ShowOwnershipData(ByVal pnode As TreeNode)
    Dim vNodeInfo As OwnershipNodeInfo
    Dim vDataSet As New DataSet
    Dim vList As New ParameterList(True)
    vNodeInfo = CType(pnode.Tag, OwnershipNodeInfo)

    Select Case vNodeInfo.NodeType
      Case OwnershipNodeType.DepartmentType
        vList("OwnershipGroup") = vNodeInfo.OwnershipGroup
        vDataSet = DataHelper.GetOwnershipDepartmentInformation(vList)
        dgr.SetColumnHeaderVisible(True)
        dgr.Populate(vDataSet)
        dgr.SetRowHeaderColumnVisible(0, False)
        ShowButtons(False)
      Case OwnershipNodeType.Department
        vList("OwnershipGroup") = vNodeInfo.OwnershipGroup
        vList("Department") = vNodeInfo.Department
        vDataSet = DataHelper.GetOwnershipDepartmentInformation(vList)
        dgr.SetColumnHeaderVisible(True)
        dgr.Populate(vDataSet)
        dgr.SetRowHeaderColumnVisible(0, False)
        ShowButtons(False)
      Case OwnershipNodeType.UserType
        vList("OwnershipGroup") = vNodeInfo.OwnershipGroup
        vDataSet = DataHelper.GetOwnershipUserInformation(vList)
        dgr.SetColumnHeaderVisible(True)
        dgr.Populate(vDataSet)
        dgr.SetRowHeaderColumnVisible(0, False)
        dgr.SetColumnVisible("ValidFrom", False)
        dgr.SetColumnVisible("ValidTo", False)
        dgr.SetColumnVisible("AmendedOn", False)
        dgr.SetColumnVisible("AmendedBy", False)
        ShowButtons(False)
      Case OwnershipNodeType.User
        vList("OwnershipGroup") = vNodeInfo.OwnershipGroup
        vList("Logname") = vNodeInfo.User
        vDataSet = DataHelper.GetOwnershipUserInformation(vList)
        dgr.SetColumnHeaderVisible(True)
        dgr.Populate(vDataSet)
        dgr.SetRowHeaderColumnVisible(0, False)
        ShowButtons(False)
      Case OwnershipNodeType.OwnershipGroup
        LoadOwnershipData()
        Dim vNewRow As DataRow
        Dim vDataTable As New DataTable("DataRow")
        vDataTable.Columns.Add()
        Dim vRow As DataRow()
        vRow = mvOwnershipSet.Tables("DataRow").Select("OwnershipGroup  = '" + vNodeInfo.OwnershipGroup + "'")
        If vRow.Length > 0 Then
          For intCtr As Integer = 0 To vRow(0).ItemArray.Length - 1
            vNewRow = vDataTable.NewRow()
            vNewRow(0) = vRow(0).Item(intCtr).ToString()
            vDataTable.Rows.Add(vNewRow)
          Next
          vDataSet.Tables.Add(vDataTable)
          dgr.Populate(vDataSet)
          dgr.SetColumnHeaderVisible(False)
          dgr.SetRowHeaderColumnVisible(0, True)
          dgr.SetRowVisible(3, False)
          dgr.SetRowHeaderVisible()
          For intIncr As Integer = 0 To mvOwnershipSet.Tables("Column").Rows.Count - 1
            dgr.SetRowHeaderValue(intIncr, 0, mvOwnershipSet.Tables("Column").Rows(intIncr).Item("Heading").ToString())
          Next
          dgr.SetPreferredRowHeaderWidth(0)
        End If
        ShowButtons(True)
      Case OwnershipNodeType.OwnershipGroupType
        LoadOwnershipData()
        dgr.Populate(mvOwnershipSet)
        dgr.SetColumnHeaderVisible(True)
        dgr.SetRowHeaderColumnVisible(0, False)
        ShowButtons(False)
    End Select
  End Sub

  Private Sub ShowButtons(ByVal pEnable As Boolean)
    cmdAmend.Enabled = pEnable
    cmdNew.Enabled = pEnable
    cmdDelete.Enabled = pEnable
  End Sub

#End Region

  Private Sub tvw_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterSelect
    Try
      ShowOwnershipData(e.Node)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub tvw_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvw.BeforeExpand
    Try
      DisplayChildNode(e.Node)
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
          tvw.ContextMenuStrip = mvOwnershipGroupMenu
          If tvw.ContextMenuStrip IsNot Nothing Then
            Dim vNodeInfo As OwnershipNodeInfo = CType(vNode.Tag, OwnershipNodeInfo)
            mvOwnershipGroupMenu.OwnershipNodeInfo = vNodeInfo
          End If
        Else
          tvw.ContextMenuStrip = Nothing
        End If
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' This is an event method which fires after menu action completed
  ''' </summary>
  ''' <param name="pItem"></param>
  ''' <remarks></remarks>
  Private Sub mvOwnershipGroupMenu_MenuActionCompleted(ByVal pItem As OwnershipMenu.OwnershipMenuItems) Handles mvOwnershipGroupMenu.MenuActionCompleted
    Try
      Dim vNodeInfo As OwnershipNodeInfo
      vNodeInfo = CType(tvw.SelectedNode.Tag, OwnershipNodeInfo)
      Select Case pItem
        Case OwnershipMenu.OwnershipMenuItems.omiChangeDefault
          ShowOwnershipData(tvw.SelectedNode)
        Case OwnershipMenu.OwnershipMenuItems.omiChangeAccess
          ShowOwnershipData(tvw.SelectedNode)
        Case OwnershipMenu.OwnershipMenuItems.omiNewGroup
          LoadOwnershipData()
          GetOwnershipTree()
        Case OwnershipMenu.OwnershipMenuItems.omiAddDepartment
          If vNodeInfo.NodeType = OwnershipNodeType.OwnershipGroup Then
            tvw.SelectedNode = tvw.SelectedNode.Nodes("Departments")
            tvw.SelectedNode.Expand()
          ElseIf vNodeInfo.NodeType = OwnershipNodeType.UserType Then
            tvw.SelectedNode = tvw.SelectedNode.Parent.Nodes("Departments")
            tvw.SelectedNode.Expand()
          End If
          tvw.SelectedNode.Nodes.Add("_DUMMY", "_DUMMY") ' Force DisplayNode function to rebuild node
          DisplayChildNode(tvw.SelectedNode)
          ShowOwnershipData(tvw.SelectedNode)
        Case OwnershipMenu.OwnershipMenuItems.omiAddUser
          If vNodeInfo.NodeType = OwnershipNodeType.OwnershipGroup Then
            tvw.SelectedNode = tvw.SelectedNode.Nodes("Users")
            tvw.SelectedNode.Expand()
          ElseIf vNodeInfo.NodeType = OwnershipNodeType.DepartmentType Then
            tvw.SelectedNode = tvw.SelectedNode.Parent.Nodes("Users")
            tvw.SelectedNode.Expand()
          End If
          tvw.SelectedNode.Nodes.Add("_DUMMY", "_DUMMY") ' Force DisplayNode function to rebuild node
          DisplayChildNode(tvw.SelectedNode)
          ShowOwnershipData(tvw.SelectedNode)
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdAmend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAmend.Click
    Try
      Dim vCurrentTableName As String
      Dim vParams As New ParameterList(True)
      Dim vCriteriaList As New ParameterList

      vCurrentTableName = "ownership_groups"
      vParams("MaintenanceTableName") = vCurrentTableName
      vCriteriaList("OwnershipGroup") = dgr.GetValue(0, 0)
      vParams("OwnershipGroup") = dgr.GetValue(0, 0)
      vParams("PrincipalDepartment") = dgr.GetValue(2, 0)
      vParams("OwnershipGroupDesc") = dgr.GetValue(1, 0)
      vParams("PrincipalDepartmentLogname") = dgr.GetValue(4, 0)
      vParams("ReadAccessText") = dgr.GetValue(6, 0) 'BR16293
      vParams("ViewAccessText") = dgr.GetValue(7, 0) 'BR16293
      vParams("Notes") = dgr.GetValue(8, 0) 'BR16293
      Dim vForm As New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmAmend, vCurrentTableName, vParams, vCriteriaList)
      vForm.Text = ControlText.FrmOwnershipGroups
      vForm.ShowDialog()
      ShowOwnershipData(tvw.SelectedNode)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Me.Close()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      If ConfirmDelete() Then
        Dim vParams As New ParameterList(True)
        vParams("MaintenanceTableName") = "ownership_groups"
        vParams("OwnershipGroup") = dgr.GetValue(0, 0)
        DataHelper.DeleteTableMaintenanceData(vParams)
        LoadOwnershipData()
        GetOwnershipTree()
      End If
    Catch vCareEx As CareException
      If vCareEx.ErrorNumber = CareException.ErrorNumbers.enCannotDelete Then
        ShowErrorMessage(vCareEx.Message)
      Else
        DataHelper.HandleException(vCareEx)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub cmdNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Try
      Dim vCurrentTableName As String = "ownership_groups"
      Dim vParams As New ParameterList(True)
      vParams("MaintenanceTableName") = vCurrentTableName
      Dim vForm As New frmTableEntry(CareNetServices.XMLTableMaintenanceMode.xtmmNew, vCurrentTableName, vParams, Nothing)
      vForm.Text = ControlText.FrmOwnershipGroups
      If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        LoadOwnershipData()
        GetOwnershipTree()
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub frmOwnershipMaintenance_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    ObjectCache.ClearCache(Of CDBNETBiz.OwnershipGroup)()
    ObjectCache.ClearCache(Of CDBNETBiz.OwnershipUserInfo)()
  End Sub
End Class
''' <summary>
''' Enum representing node type
''' </summary>
''' <remarks></remarks>
Public Enum OwnershipNodeType
  OwnershipGroupType
  OwnershipGroup
  UserType
  User
  DepartmentType
  Department
End Enum
''' <summary>
''' Class to store tag level information
''' </summary>
''' <remarks></remarks>
Public Class OwnershipNodeInfo
  Private mvNodeType As OwnershipNodeType
  Private mvOwnershipGroup As String
  Private mvUser As String
  Private mvDepartment As String
  Private mvAccessLevel As String

  Public Sub New(ByVal pNodeType As OwnershipNodeType, Optional ByVal pOwnershipGroup As String = "", Optional ByVal pUser As String = "", Optional ByVal pDepartment As String = "", Optional ByVal pAccessLevel As String = "")
    mvNodeType = pNodeType
    mvOwnershipGroup = pOwnershipGroup
    mvUser = pUser
    mvDepartment = pDepartment
    mvAccessLevel = pAccessLevel
  End Sub

  Public ReadOnly Property NodeType() As Integer
    Get
      Return mvNodeType
    End Get
  End Property
  Public ReadOnly Property OwnershipGroup() As String
    Get
      Return mvOwnershipGroup
    End Get
  End Property
  Public ReadOnly Property User() As String
    Get
      Return mvUser
    End Get
  End Property
  Public ReadOnly Property Department() As String
    Get
      Return mvDepartment
    End Get
  End Property
  Public ReadOnly Property AcceesLevel() As String
    Get
      Return mvAccessLevel
    End Get
  End Property

End Class