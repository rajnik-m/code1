Public Class TabSelector
  Inherits System.Windows.Forms.UserControl

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
  Friend WithEvents tvw As System.Windows.Forms.TreeView
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.tvw = New System.Windows.Forms.TreeView
    Me.SuspendLayout()
    '
    'tvw
    '
    Me.tvw.BorderStyle = System.Windows.Forms.BorderStyle.None
    Me.tvw.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tvw.HideSelection = False
    Me.tvw.Location = New System.Drawing.Point(0, 0)
    Me.tvw.Name = "tvw"
    Me.tvw.Size = New System.Drawing.Size(216, 328)
    Me.tvw.TabIndex = 0
    '
    'TabSelector
    '
    Me.Controls.Add(Me.tvw)
    Me.Name = "TabSelector"
    Me.Size = New System.Drawing.Size(216, 328)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Public Event ContactTabSelected(ByVal sender As Object, ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pCustomForm As Integer, ByVal pReadOnlyPage As Boolean)
  Public Event CampaignTabSelected(ByVal sender As Object, ByVal pType As CareServices.XMLCampaignDataSelectionTypes, ByVal pCampaignItem As CampaignItem)
  Public Event EventTabSelected(ByVal sender As Object, ByVal pType As CareServices.XMLEventDataSelectionTypes, ByVal pCode As String)
  Public Event BeforeSelect(ByVal sender As Object, ByRef pCancel As Boolean)

  Private mvBackgroundExtender As BackGroundExtender
  Private mvCampaignName As String
  Private mvCampaignRestrictions As ParameterList

  Public Sub InitialiseControls()
    tvw.BackColor = DisplayTheme.SelectionPanelTreeBackColor
    SetStyle(System.Windows.Forms.ControlStyles.DoubleBuffer, True)
    SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, False)
    SetStyle(System.Windows.Forms.ControlStyles.ResizeRedraw, True)
    SetStyle(System.Windows.Forms.ControlStyles.UserPaint, True)
    SetStyle(ControlStyles.SupportsTransparentBackColor, True)
    Me.BackColor = System.Drawing.Color.Transparent
    mvBackgroundExtender = New BackGroundExtender(BackGroundExtender.BackgroundExtenderTypes.betSelectionPanel)
    tvw.Nodes.Clear()
  End Sub

  Public Sub Init(ByVal pCampaignItem As CampaignItem, Optional ByVal pSelectItem As CampaignItem = Nothing, Optional ByVal pImageList As ImageList = Nothing)
    With tvw
      If .TopNode IsNot Nothing Then tvw.Nodes.Clear()
      Dim vList As New ParameterList(True)
      vList("Campaign") = pCampaignItem.Campaign
      Dim vCampaignNode As TreeNode = Nothing
      Dim vItem As SelectionItem
      Dim vNodeSelected As Boolean
      If pImageList IsNot Nothing Then
        tvw.ImageList = pImageList
        tvw.ImageKey = "citgen"
        tvw.SelectedImageKey = "citgen"
        'tvw.ImageIndex = 0
        'tvw.SelectedImageIndex = 0
      End If
      If pCampaignItem.Campaign.Length = 0 Then
        vItem = New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCampaign, "")
        vCampaignNode = NewNode(vItem)
        vCampaignNode.SelectedImageKey = [Enum].GetName(GetType(CampaignItem.CampaignItemTypes), CampaignItem.CampaignItemTypes.citCampaign)
        vCampaignNode.ImageKey = [Enum].GetName(GetType(CampaignItem.CampaignItemTypes), CampaignItem.CampaignItemTypes.citCampaign)
        tvw.Nodes.Add(vCampaignNode)
        mvCampaignName = vItem.Description
      Else
        If mvCampaignRestrictions IsNot Nothing Then
          With mvCampaignRestrictions
            For Each vValue As DictionaryEntry In mvCampaignRestrictions
              vList.Add(vValue.Key.ToString, vValue.Value.ToString)
            Next
          End With
        End If
        Dim vSelectionPages As DataTable = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtCampaignPages, vList)
        Dim vAppealNode As TreeNode = Nothing
        Dim vCollectionNode As TreeNode = Nothing
        Dim vNewNode As TreeNode = Nothing
        For Each vRow As DataRow In vSelectionPages.Rows
          vItem = New SelectionItem(vRow, SelectionItem.SelectionItemTypes.sitCampaignSelection)
          If vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtCampaign Then
            vNewNode = NewNode(vItem)
            vCampaignNode = vNewNode
            vCampaignNode.SelectedImageKey = [Enum].GetName(GetType(CampaignItem.CampaignItemTypes), CampaignItem.CampaignItemTypes.citCampaign)
            vCampaignNode.ImageKey = [Enum].GetName(GetType(CampaignItem.CampaignItemTypes), CampaignItem.CampaignItemTypes.citCampaign)
            tvw.Nodes.Add(vCampaignNode)
            mvCampaignName = vItem.Description
            Dim vCode As String = vRow.Item("Code").ToString
            vCampaignNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers, vCode)))
          ElseIf vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal Then
            vNewNode = NewNode(vItem)
            vAppealNode = vNewNode
            vAppealNode.ImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vAppealNode.SelectedImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vCampaignNode.Nodes.Add(vAppealNode)
            Dim vCode As String = vRow.Item("Code").ToString
            If vItem.CampaignItem.AppealType = CampaignItem.AppealTypes.atSegment Then
              vAppealNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets, vCode)))
              vAppealNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers, vCode)))
            ElseIf (vItem.CampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection _
               Or vItem.CampaignItem.AppealType = CampaignItem.AppealTypes.atUnMannedCollection) Then
              vAppealNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources, vCode)))
            End If
          ElseIf vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtSegment Then
            vNewNode = NewNode(vItem)
            vNewNode.ImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vNewNode.SelectedImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vAppealNode.Nodes.Add(vNewNode)
            Dim vCode As String = vRow.Item("Code").ToString
            vNewNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCostCentres, vCode)))
            vNewNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtTickBoxes, vCode)))
            vNewNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentProducts, vCode)))
            vNewNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers, vCode)))
          ElseIf vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtCollection Then
            vNewNode = NewNode(vItem)
            vCollectionNode = vNewNode
            vCollectionNode.ImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vCollectionNode.SelectedImageKey = [Enum].GetName(GetType(CampaignItem.AppealTypes), vItem.CampaignItem.AppealType)
            vAppealNode.Nodes.Add(vCollectionNode)
            Dim vCode As String = vRow.Item("Code").ToString
            Select Case vItem.CampaignItem.AppealType
              Case CampaignItem.AppealTypes.atH2HCollection
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectors, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectionPIS, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments, vCode)))
              Case CampaignItem.AppealTypes.atMannedCollection
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectionBoxes, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments, vCode)))
              Case CampaignItem.AppealTypes.atUnMannedCollection
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtUnMannedCollectionBoxes, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources, vCode)))
                vCollectionNode.Nodes.Add(NewNode(New SelectionItem(CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments, vCode)))
            End Select
          End If
          If pSelectItem IsNot Nothing AndAlso vItem.CampaignItem.Code = pSelectItem.Code Then
            tvw.SelectedNode = vNewNode
            vNodeSelected = True
          End If
        Next
        If vNodeSelected = False AndAlso vCampaignNode IsNot Nothing Then tvw.SelectedNode = vCampaignNode
        If vCampaignNode IsNot Nothing Then vCampaignNode.Expand()
      End If
    End With
  End Sub
  Public Sub Init(ByVal pContactInfo As ContactInfo)
    With tvw
      If .TopNode Is Nothing Then
        Dim vSelectionPages As DataTable = DataHelper.ContactAndOrganisationGroups(pContactInfo.ContactGroup).SelectionPages
        Dim vNode As TreeNode = Nothing
        Dim vItem As SelectionItem
        For Each vRow As DataRow In vSelectionPages.Rows
          vItem = New SelectionItem(vRow, SelectionItem.SelectionItemTypes.sitContactSelection)
          If vItem.ContactSelectionType = CareServices.XMLContactDataSelectionTypes.xcdtNone Then
            vNode = NewNode(vItem)
            tvw.Nodes.Add(vNode)
          Else
            vNode.Nodes.Add(NewNode(vItem))
          End If
        Next
      End If
    End With
  End Sub
  Public Sub Init(ByVal pEventInfo As CareEventInfo)
    With tvw
      If .TopNode Is Nothing Then
        Dim vSelectionPages As DataTable = DataHelper.EventGroups(pEventInfo.EventGroup).SelectionPages
        Dim vNode As TreeNode = Nothing
        Dim vItem As SelectionItem
        For Each vRow As DataRow In vSelectionPages.Rows
          vItem = New SelectionItem(vRow, SelectionItem.SelectionItemTypes.sitEventSelection)
          If vItem.EventSelectionType = -1 Then
            vNode = NewNode(vItem)
            tvw.Nodes.Add(vNode)
          Else
            If vNode IsNot Nothing Then
              vNode.Nodes.Add(NewNode(vItem))
            Else
              tvw.Nodes.Add(NewNode(vItem))
            End If
          End If
        Next
      End If
    End With
  End Sub

  Public Sub AddNode(ByVal pType As CareServices.XMLCampaignDataSelectionTypes, ByVal pCode As String)
    If tvw.SelectedNode IsNot Nothing Then
      Dim vItem As SelectionItem = DirectCast(tvw.SelectedNode.Tag, SelectionItem)
      Dim vCode As String = vItem.Code
      vItem = New SelectionItem(pType, pCode & "_")
      Dim vNode As TreeNode = NewNode(vItem)
      tvw.SelectedNode.Nodes.Add(vNode)
      tvw.SelectedNode = vNode
    End If
  End Sub

  Public Sub RemoveSelectedNode()
    If tvw.SelectedNode IsNot Nothing Then tvw.Nodes.Remove(tvw.SelectedNode)
  End Sub

  Public Sub SetSelectionType(ByRef pType As CareServices.XMLContactDataSelectionTypes)
    Dim vNode As TreeNode = tvw.Nodes(0)
    vNode = FindNode(vNode, pType)
    If Not vNode Is Nothing Then
      Dim vItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
      pType = vItem.ContactSelectionType
      tvw.SelectedNode = vNode
      tvw.SelectedNode.Expand()
      vNode.EnsureVisible()
    End If
  End Sub
  Public Sub SetSelectionType(ByRef pType As CareServices.XMLEventDataSelectionTypes)
    Dim vNode As TreeNode = tvw.Nodes(0)
    vNode = FindNode(vNode, pType)
    If Not vNode Is Nothing Then
      Dim vItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
      pType = vItem.EventSelectionType
      tvw.SelectedNode = vNode
      tvw.SelectedNode.Expand()
      vNode.EnsureVisible()
    End If
  End Sub

  Public Sub SetSelectionType(ByRef pType As CareServices.XMLCampaignDataSelectionTypes)
    Dim vNode As TreeNode = tvw.SelectedNode
    vNode = FindNode(vNode, pType)
    If Not vNode Is Nothing Then
      Dim vItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
      pType = vItem.CampaignSelectionType
      tvw.SelectedNode = vNode
      tvw.SelectedNode.Expand()
      vNode.EnsureVisible()
    End If
  End Sub

  Public ReadOnly Property SelectedContactDataType() As CareServices.XMLContactDataSelectionTypes
    Get
      Dim vNode As TreeNode = tvw.SelectedNode
      If vNode IsNot Nothing Then
        Dim vItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
        Return vItem.ContactSelectionType
      End If
    End Get
  End Property

  Public Sub SetFocus()
    tvw.Focus()
  End Sub

  Public Function HasDependants() As Boolean
    ' this is specific to campaign structures at the moment
    Dim vNode As TreeNode = tvw.SelectedNode
    Dim vMainItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
    If vNode.Nodes.Count > 0 Then
      Dim vChildNode As TreeNode
      For Each vChildNode In vNode.Nodes
        Dim vItem As SelectionItem = DirectCast(vChildNode.Tag, SelectionItem)
        If vItem.CampaignItem.ItemType <> vMainItem.CampaignItem.ItemType Then
          Return True
        End If
      Next
    End If
    Return False
  End Function

  Public Function MaxSequenceNumber() As Integer
    Dim vMaxSequenceNumber As Integer
    Dim vNode As TreeNode = AppealNode()
    Dim vItem As SelectionItem
    If vNode IsNot Nothing Then
      For Each vSegmentNode As TreeNode In vNode.Nodes
        vItem = DirectCast(vSegmentNode.Tag, SelectionItem)
        If vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtSegment Then
          vMaxSequenceNumber = Math.Max(vMaxSequenceNumber, vItem.Sequence)
        End If
      Next
    End If
    Return vMaxSequenceNumber
  End Function

  Private Function AppealNode() As TreeNode
    Dim vNode As TreeNode = tvw.SelectedNode
    If vNode IsNot Nothing Then
      Dim vItem As SelectionItem = DirectCast(vNode.Tag, SelectionItem)
      Select Case vItem.CampaignSelectionType
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal
          Return vNode                'It is this node
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegment
          Return vNode.Parent         'This nodes parent
      End Select
    End If
    Return vNode
  End Function

  Public Function NewSequenceValid(ByVal pSequenceNumber As Integer) As Boolean
    Dim vNode As TreeNode = AppealNode()
    Dim vItem As SelectionItem
    For Each vSegmentNode As TreeNode In vNode.Nodes
      vItem = DirectCast(vSegmentNode.Tag, SelectionItem)
      If vItem.CampaignSelectionType = CareServices.XMLCampaignDataSelectionTypes.xcadtSegment Then
        If pSequenceNumber = vItem.Sequence AndAlso vSegmentNode IsNot tvw.SelectedNode Then Return False
      End If
    Next
    Return True
  End Function

  Public ReadOnly Property Caption() As String
    Get
      Return mvCampaignName
    End Get
  End Property

  Public Property TreeContextMenu() As ContextMenuStrip
    Get
      Return tvw.ContextMenuStrip
    End Get
    Set(ByVal value As ContextMenuStrip)
      tvw.ContextMenuStrip = value
    End Set
  End Property

  Private Function FindNode(ByVal pNode As TreeNode, ByVal pType As CareServices.XMLContactDataSelectionTypes) As TreeNode
    Dim vNode As TreeNode
    Dim vFindNode As TreeNode
    Dim vItem As SelectionItem

    For Each vNode In pNode.Nodes
      vItem = DirectCast(vNode.Tag, SelectionItem)
      If vItem.ContactSelectionType = pType OrElse ((pType = CareServices.XMLContactDataSelectionTypes.xcdtNone) And (vItem.ContactSelectionType <> CareServices.XMLContactDataSelectionTypes.xcdtNone)) Then
        Return vNode
        Exit For
      Else
        vFindNode = FindNode(vNode, pType)
        If Not vFindNode Is Nothing Then
          Return vFindNode
          Exit For
        End If
      End If
    Next
    Return Nothing
  End Function
  Private Function FindNode(ByVal pNode As TreeNode, ByVal pType As CareServices.XMLEventDataSelectionTypes) As TreeNode
    Dim vNode As TreeNode
    Dim vFindNode As TreeNode
    Dim vItem As SelectionItem

    For Each vNode In pNode.Nodes
      vItem = DirectCast(vNode.Tag, SelectionItem)
      If vItem.EventSelectionType = pType OrElse ((pType = -1) And (vItem.EventSelectionType <> -1)) Then
        Return vNode
        Exit For
      Else
        vFindNode = FindNode(vNode, pType)
        If Not vFindNode Is Nothing Then
          Return vFindNode
          Exit For
        End If
      End If
    Next
    Return Nothing
  End Function

  Private Function FindNode(ByVal pNode As TreeNode, ByVal pType As CareServices.XMLCampaignDataSelectionTypes) As TreeNode
    Dim vNode As TreeNode
    Dim vFindNode As TreeNode
    Dim vItem As SelectionItem

    For Each vNode In pNode.Nodes
      vItem = DirectCast(vNode.Tag, SelectionItem)
      If vItem.CampaignSelectionType = pType OrElse ((pType = -1) And (vItem.CampaignSelectionType <> -1)) Then
        Return vNode
        Exit For
      Else
        vFindNode = FindNode(vNode, pType)
        If Not vFindNode Is Nothing Then
          Return vFindNode
          Exit For
        End If
      End If
    Next
    Return Nothing
  End Function

  Private Function NewNode(ByVal pItem As SelectionItem) As TreeNode
    Dim vNode As New TreeNode(pItem.Description)
    vNode.Tag = pItem
    Return vNode
  End Function

  Private Sub tvw_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvw.AfterSelect
    Dim vItem As SelectionItem
    Dim vCustomForm As Integer = 0

    vItem = DirectCast(e.Node.Tag, SelectionItem)
    If vItem.ContactSelectionType <> CareServices.XMLContactDataSelectionTypes.xcdtNone Then
      If vItem.Code.StartsWith("CustomForm") Then vCustomForm = CInt(vItem.Code.Substring(10))
      RaiseEvent ContactTabSelected(Me, vItem.ContactSelectionType, vCustomForm, vItem.ReadOnlyPage)
      tvw.Focus()
    ElseIf vItem.CampaignSelectionType <> CareServices.XMLCampaignDataSelectionTypes.xcadtNone Then
      RaiseEvent CampaignTabSelected(Me, vItem.CampaignSelectionType, vItem.CampaignItem)
    ElseIf vItem.EventSelectionType <> -1 Then
      RaiseEvent EventTabSelected(Me, vItem.EventSelectionType, vItem.Code)
    End If
  End Sub
  Private Sub tvw_BeforeSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles tvw.BeforeSelect
    Dim vCancel As Boolean
    RaiseEvent BeforeSelect(Me, vCancel)
    e.Cancel = vCancel
  End Sub
  Private Sub tvw_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvw.MouseDown
    If e.Button = Windows.Forms.MouseButtons.Right Then
      Dim vNode As TreeNode = tvw.GetNodeAt(e.X, e.Y)
      tvw.SelectedNode = vNode
      Dim vItem As SelectionItem = Nothing
      If vNode IsNot Nothing Then
        tvw.SelectedNode.Expand()
        If vNode.Tag IsNot Nothing Then
          vItem = DirectCast(vNode.Tag, SelectionItem)
        End If
      End If
      If tvw.ContextMenuStrip IsNot Nothing AndAlso vItem IsNot Nothing Then
        If TypeOf tvw.ContextMenuStrip Is CampaignMenu Then
          DirectCast(tvw.ContextMenuStrip, CampaignMenu).CampaignItem = vItem.CampaignItem
        ElseIf TypeOf tvw.ContextMenuStrip Is EventMenu Then
          DirectCast(tvw.ContextMenuStrip, EventMenu).EventDataType = vItem.EventSelectionType
        End If
      End If
    End If
  End Sub

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    Me.BackColor = System.Drawing.Color.Transparent
    MyBase.OnPaintBackground(e)
    mvBackgroundExtender.PaintBackGround(Me.ClientRectangle, e)
  End Sub

  Private Class SelectionItem
    Public Description As String
    Public Code As String
    Public Sequence As Integer
    Public ReadOnlyPage As Boolean
    Public CampaignItem As CampaignItem
    Public ContactSelectionType As CareServices.XMLContactDataSelectionTypes
    Public CampaignSelectionType As CareServices.XMLCampaignDataSelectionTypes
    Public EventSelectionType As CareServices.XMLEventDataSelectionTypes

    Public Enum SelectionItemTypes
      sitContactSelection
      sitCampaignSelection
      sitEventSelection
    End Enum

    Public Sub New(ByVal pRow As DataRow, ByVal pType As SelectionItemTypes)
      Description = pRow("Description").ToString
      Code = pRow("Code").ToString
      Select Case pType
        Case SelectionItemTypes.sitContactSelection
          ContactSelectionType = CType(pRow("Type"), CareServices.XMLContactDataSelectionTypes)
          ReadOnlyPage = pRow("ReadOnly").ToString = "Y"
        Case SelectionItemTypes.sitCampaignSelection
          CampaignSelectionType = CType(pRow("Type"), CareServices.XMLCampaignDataSelectionTypes)
          Dim vStartDate As String = pRow("StartDate").ToString
          Dim vEndDate As String = pRow("EndDate").ToString
          CampaignItem = New CampaignItem(Code, vStartDate, vEndDate)
          Sequence = IntegerValue(pRow("Sequence").ToString)
        Case SelectionItemTypes.sitEventSelection
          EventSelectionType = CType(pRow("Type"), CareServices.XMLEventDataSelectionTypes)
      End Select
    End Sub

    Public Sub New(ByVal pType As CareServices.XMLCampaignDataSelectionTypes, ByVal pCode As String)
      Select Case pType
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCampaign
          Description = "New Campaign"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppeal
          Description = "New Appeal"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegment
          Description = "New Segment"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollection
          Description = "New Collection"

        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionRegions
          Description = "Regions"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectors, CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectors
          Description = "Collectors"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPIS, CareServices.XMLCampaignDataSelectionTypes.xcadtH2HCollectionPIS
          Description = "Paying In Slips"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtMannedCollectionBoxes, CareServices.XMLCampaignDataSelectionTypes.xcadtUnMannedCollectionBoxes
          Description = "Boxes"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionResources
          Description = "Resources"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCollectionPayments
          Description = "Income"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtTickBoxes
          Description = "Tick Boxes"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtSegmentProducts
          Description = "Product Allocation"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtCostCentres
          Description = "Cost Centres"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealBudgets
          Description = "Budgets"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtSuppliers
          Description = "Suppliers"
        Case CareServices.XMLCampaignDataSelectionTypes.xcadtAppealResources
          Description = "Resources"
      End Select
      CampaignSelectionType = pType
      Code = pCode
      CampaignItem = New CampaignItem(pCode, "", "")
    End Sub
  End Class

  Public Sub SetCampaignRestrictions(ByVal pCampaignRestrictions As ParameterList)
    mvCampaignRestrictions = pCampaignRestrictions
  End Sub

  Public Function GetParentCampaignItem(ByVal pItemType As CampaignItem.CampaignItemTypes) As CampaignItem
    Dim vNode As TreeNode = tvw.SelectedNode
    While vNode.Parent IsNot Nothing
      Dim vItem As SelectionItem = DirectCast(vNode.Parent.Tag, SelectionItem)
      If vItem.CampaignItem.ItemType = pItemType Then
        Return vItem.CampaignItem
      End If
      vNode = vNode.Parent
    End While
    Return Nothing
  End Function
End Class
