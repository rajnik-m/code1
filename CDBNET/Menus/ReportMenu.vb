Public Class ReportMenu
  Inherits ContextMenuStrip

  'Private mvParent As MaintenanceParentForm
  Private mvReportNumber As Integer
  Private mvNodeInfo As TreeViewNodeInfo
  Public Event MenuActionCompleted(ByVal pItem As ReportMenuItems)

  Public Enum ReportMenuItems
    rmiDuplicateReport
    rmiRenumberParameter
    rmiRenumberSection
    rmiDuplicateSection
    rmiRenumberItem
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  '  Public Sub New(ByVal pParent As MaintenanceParentForm)
  Public Sub New()
    MyBase.New()
    'mvParent = pParent
    With mvMenuItems
      .Add(ReportMenuItems.rmiDuplicateReport.ToString, New MenuToolbarCommand("Duplicate Report", ControlText.MnuDuplicateReport, ReportMenuItems.rmiDuplicateReport))
      .Add(ReportMenuItems.rmiRenumberParameter.ToString, New MenuToolbarCommand("Renumber Paramters", ControlText.MnuRenumberParamaters, ReportMenuItems.rmiRenumberParameter))
      .Add(ReportMenuItems.rmiRenumberSection.ToString, New MenuToolbarCommand("Renumber Sections", ControlText.MnuRenumberSections, ReportMenuItems.rmiRenumberSection))
      .Add(ReportMenuItems.rmiDuplicateSection.ToString, New MenuToolbarCommand("Duplicate Section", ControlText.MnuDuplicateSection, ReportMenuItems.rmiDuplicateSection))
      .Add(ReportMenuItems.rmiRenumberItem.ToString, New MenuToolbarCommand("Renumber Items", ControlText.MnuRenumberItems, ReportMenuItems.rmiRenumberItem))
    End With
    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)
  End Sub
  Public ReadOnly Property ReportNumber() As Integer
    Get
      Return mvReportNumber
    End Get
  End Property

  Public Property ReportNodeInfo() As TreeViewNodeInfo
    Get
      Return mvNodeInfo
    End Get
    Set(ByVal Value As TreeViewNodeInfo)
      mvNodeInfo = Value
    End Set
  End Property

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ReportMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, ReportMenuItems)

      Select Case vMenuItem
        Case ReportMenuItems.rmiDuplicateReport
          Dim vForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptReportNumber, Nothing, Nothing)
          Dim vList As ParameterList = New ParameterList(True)
          Dim vReturnList As ParameterList = Nothing
          If vForm.ShowDialog() = DialogResult.OK Then
            vReturnList = vForm.ReturnList
            vList("ReportNumber") = mvNodeInfo.ReportNumber.ToString()
            vList("ReportCode") = mvNodeInfo.ReportCode.ToString()
            vList("NewReportNumber") = vReturnList("ReportNumber")
            DataHelper.DuplicateReport(vList)
            mvReportNumber = IntegerValue(vReturnList("ReportNumber"))
          End If
          

          
        Case ReportMenuItems.rmiRenumberParameter
          Dim vList As ParameterList = New ParameterList(True)
          vList("ReportNumber") = mvNodeInfo.ReportNumber.ToString()
          DataHelper.RenumberParameters(vList)
        Case ReportMenuItems.rmiRenumberSection
          Dim vList As ParameterList = New ParameterList(True)
          vList("ReportNumber") = mvNodeInfo.ReportNumber.ToString()
          DataHelper.RenumberSections(vList)
        Case ReportMenuItems.rmiDuplicateSection
          Dim vList As ParameterList = New ParameterList(True)
          vList("ReportNumber") = mvNodeInfo.ReportNumber.ToString()
          vList("SectionNumber") = mvNodeInfo.SectionNumber.ToString()
          DataHelper.DuplicateSection(vList)
        Case ReportMenuItems.rmiRenumberItem
          Dim vList As ParameterList = New ParameterList(True)
          vList("ReportNumber") = mvNodeInfo.ReportNumber.ToString()
          vList("SectionNumber") = mvNodeInfo.SectionNumber.ToString()
          DataHelper.RenumberItems(vList)
      End Select
      RaiseEvent MenuActionCompleted(vMenuItem)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub ReportMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    Dim vCursor As New BusyCursor
    Try
      'make DeletAllContacts visible Only if the User has Access Rights
      'Me.Items(ReportMenuItems.ssmiDeleteAllContacts).Visible = Not mvMenuItems.Item(ReportMenuItems.ssmiDeleteAllContacts.ToString).HideItem
      If mvNodeInfo.NodeType = TreeViewNodeType.SubSection Then
        Me.Items(ReportMenuItems.rmiDuplicateSection).Visible = True
        Me.Items(ReportMenuItems.rmiRenumberItem).Visible = True
      Else
        Me.Items(ReportMenuItems.rmiDuplicateSection).Visible = False
        Me.Items(ReportMenuItems.rmiRenumberItem).Visible = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

End Class
