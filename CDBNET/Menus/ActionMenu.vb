Public Class ActionMenu
  Inherits BaseActionMenu

  Public Shadows Event NewAction(ByVal sender As Object, ByVal e As System.EventArgs)
  Public Shadows Event EditAction(ByVal sender As Object, ByVal e As System.EventArgs)
  Public Shadows Event NewActionFromTemplate(ByVal sender As Object, ByVal e As System.EventArgs, ByVal pActionNumber As Integer)
  Public Shadows Event DeleteAction(ByVal sender As Object, ByVal e As System.EventArgs)
  Public Shadows Event RefreshCard(ByVal sender As Object)

  <Obsolete("For Form Designer use only", True)>
  Public Sub New()
    MyBase.New()
    If Not Me.DesignMode Then
      Throw New NotSupportedException("Default constructor is provided for Form Designer compatibiliy only and should not be used in code.")
    End If
  End Sub

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent)
  End Sub

  Protected Overrides Sub AddFollowupAction(ByVal pList As ParameterList)
    FormHelper.EditAction(ActionNumber, mvParent, pList, mvContactInfo)
  End Sub

  Protected Overrides Function AddNewActionFromTemplate(ByVal sender As Object, ByVal e As EventArgs) As Integer
    Dim vMasterAction As Integer = FormHelper.NewActionFromTemplate(mvParent, ContactNumber, mvList)
    Select Case ActionType
      Case ActionTypes.CampaignActions, ActionTypes.EventActions
        RaiseEvent NewActionFromTemplate(sender, e, vMasterAction)
    End Select
    Return vMasterAction
  End Function

  Protected Overrides Sub HandleDeleteAction(sender As Object, e As EventArgs)
    RaiseEvent DeleteAction(sender, e)
  End Sub

  Protected Overrides Sub DoNotify()
    If Not FormHelper.NotifyForm Is Nothing Then FormHelper.NotifyForm.DoRefresh()
  End Sub

  Protected Overrides Function DoRefresh() As Boolean
    If TypeOf mvParent Is frmEventSet OrElse TypeOf mvParent Is frmCampaignSet Then
      RaiseEvent RefreshCard(Me)
    End If
    Return False
  End Function

  Protected Overrides Sub HandleNewAction(ByVal sender As Object, ByVal e As System.EventArgs)
    ProcessNewEditAction(sender, e)
  End Sub

  Protected Overrides Sub HandleEditAction(ByVal sender As Object, ByVal e As System.EventArgs)
    ProcessNewEditAction(sender, e)
  End Sub

  Private Sub ProcessNewEditAction(ByVal sender As Object, ByVal e As EventArgs)
    Dim vMenuItem As ActionMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, ActionMenuItems)

    If TypeOf mvParent Is frmCardSet Or TypeOf mvParent Is frmEventSet Or TypeOf mvParent Is frmCampaignSet Or TypeOf mvParent Is frmExams Then
      If vMenuItem = ActionMenuItems.amiNew Then
        RaiseEvent NewAction(sender, e)
      Else
        RaiseEvent EditAction(sender, e)
      End If
    Else
      Dim vActionToEdit As Integer = 0
      If vMenuItem = ActionMenuItems.amiEdit Then vActionToEdit = ActionNumber
      FormHelper.EditAction(vActionToEdit, mvParent, mvContactInfo)
    End If
  End Sub
  Protected Overrides Sub HandleNewActionLink(sender As Object, e As EventArgs)
    Try
      Dim vList As New ParameterList(True)
      vList("PositionActionFinder") = "Y"
      Dim vActionNumber As Integer = FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftActions, vList, mvParent)
      If vActionNumber > 0 Then
        vList = New ParameterList(True, True)
        vList.IntegerValue("ActionNumber") = vActionNumber
        vList.IntegerValue("ContactPositionNumber") = ContactPositionNumber
        vList("ActionLinkType") = "R"   'These are always 'related-to' links
        DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
        If mvParent IsNot Nothing Then mvParent.RefreshData()
      End If

    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Protected Overrides Sub HandleDeleteActionLink(sender As Object, e As EventArgs)
    Try
      If ConfirmDelete() Then
        Dim vList As New ParameterList(True, True)
        vList.IntegerValue("ActionNumber") = ActionNumber
        vList("ActionLinkType") = "R"   'These are always 'related-to' links
        If ActionType = BaseActionMenu.ActionTypes.PositionActions Then
          vList.IntegerValue("ContactPositionNumber") = ContactPositionNumber
        Else
          'Not supported
          Throw New NotImplementedException("Delete Action Link menu only available for Position Links")
        End If
        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
        If mvParent IsNot Nothing Then mvParent.RefreshData()
      End If

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

End Class
