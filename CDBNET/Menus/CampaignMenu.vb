Public Class CampaignMenu
  Inherits CampaignBaseMenu

  Private mvParent As MaintenanceParentForm
  Private mvCriteriaSet As Integer
  Private mvMailingType As String
  Public Event MenuSelected(ByVal pItem As CampaignMenuItems)

  Public Enum CampaignMenuItems
    cmiNewAppeal
    cmiNewSegment
    cmiNewCollection
    cmiSegmentCriteria
    cmiSegmentSteps
    cmiSegmentPrint
    cmiSumAppeal
    cmiCountAppealOrSegment
    cmiMailAppeal
    cmiCopyData
    cmiCopyCriteria
    cmiReports
    cmiActions
    cmiCalculateIncome
    cmiPaste
    cmiAddCollectionBoxes
    cmiCountCollectors
    cmiCollectionFulfilment
    cmiMailAppealParams
    'cmiSelectionCriteria
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Property CriteriaSet() As Integer
    Get
      Return mvCriteriaSet
    End Get
    Set(ByVal value As Integer)
      mvCriteriaSet = value
    End Set
  End Property

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New()
    mvParent = pParent

    With mvMenuItems
      .Add(CampaignMenuItems.cmiNewAppeal.ToString, New MenuToolbarCommand("NewAppeal", ControlText.MnuCampaignNewAppeal, CampaignMenuItems.cmiNewAppeal, "SCCPNA"))
      .Add(CampaignMenuItems.cmiNewSegment.ToString, New MenuToolbarCommand("NewSegment", ControlText.MnuCampaignNewSegment, CampaignMenuItems.cmiNewSegment, "SCCPNS"))
      .Add(CampaignMenuItems.cmiNewCollection.ToString, New MenuToolbarCommand("NewCollection", ControlText.MnuCampaignNewCollection, CampaignMenuItems.cmiNewCollection, "SCCPNC"))
      .Add(CampaignMenuItems.cmiSegmentCriteria.ToString, New MenuToolbarCommand("SegmentCriteria", ControlText.MnuSelectionCriteria, CampaignMenuItems.cmiSegmentCriteria, "SCCPSC"))
      .Add(CampaignMenuItems.cmiSegmentSteps.ToString, New MenuToolbarCommand("SegmentSteps", ControlText.MnuCampaignSegmentSteps, CampaignMenuItems.cmiSegmentSteps, "SCCPSS"))
      .Add(CampaignMenuItems.cmiSegmentPrint.ToString, New MenuToolbarCommand("SegmentPrint", ControlText.MnuCampaignSegmentPrint, CampaignMenuItems.cmiSegmentPrint, "SCCPSP"))
      .Add(CampaignMenuItems.cmiSumAppeal.ToString, New MenuToolbarCommand("SumAppeal", ControlText.MnuCampaignSumAppeal, CampaignMenuItems.cmiSumAppeal, "SCCPSA"))
      .Add(CampaignMenuItems.cmiCountAppealOrSegment.ToString, New MenuToolbarCommand("CountAppeal", ControlText.MnuCampaignCountAppealOrSegment, CampaignMenuItems.cmiCountAppealOrSegment, "SCCPCA"))
      .Add(CampaignMenuItems.cmiMailAppeal.ToString, New MenuToolbarCommand("MailAppeal", ControlText.MnuCampaignMailAppeal, CampaignMenuItems.cmiMailAppeal, "SCCPMA"))
      .Add(CampaignMenuItems.cmiCopyData.ToString, New MenuToolbarCommand("CopyAppeal", ControlText.MnuCampaignCopyAppealOrSegment, CampaignMenuItems.cmiCopyData, "SCCPCP"))
      .Add(CampaignMenuItems.cmiCopyCriteria.ToString, New MenuToolbarCommand("CopyCriteria", ControlText.MnuCampaignCopyCriteria, CampaignMenuItems.cmiCopyCriteria, "SCCPCC"))
      .Add(CampaignMenuItems.cmiReports.ToString, New MenuToolbarCommand("Reports", ControlText.MnuCampaignReports, CampaignMenuItems.cmiReports, "SCCPRP"))
      .Add(CampaignMenuItems.cmiActions.ToString, New MenuToolbarCommand("CampaignActions", ControlText.MnuCampaignActions, CampaignMenuItems.cmiActions, "SCCPAC"))
      .Add(CampaignMenuItems.cmiCalculateIncome.ToString, New MenuToolbarCommand("CalculateIncome", ControlText.MnuCampaignCalculateIncome, CampaignMenuItems.cmiCalculateIncome, "SCCPCI"))
      .Add(CampaignMenuItems.cmiPaste.ToString, New MenuToolbarCommand("Paste", ControlText.MnuPaste, CampaignMenuItems.cmiPaste, "SCCPPA"))
      .Add(CampaignMenuItems.cmiAddCollectionBoxes.ToString, New MenuToolbarCommand("AddCollectionBoxes", ControlText.MnuAddCollectionBoxes, CampaignMenuItems.cmiAddCollectionBoxes, "SCCPAB"))
      .Add(CampaignMenuItems.cmiCountCollectors.ToString, New MenuToolbarCommand("CountCollectors", ControlText.MnuCampaignCountCollectors, CampaignMenuItems.cmiCountCollectors, "SCCPCO"))
      .Add(CampaignMenuItems.cmiCollectionFulfilment.ToString, New MenuToolbarCommand("CollectionFulfilment", ControlText.MnuCampaignCollectionsFulfilment, CampaignMenuItems.cmiCollectionFulfilment, "SCCPCF"))
      .Add(CampaignMenuItems.cmiMailAppealParams.ToString, New MenuToolbarCommand("MailAppealParams", ControlText.MnuCampaignMailAppealParams, CampaignMenuItems.cmiMailAppealParams, "SCCPMP"))
      '.Add(CampaignMenuItems.cmiSelectionCriteria.ToString, New MenuToolbarCommand("SelectionCriteria", ControlText.mnuSelectionCriteria, CampaignMenuItems.cmiSelectionCriteria, "SCCPSC"))
    End With

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    MenuToolbarCommand.SetAccessControl(mvMenuItems)

  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As CampaignMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, CampaignMenuItems)

      Select Case vMenuItem
        Case CampaignMenuItems.cmiNewAppeal, CampaignMenuItems.cmiNewSegment, CampaignMenuItems.cmiNewCollection, CampaignMenuItems.cmiSumAppeal, _
        CampaignMenuItems.cmiCalculateIncome, CampaignMenuItems.cmiAddCollectionBoxes, CampaignMenuItems.cmiCountCollectors, _
        CampaignMenuItems.cmiCopyData, CampaignMenuItems.cmiPaste, CampaignMenuItems.cmiSegmentSteps, CampaignMenuItems.cmiCountAppealOrSegment, CampaignMenuItems.cmiMailAppeal, _
        CampaignMenuItems.cmiSegmentCriteria, CampaignMenuItems.cmiMailAppealParams
          RaiseEvent MenuSelected(vMenuItem)
          'Case CampaignMenuItems.cmiSegmentCriteria
          '
        Case CampaignMenuItems.cmiSegmentPrint
          Dim vList As New ParameterList(True)
          vList("ReportCode") = "MACP"
          vList("RP1") = CampaignItem.Campaign
          vList("RP2") = CampaignItem.Appeal
          vList("RP3") = CampaignItem.Segment
          vList("RP4") = mvCriteriaSet.ToString
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case CampaignMenuItems.cmiCopyCriteria
          RaiseEvent MenuSelected(vMenuItem)
        Case CampaignMenuItems.cmiActions
          RaiseEvent MenuSelected(vMenuItem)

      End Select
    Catch vException As CareException
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub



  Private Sub CampaignMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    Dim vCursor As New BusyCursor
    For Each vItem As ToolStripItem In Me.Items
      vItem.Visible = False
    Next
    DirectCast(Me.Items(CampaignMenuItems.cmiReports), ToolStripMenuItem).DropDownItems.Clear()
    DirectCast(Me.Items(CampaignMenuItems.cmiCollectionFulfilment), ToolStripMenuItem).DropDownItems.Clear()

    Try
      e.Cancel = False
      If CampaignItem IsNot Nothing AndAlso CampaignItem.Existing Then
        Dim vAppealLocked As Boolean = CampaignItem.AppealLocked
        Select Case CampaignItem.ItemType
          Case CampaignItem.CampaignItemTypes.citCampaign
            mvMenuItems(CampaignMenuItems.cmiNewAppeal).SetContextItemVisible(Me, Not CampaignItem.MarkHistorical)
            mvMenuItems(CampaignMenuItems.cmiCalculateIncome).SetContextItemVisible(Me, True)
            mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemVisible(Me, True)
            mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemEnabled(Me, EnablePaste)
            If GetReports("MKTCAM") > 0 Then mvMenuItems(CampaignMenuItems.cmiReports).SetContextItemVisible(Me, True)
            If GetFulfilmentsMenu() > 0 Then mvMenuItems(CampaignMenuItems.cmiCollectionFulfilment).SetContextItemVisible(Me, True)

          Case CampaignItem.CampaignItemTypes.citAppeal
            If vAppealLocked Then                   'BR11765: Dont display menu if the appeal is locked.
              'ShowInformationMessage(InformationMessages.ImAppealLocked)
            Else
              If CampaignItem.Appeal.Length > 0 Then
                Dim vNewSegment As Boolean = True
                Dim vSum As Boolean = True
                Dim vCount As Boolean = True
                Dim vMail As Boolean = True
                Dim vCopy As Boolean = True
                Dim vIncome As Boolean = True
                If CampaignItem.SegmentCount = 0 Then
                  'vSum = False
                  vCount = False
                  vMail = False
                Else
                  vMail = Not CampaignItem.MarkHistorical
                End If
                If CampaignItem.AppealType = CampaignItem.AppealTypes.atSegment Then
                  vNewSegment = Not CampaignItem.MarkHistorical
                  vCopy = Not CampaignItem.MarkHistorical
                Else
                  vNewSegment = False
                  ' BR 12517 - Previous BR change 11765 incorrectly made most items invisible. 
                  ' By commenting out the following code we return them to pre 6.2 state
                  'vSum = False
                  'vCount = False
                  'vMail = False
                  'vCopy = False
                  'vIncome = False
                End If
                mvMenuItems(CampaignMenuItems.cmiNewCollection).SetContextItemVisible(Me, (CampaignItem.AppealType <> CampaignItem.AppealTypes.atSegment))
                mvMenuItems(CampaignMenuItems.cmiNewSegment).SetContextItemVisible(Me, vNewSegment)
                mvMenuItems(CampaignMenuItems.cmiSumAppeal).SetContextItemVisible(Me, vSum)
                mvMenuItems(CampaignMenuItems.cmiCountAppealOrSegment).SetContextItemVisible(Me, vCount)
                mvMenuItems(CampaignMenuItems.cmiMailAppeal).SetContextItemVisible(Me, vMail)
                mvMenuItems(CampaignMenuItems.cmiCopyData).SetContextItemVisible(Me, vCopy)
                mvMenuItems(CampaignMenuItems.cmiCopyData).SetContextItemEnabled(Me, vCopy)
                mvMenuItems(CampaignMenuItems.cmiActions).SetContextItemVisible(Me, True)                'Menu not available yet
                mvMenuItems(CampaignMenuItems.cmiCalculateIncome).SetContextItemVisible(Me, vIncome)
                mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemEnabled(Me, EnablePaste())

                'BR 17167 only show this menu if there are parameters
                Dim vList As ParameterList = Nothing
                Dim pType As CareServices.TaskJobTypes = CareNetServices.TaskJobTypes.tjtMailingRun
                vList = New ParameterList(True, True)
                vList.Add("Campaign", CampaignItem.Campaign)
                vList.Add("Appeal", CampaignItem.Appeal)
                vList.Add("Segment", CampaignItem.Segment)

                vList.Add("Mail", "Y")
                Dim vDataSet As DataSet = DataHelper.GetCampaignCriteriaVariableControls(vList, pType)

                If vDataSet IsNot Nothing Then
                  Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
                  If vTable IsNot Nothing Then
                    If vDataSet.Tables.Contains("Parameters") Then
                      mvMenuItems(CampaignMenuItems.cmiMailAppealParams).SetContextItemVisible(Me, vMail)
                    End If
                  End If
                End If
                vDataSet.Dispose()

                If GetReports("MKTAPP") > 0 Then mvMenuItems(CampaignMenuItems.cmiReports).SetContextItemVisible(Me, True)
                If GetFulfilmentsMenu() > 0 Then mvMenuItems(CampaignMenuItems.cmiCollectionFulfilment).SetContextItemVisible(Me, True)
              Else
                e.Cancel = True
              End If
            End If
          Case CampaignItem.CampaignItemTypes.citSegment
            If vAppealLocked Then                     'BR11765: Dont display menu if the appeal is locked.
              'ShowInformationMessage(InformationMessages.ImAppealLocked)
            Else
              If CampaignItem.Segment.Length > 0 Then
                'mvMenuItems(CampaignMenuItems.cmiSegmentCriteria).SetContextItemVisible(Me, True)        'Menu not available yet
                mvMenuItems(CampaignMenuItems.cmiSegmentSteps).SetContextItemVisible(Me, CampaignItem.AppealMailingTypeCode = "MKTG")
                mvMenuItems(CampaignMenuItems.cmiSegmentPrint).SetContextItemVisible(Me, (mvCriteriaSet > 0))
                mvMenuItems(CampaignMenuItems.cmiCountAppealOrSegment).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiCopyData).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiCopyCriteria).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiCopyCriteria).SetContextItemEnabled(Me, False)
                mvMenuItems(CampaignMenuItems.cmiCalculateIncome).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemVisible(Me, True)
                mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemEnabled(Me, EnablePaste())
                mvMenuItems(CampaignMenuItems.cmiSegmentCriteria).SetContextItemVisible(Me, True)

                ' Display Segment Criteria menu only if count of Selection_Steps for the criteria is Zero
                If mvCriteriaSet > 0 Then
                  Dim vList As New ParameterList(True)
                  vList.IntegerValue("CriteriaSetNumber") = mvCriteriaSet
                  Dim vHasSelectionSteps As Boolean = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctSelectionSteps, vList) > 0
                  mvMenuItems(CampaignMenuItems.cmiSegmentCriteria).SetContextItemEnabled(Me, vHasSelectionSteps = False)
                  vList = New ParameterList(True)
                  vList.IntegerValue("CriteriaSet") = mvCriteriaSet
                  Dim vHasCriteriaSetDetails As Boolean = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCriteriaSetDetails, vList) > 0
                  mvMenuItems(CampaignMenuItems.cmiSegmentSteps).SetContextItemEnabled(Me, vHasCriteriaSetDetails = False)
                  'Prevent copying criteria and segment itself where we have both criteria set details and selection steps as this is invalid
                  mvMenuItems(CampaignMenuItems.cmiCopyCriteria).SetContextItemEnabled(Me, Not (vHasCriteriaSetDetails AndAlso vHasSelectionSteps))
                  mvMenuItems(CampaignMenuItems.cmiCopyData).SetContextItemEnabled(Me, Not (vHasCriteriaSetDetails AndAlso vHasSelectionSteps))
                Else
                  mvMenuItems(CampaignMenuItems.cmiSegmentCriteria).SetContextItemEnabled(Me, True)
                  mvMenuItems(CampaignMenuItems.cmiSegmentSteps).SetContextItemEnabled(Me, True)
                End If
                If GetReports("MKTSEG") > 0 Then mvMenuItems(CampaignMenuItems.cmiReports).SetContextItemVisible(Me, True)
              Else
                e.Cancel = True
              End If
            End If
          Case CampaignItem.CampaignItemTypes.citCollection
            mvMenuItems(CampaignMenuItems.cmiAddCollectionBoxes).SetContextItemVisible(Me, (CampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection Or CampaignItem.AppealType = CampaignItem.AppealTypes.atUnMannedCollection))
            mvMenuItems(CampaignMenuItems.cmiCountCollectors).SetContextItemVisible(Me, (CampaignItem.AppealType = CampaignItem.AppealTypes.atMannedCollection Or CampaignItem.AppealType = CampaignItem.AppealTypes.atH2HCollection))
            mvMenuItems(CampaignMenuItems.cmiCalculateIncome).SetContextItemVisible(Me, True)
            mvMenuItems(CampaignMenuItems.cmiCopyData).SetContextItemVisible(Me, True)
            mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemVisible(Me, True)
            mvMenuItems(CampaignMenuItems.cmiPaste).SetContextItemEnabled(Me, EnablePaste())
            If GetFulfilmentsMenu() > 0 Then mvMenuItems(CampaignMenuItems.cmiCollectionFulfilment).SetContextItemVisible(Me, True)
          Case Else
            e.Cancel = True
        End Select
      Else
        e.Cancel = True
      End If
    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enVarInMultipleAreas
          ShowInformationMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enVariableNameContainsInvalidCharacters
          ShowErrorMessage(vCareException.Message)
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Function GetReports(ByVal pReportCode As String) As Integer
    Dim vList As New ParameterList(True)
    vList("ReportCode") = pReportCode
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtReports, vList)
    Dim vReports As ToolStripMenuItem = mvMenuItems(CampaignMenuItems.cmiReports).FindMenuStripItem(Me)

    If vTable IsNot Nothing And vReports IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        vReports.DropDownItems.Add(vRow.Item("ReportName").ToString, Nothing, AddressOf ReportMenuHandler).Tag = vRow.Item("ReportNumber")
      Next
      If vReports.DropDownItems.Count = 0 Then vReports.Visible = False
      Return vReports.DropDownItems.Count
    End If

  End Function

  Private Sub ReportMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vReportNumber As Integer = CInt(vMenuItem.Tag)
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ReportNumber") = vReportNumber
      vList("RP1") = CampaignItem.Campaign
      If CampaignItem.Appeal.Length > 0 Then vList("RP2") = CampaignItem.Appeal
      If CampaignItem.Segment.Length > 0 Then vList("RP3") = CampaignItem.Segment
      Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Function EnablePaste() As Boolean
    With CampaignItem
      If (.ItemType = CampaignItem.CampaignItemTypes.citCampaign _
         Or .ItemType = CampaignItem.CampaignItemTypes.citAppeal) _
         AndAlso FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctAppeal) Then
        Return True
      ElseIf .AppealType = CampaignItem.AppealTypes.atH2HCollection _
             AndAlso FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctH2HCollection) Then
        'copy h2h collection
        Return True
      ElseIf .AppealType = CampaignItem.AppealTypes.atMannedCollection _
             AndAlso FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctMannedCollection) Then
        'copy Manned collection
        Return True
      ElseIf .AppealType = CampaignItem.AppealTypes.atUnMannedCollection _
             AndAlso FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctUnMannedCollection) Then
        'copy Unmanned collection
        Return True
      ElseIf .AppealType = CampaignItem.AppealTypes.atSegment _
             AndAlso (FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctSegment) OrElse FormHelper.ClipboardContainsCampaignData(CampaignCopyInfo.CampaignCopyTypes.cctSegmentCriteria)) Then
        'copy segment or segment criteria set
        Return True
      Else
        Return False
      End If
    End With
  End Function

  Private Function GetFulfilmentsMenu() As Integer
    'Menu item tag is the Fulfilment type
    Dim vFulfilmentMenu As ToolStripMenuItem = mvMenuItems(CampaignMenuItems.cmiCollectionFulfilment).FindMenuStripItem(Me)

    Select Case CampaignItem.ItemType
      Case CampaignItem.CampaignItemTypes.citCampaign
        With vFulfilmentMenu
          .DropDownItems.Add(ControlText.MnuCampaignAcknowledgementFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "A"
          .DropDownItems.Add(ControlText.MnuCampaignConfirmationFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "C"
          .DropDownItems.Add(ControlText.MnuCampaignEndOfCollectionFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "E"
          .DropDownItems.Add(ControlText.MnuCampaignLabelsFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "L"
          .DropDownItems.Add(ControlText.MnuCampaignReminderFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "R"
          .DropDownItems.Add(ControlText.MnuCampaignResourcesFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "O"
        End With

      Case CampaignItem.CampaignItemTypes.citAppeal, CampaignItem.CampaignItemTypes.citCollection
        Select Case CampaignItem.AppealType
          Case CampaignItem.AppealTypes.atMannedCollection
            With vFulfilmentMenu
              .DropDownItems.Add(ControlText.MnuCampaignAcknowledgementFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "A"
              .DropDownItems.Add(ControlText.MnuCampaignConfirmationFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "C"
              .DropDownItems.Add(ControlText.MnuCampaignLabelsFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "L"
            End With
          Case CampaignItem.AppealTypes.atUnMannedCollection
            With vFulfilmentMenu
              .DropDownItems.Add(ControlText.MnuCampaignAcknowledgementFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "A"
              .DropDownItems.Add(ControlText.MnuCampaignConfirmationFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "C"
              .DropDownItems.Add(ControlText.MnuCampaignEndOfCollectionFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "E"
              .DropDownItems.Add(ControlText.MnuCampaignLabelsFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "L"
              .DropDownItems.Add(ControlText.MnuCampaignReminderFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "R"
              .DropDownItems.Add(ControlText.MnuCampaignResourcesFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "O"
            End With
          Case CampaignItem.AppealTypes.atH2HCollection
            With vFulfilmentMenu
              .DropDownItems.Add(ControlText.MnuCampaignConfirmationFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "C"
              .DropDownItems.Add(ControlText.MnuCampaignReminderFulfilment, Nothing, AddressOf FulfilmentMenuHandler).Tag = "R"
            End With
        End Select

    End Select
    Return vFulfilmentMenu.DropDownItems.Count

  End Function

  Private Sub FulfilmentMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor()

    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Dim vFulfilmentType As String = DirectCast(vMenuItem.Tag, String)
      Dim vDefaults As New ParameterList    'ApplicationParameters default values
      Dim vList As New ParameterList

      vDefaults("FulfilmentType") = vFulfilmentType
      With CampaignItem
        vDefaults("Campaign") = .Campaign
        If .Appeal.Length > 0 Then
          vDefaults("Appeal") = .Appeal
          vDefaults("CollectionType") = .AppealTypeCode
          If .CollectionNumber > 0 Then
            vDefaults("CollectionNumber") = .CollectionNumber.ToString
          End If
          If .AppealType = CampaignItem.AppealTypes.atH2HCollection AndAlso vFulfilmentType = "C" Then
            'House 2 House Confirmation Fulfilment
            vDefaults("CollectorStatus") = AppValues.ControlValue(AppValues.ControlValues.default_collector_status)
          End If
        Else
          Select Case vFulfilmentType
            Case "A", "C", "L"
              vDefaults("CollectionType") = "M"
            Case "E", "O"
              vDefaults("CollectionType") = "U"
            Case "R"
              vDefaults("CollectionType") = "U"
          End Select
        End If
        If vFulfilmentType = "L" Then vDefaults("NoMailingHistory") = "Y"
      End With
      vDefaults("FulfilmentType") = vFulfilmentType
      FormHelper.ProcessTask(CareServices.TaskJobTypes.tjtPublicCollectionsFulfilment, vDefaults)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub



End Class

