Imports System.IO
Imports System.Drawing.Printing
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports CDBNETBiz

Public Class FormHelper

  Private Shared mvNotifyForm As frmNotify
  Private Shared mvTaskStatusForm As frmTaskStatus

  'Public Enum RunMailingResult
  '  NoMailingRun
  '  MailingRunSynchSuccess
  '  MailingRunSynchFail
  '  MailingRunAsych
  'End Enum

  Public Enum ProcessTaskScheduleType
    ptsNone
    ptsAskToSchedule
    ptsAlwaysSchedule
    ptsAlwaysRun
  End Enum
  'Adjust MarginBounds rectangle when printing based on the physical characteristics of the printer
  Public Shared Function GetRealMarginBounds(ByVal e As PrintPageEventArgs, ByVal pPreview As Boolean) As Rectangle
    If pPreview Then Return e.MarginBounds

    'Get printer’s offsets
    Dim vCX As Single = e.PageSettings.HardMarginX
    Dim vCY As Single = e.PageSettings.HardMarginY
    'Create the real margin bounds by scaling the offset by the printer resolution and then rescaling it back to 1/100th of an inch
    Dim vMarginBounds As Rectangle = e.MarginBounds
    Dim vDPIX As Single = e.Graphics.DpiX
    Dim vDPIY As Single = e.Graphics.DpiY
    vMarginBounds.Offset(CInt(-vCX * 100 / vDPIX), CInt(-vCY * 100 / vDPIY))
    Return vMarginBounds
  End Function

  Public Shared Sub ShowSelectionSet(ByVal pNumber As Integer, ByVal pDesc As String)
    Dim vForm As New frmSelectionSet(pNumber, pDesc)
    vForm.Show()
  End Sub

  Public Shared Function AddSelectionSet(ByVal pOwner As Form) As Integer
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctSelectionSet)
    If vForm.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
      Dim vSSNo As Integer = vForm.ReturnList.IntegerValue("SelectionSetNumber")
      Dim vSSDesc As String = vForm.ReturnList("SelectionSetDesc")
      If vSSNo > 0 Then ShowSelectionSet(vSSNo, vSSDesc)
      Return vSSNo
    End If
  End Function
  Public Shared Function MergeSelectionSet(ByVal pOwner As Form, ByVal pSelectionSetNumber As Integer) As Integer
    Dim vList As New ParameterList(True)
    vList("SelectionSetNumber") = pSelectionSetNumber.ToString
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctMergeSelectionSet, vList)
    If vForm.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
      Dim vSSNo As Integer = vForm.ReturnList.IntegerValue("SelectionSetNumber")
      If vSSNo > 0 Then
        Dim vDataSet As DataSet = DataHelper.GetSelectionSetData(vSSNo)
        Dim vContactCount As Integer = vDataSet.Tables("DataRow").Rows.Count
        Dim vSSDesc As String = vForm.ReturnList("SelectionSetDesc")
        ShowInformationMessage(InformationMessages.ImSelectionSetMergeDetails, vSSDesc, vContactCount.ToString)
        Return vSSNo
      End If
    End If
  End Function

  Public Shared Function NewActionFromTemplate(ByVal pParentForm As MaintenanceParentForm, ByVal pContactNumber As Integer) As Integer
    NewActionFromTemplate(pParentForm, pContactNumber, Nothing)
  End Function
  Public Shared Function NewActionFromTemplate(ByVal pParentForm As MaintenanceParentForm, ByVal pContactNumber As Integer, ByVal pList As ParameterList) As Integer
    Return ActionsHelper.NewActionFromTemplate(pParentForm, pContactNumber, pList)
  End Function


  Public Shared Sub EditActionTemplate(ByVal pNumber As Integer)
    Dim vForm As frmActionSet = Nothing
    Try
      vForm = New frmActionSet(pNumber, Nothing)
      vForm.Show()
      If pNumber = 0 Then 'BR17953 Clear Action Template cache
        DataHelper.ClearCachedTable("ActionTemplates", CareNetServices.XMLLookupDataTypes.xldtActionTemplates)
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
        vForm.Close()
        ShowInformationMessage(InformationMessages.ImCannotFindAction)
      Else
        Throw vException
      End If
    End Try
  End Sub

  Public Shared Sub EditMeeting(pNumber As Integer)
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetings, CareServices.XMLMaintenanceControlTypes), pNumber)
    vForm.Show()
  End Sub

  Public Shared Sub EditAction(ByVal pNumber As Integer)
    EditAction(pNumber, Nothing, Nothing, Nothing)
  End Sub
  Public Shared Sub EditAction(ByVal pNumber As Integer, ByVal pParentForm As MaintenanceParentForm)
    EditAction(pNumber, pParentForm, Nothing, Nothing)
  End Sub
  Public Shared Sub EditAction(ByVal pNumber As Integer, ByVal pParentForm As MaintenanceParentForm, ByVal pContactInfo As ContactInfo)
    EditAction(pNumber, pParentForm, Nothing, pContactInfo)
  End Sub
  Public Shared Sub EditAction(ByVal pNumber As Integer, ByVal pParentForm As MaintenanceParentForm, ByVal pList As ParameterList, ByVal pContactInfo As ContactInfo)
    ActionsHelper.EditAction(pNumber, pParentForm, pList, pContactInfo)
  End Sub
  Public Shared Sub EditMeeting(ByVal pMeetingNumber As Integer, ByVal pParentForm As MaintenanceParentForm)    'Edit the specified meeting
    EditMeeting(pMeetingNumber, pParentForm, Nothing)
  End Sub

  Private Shared Function CheckAddresseeStatusAndSuppression(ByVal pCurrentContact As ContactInfo) As Boolean
    Dim vIndex As Integer
    Dim vContinue As Boolean
    Dim vStatus As String

    vContinue = True
    If pCurrentContact IsNot Nothing Then
      vStatus = pCurrentContact.Status
      If vStatus.Length > 0 Then
        Dim vWarningStatuses() As String
        vWarningStatuses = Split(AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_document_warning_statuses), "|")
        For vIndex = 0 To UBound(vWarningStatuses)
          If vStatus = vWarningStatuses(vIndex) Then
            vContinue = ShowQuestion(QuestionMessages.QmDocumentAddresseeStatus, MessageBoxButtons.YesNo, pCurrentContact.StatusDesc) = System.Windows.Forms.DialogResult.Yes
          End If
        Next
      End If
      If vContinue AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_document_warn_suppressions).Length > 0 Then
        Dim vSuppressionCodes() As String = Split(AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_document_warn_suppressions), "|")
        Dim vSuppressionDescriptions As String = ""
        Dim vList As New ParameterList(True)
        vList("Current") = "Y"
        Dim vDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactSuppressions, pCurrentContact.ContactNumber, vList))
        If vDT IsNot Nothing Then
          For Each vDR As DataRow In vDT.Rows
            For Each vSuppression As String In vSuppressionCodes
              If vSuppression.Length > 0 AndAlso vDR.Item("SuppressionCode").ToString = vSuppression Then
                vSuppressionDescriptions = CStr(IIf(Len(vSuppressionDescriptions) > 0, vSuppressionDescriptions & ", " & vDR.Item("SuppressionDesc").ToString, vDR.Item("SuppressionDesc").ToString))
              End If
            Next
          Next
          If vSuppressionDescriptions.Length > 0 Then
            vContinue = ShowQuestion(QuestionMessages.QmDocumentAddresseeSuppression, MessageBoxButtons.YesNo, vSuppressionDescriptions) = System.Windows.Forms.DialogResult.Yes
          End If
        End If
      End If
    End If
    Return vContinue
  End Function

  Public Shared Sub NewDocument()     'Create a new document
    EditDocument(0, Nothing, Nothing)
  End Sub

  Public Shared Sub NewDocument(ByVal pParentForm As MaintenanceParentForm)     'Create a new document
    EditDocument(0, pParentForm, Nothing)
  End Sub
  Public Shared Sub NewDocument(ByVal pParentForm As MaintenanceParentForm, ByVal pDestinationContactInfo As ContactInfo, ByVal pParams As ParameterList)     'Create a new document
    EditDocument(0, pParentForm, pDestinationContactInfo, pParams)
  End Sub

  Public Shared Sub NewDocument(ByVal pParentForm As MaintenanceParentForm, ByVal pDestinationContactInfo As ContactInfo)     'Create a new document
    EditDocument(0, pParentForm, pDestinationContactInfo)
  End Sub

  Public Shared Sub EditDocument(ByVal pDocumentNumber As Integer)    'Edit the specified document
    EditDocument(pDocumentNumber, Nothing, Nothing)
  End Sub

  Public Shared Sub EditDocument(ByVal pDocumentNumber As Integer, ByVal pParentForm As MaintenanceParentForm)    'Edit the specified document
    EditDocument(pDocumentNumber, pParentForm, Nothing)
  End Sub

  Public Shared Sub EditDocument(ByVal pDocumentNumber As Integer, ByVal pParentForm As MaintenanceParentForm, ByVal pDestinationContactInfo As ContactInfo)
    EditDocument(pDocumentNumber, pParentForm, pDestinationContactInfo, Nothing)
  End Sub
  Public Shared Sub EditDocument(ByVal pDocumentNumber As Integer, ByVal pParentForm As MaintenanceParentForm, ByVal pDestinationContactInfo As ContactInfo, ByVal pParams As ParameterList)
    Dim vForm As frmCardMaintenance = Nothing
    Try
      Dim vContactInfo As ContactInfo = pDestinationContactInfo
      Dim vValid As Boolean = True
      Dim vNewDocument As Boolean
      If pDocumentNumber = 0 Then       'We are creating a new document
        vNewDocument = True
        If vContactInfo Is Nothing Then vContactInfo = MainHelper.CurrentContact
      End If
      'See if we have the correct access rights to communicate with this Contact
      If vContactInfo IsNot Nothing Then
        vValid = AppValues.ValidCommunicationsAccessLevel(vContactInfo, Nothing, "")
        If vValid Then vValid = CheckAddresseeStatusAndSuppression(vContactInfo)
      End If
      If vValid = True AndAlso pDocumentNumber > 0 Then
        'Check user access rights
        Dim vDataSet As DataSet = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentInformation, pDocumentNumber)
        If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Count > 0 AndAlso vDataSet.Tables(0).Rows.Count > 0 Then
          Dim vRow As DataRow = vDataSet.Tables(0).Rows(0)
          Dim vRights As DataHelper.DocumentAccessRights = CType(vRow.Item("AccessRights"), DataHelper.DocumentAccessRights)
          If vRights.HasFlag(DataHelper.DocumentAccessRights.darEditHeader) = False Then vValid = False 'User does not have permission
        End If
      End If
      If vValid Then
        Dim vRow As DataRow = Nothing
        If Not vNewDocument Then
          Dim vList As New ParameterList(True)
          vList.Add("DocumentNumber", pDocumentNumber)
          vList.Add("IncludeEmailDocSource", "Y")
          vRow = DataHelper.GetRowFromDataSet(DataHelper.FindData(CareServices.XMLDataFinderTypes.xdftDocuments, vList))
        End If
        If vRow IsNot Nothing Or vNewDocument Then
          Dim vCustomiseMenu As New CustomiseMenu
          vCustomiseMenu.SetContext(vForm, CareServices.XMLMaintenanceControlTypes.xmctDocument, "")
          If pParams IsNot Nothing AndAlso (pParams.ContainsKey("ExamUnitLinkId") OrElse pParams.ContainsKey("ExamCentreId") OrElse pParams.ContainsKey("ExamCentreUnitId") _
          OrElse pParams.ContainsKey("ContactCpdPeriodNumber") OrElse pParams.ContainsKey("ContactCpdPointNumber") OrElse pParams.ContainsKey("ContactPositionNumber")) Then
            vForm = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctDocument, pDocumentNumber, pParentForm, pParams, vContactInfo)
            vForm.SetCustomiseMenu(vCustomiseMenu)
            vForm.Show()
          Else
            vForm = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctDocument, pDocumentNumber, pParentForm, Nothing, vContactInfo)
            vForm.SetCustomiseMenu(vCustomiseMenu)
            vForm.Show()
          End If
        Else
          ShowInformationMessage(InformationMessages.ImCannotFindDocument)
        End If
      End If
    Catch vCareEx As CareException
      If vCareEx.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
        ShowInformationMessage(InformationMessages.ImCannotFindDocument)
        If vForm IsNot Nothing Then vForm.Close()
      Else
        Throw vCareEx
      End If
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub

  Public Shared Sub EditMeeting(ByVal pMeetingNumber As Integer, ByVal pParentForm As MaintenanceParentForm, ByVal pDestinationContactInfo As ContactInfo)
    Dim vForm As frmCardMaintenance = Nothing
    Try
      Dim vContactInfo As ContactInfo = pDestinationContactInfo
      Dim vValid As Boolean = True
      Dim vNewMeeting As Boolean
      If pMeetingNumber = 0 Then       'We are creating a new document
        vNewMeeting = True
        If vContactInfo Is Nothing Then vContactInfo = MainHelper.CurrentContact
      End If
      'See if we have the correct access rights to comminucate with this Contact
      If vContactInfo IsNot Nothing Then
        vValid = AppValues.ValidCommunicationsAccessLevel(vContactInfo, Nothing, "")
        If vValid Then vValid = CheckAddresseeStatusAndSuppression(vContactInfo)
      End If
      If vValid Then
        Dim vRow As DataRow = Nothing
        If Not vNewMeeting Then
          Dim vList As New ParameterList(True)
          vList.Add("MeetingNumber", pMeetingNumber)
          vRow = DataHelper.GetRowFromDataSet(DataHelper.FindData(CType(CareNetServices.XMLDataFinderTypes.xdftMeetings, CareServices.XMLDataFinderTypes), vList))
        End If
        If vRow IsNot Nothing Or vNewMeeting Then
          Dim vCustomiseMenu As New CustomiseMenu
          vCustomiseMenu.SetContext(vForm, CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetings, CareServices.XMLMaintenanceControlTypes), "")
          vForm = New frmCardMaintenance(CType(CareNetServices.XMLMaintenanceControlTypes.xmctMeetings, CareServices.XMLMaintenanceControlTypes), pMeetingNumber, pParentForm)
          vForm.SetCustomiseMenu(vCustomiseMenu)
          vForm.Show()
        Else
          '  ShowInformationMessage(InformationMessages.ImCannotFindDocument)
        End If
      End If
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub

  Public Shared Sub MaintainContactData(ByVal pContactNumber As Integer, ByVal pType As CareServices.XMLContactDataSelectionTypes, Optional ByVal pParentForm As MaintenanceParentForm = Nothing, Optional ByVal pList As ParameterList = Nothing)
    Dim vContactInfo As New ContactInfo(pContactNumber)
    Dim vDataSet As DataSet = DataHelper.GetContactData(pType, pContactNumber, pList)
    Dim vForm As frmCardMaintenance = New frmCardMaintenance(pParentForm, vContactInfo, pType, vDataSet, True, 0, , pList)
    vForm.Show()
  End Sub

  Public Shared Function MDIPointToScreen(ByVal pPoint As Point) As Point
    For Each vControl As Control In MDIForm.Controls
      If TypeOf vControl Is MdiClient Then
        Return vControl.PointToScreen(pPoint)
      End If
    Next
    Return New Point(0, 0)
  End Function

  Public Shared Property NotifyForm() As frmNotify
    Get
      Return mvNotifyForm
    End Get
    Set(ByVal pValue As frmNotify)
      mvNotifyForm = pValue
    End Set
  End Property

  Public Shared Property TaskStatusForm() As frmTaskStatus
    Get
      Return mvTaskStatusForm
    End Get
    Set(ByVal pValue As frmTaskStatus)
      mvTaskStatusForm = pValue
    End Set
  End Property

  Public Shared Function ShowNewContactOrDedup(ByVal pContactType As ContactInfo.ContactTypes, ByVal pList As ParameterList, Optional ByVal pOwner As Form = Nothing, Optional ByVal pAlwaysNew As Boolean = False) As Integer
    Dim vGroupParameterName As String = ""
    Dim vDataFinderType As CareServices.XMLDataFinderTypes

    If AppValues.ConfigurationValue(AppValues.ConfigurationValues.cd_contact_deduplication) = "I" Then pAlwaysNew = True

    If pContactType = ContactInfo.ContactTypes.ctContact Then
      vGroupParameterName = "ContactGroup"
      vDataFinderType = CareServices.XMLDataFinderTypes.xdftDuplicateContacts
    Else
      vGroupParameterName = "OrganisationGroup"
      vDataFinderType = CareServices.XMLDataFinderTypes.xdftDuplicateOrganisations
    End If
    If Not pList.ContainsKey(vGroupParameterName) Then pList(vGroupParameterName) = IIf(pContactType = ContactInfo.ContactTypes.ctContact, EntityGroup.DefaultContactGroupCode, EntityGroup.DefaultOrganisationGroupCode).ToString

    Dim vContactInfo As ContactInfo = New ContactInfo(pContactType, pList(vGroupParameterName))
    If pContactType = ContactInfo.ContactTypes.ctContact AndAlso pList.ContainsKey("CreateAtAddressNumber") Then
      vContactInfo.CreateAtAddressNumber = pList.IntegerValue("CreateAtAddressNumber")
    ElseIf DataHelper.ContactAndOrganisationGroups.ContainsKey(pList(vGroupParameterName)) Then
      Dim vEntityGroup As EntityGroup = DataHelper.ContactAndOrganisationGroups.Item(pList(vGroupParameterName))
      If vEntityGroup.AllAddressesUnknown Then
        pList("Town") = vEntityGroup.UnknownTown
        pList("Address") = vEntityGroup.UnknownAddress
        pList("Country") = AppValues.DefaultCountryCode
        If AppValues.IsBuildingNumberCountry(pList("Country")) Then pList("BuildingNumber") = "0"
        Return FormHelper.ShowNewContact(vContactInfo, pList, pOwner)
      End If
    End If
    If pAlwaysNew Then
      Return FormHelper.ShowNewContact(vContactInfo, pList, pOwner)
    Else
      Return FormHelper.ShowFinder(vDataFinderType, pList, pOwner)
    End If

  End Function

  Public Shared Function ShowNewContact(ByVal pContactInfo As ContactInfo, ByVal pList As ParameterList, Optional ByVal pOwner As Form = Nothing) As Integer
    Dim vForm As frmCardMaintenance
    Dim vResult As Integer

    Dim vCursor As New BusyCursor
    Try
      Dim vCustomiseMenu As New CustomiseMenu

      vForm = New frmCardMaintenance(pContactInfo, pList, pOwner Is Nothing)
      vCustomiseMenu.SetContext(vForm, vForm.MaintenanceType, pContactInfo.ContactGroup)

      vForm.SetCustomiseMenu(vCustomiseMenu)
      'BR14553 Set TopMost property to false
      vForm.TopMost = False
      If pOwner Is Nothing Then
        vForm.Show()
      Else
        If vForm.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
          vResult = pContactInfo.ContactNumber
        End If
      End If
    Finally
      vCursor.Dispose()
    End Try
    Return vResult
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes) As Integer
    Return ShowFinder(pType, Nothing, Nothing, False, False, False)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pEnableCheckboxSelect As Boolean) As Integer
    Return ShowFinder(pType, Nothing, Nothing, False, False, pEnableCheckboxSelect)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pList As ParameterList) As Integer
    Return ShowFinder(pType, pList, Nothing, False, False, False)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pList As ParameterList,
                                    ByVal pOwner As Form) As Integer
    Return ShowFinder(pType, pList, pOwner, False, False, False)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pOwner As Form,
                                    ByVal pAllContactGroups As Boolean) As Integer
    Return ShowFinder(pType, Nothing, pOwner, False, False, False)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pList As ParameterList,
                                    ByVal pOwner As Form,
                                    ByVal pAllContactGroups As Boolean,
                                    ByRef pNewContactAtOrg As Boolean) As Integer
    Return ShowFinder(pType, pList, pOwner, pAllContactGroups, pNewContactAtOrg, False)
  End Function

  Public Shared Function ShowFinder(ByVal pType As CareServices.XMLDataFinderTypes,
                                    ByVal pList As ParameterList,
                                    ByVal pOwner As Form,
                                    ByVal pAllContactGroups As Boolean,
                                    ByRef pNewContactAtOrg As Boolean,
                                    ByVal pMultipleSelect As Boolean) As Integer
    Dim vFinder As frmFinder = Nothing
    Dim vResult As Integer

    Dim vCursor As New BusyCursor
    Try
      If pOwner Is Nothing Then
        Dim vFound As Boolean
        For Each vForm As Form In MainHelper.Forms
          If TypeName(vForm) = GetType(frmFinder).Name Then
            Dim vFinderForm As frmFinder = DirectCast(vForm, frmFinder)
            If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_contact_finder_new) AndAlso (pType = CareNetServices.XMLDataFinderTypes.xdftContacts OrElse pType = CareNetServices.XMLDataFinderTypes.xdftOrganisations) Then
              If vFinderForm.FinderType = pType AndAlso vFinderForm.IsMainFinder Then
                vFinder = vFinderForm
                vFinder.Close()
                vFound = False
              End If
            Else
              If vFinderForm.FinderType = pType AndAlso vFinderForm.IsMainFinder Then
                vFinder = vFinderForm
                vFinder.BringToFront()
                vFound = True
              End If
            End If
          End If
        Next
        If Not vFound Then
          vFinder = New frmFinder(pType, pList, pAllContactGroups)
          vFinder.MultipleSelect = pMultipleSelect
          vFinder.IsMainFinder = True
          MainHelper.SetMDIParent(vFinder)
          vFinder.Show()
        End If
      Else
        vFinder = New frmFinder(pType, pList, pAllContactGroups)
        'If this is the document finder and a parent is passed in then we can get a problem since we may choose to edit
        'one of the documents we have found and this will result in opening a non-modal form underneath this one (modal)
        'Accordingly I have removed the parent from calls to show finder for the document finder apart from those from the TextLookupBox
        'For these we disable the finder popup menus
        If vFinder.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
          vResult = vFinder.SelectedItem
          pNewContactAtOrg = vFinder.NewContactAtOrg      'Flag if we just added a contact to a given organisation
        End If
      End If
      If vFinder IsNot Nothing Then
        vFinder.MultipleSelect = pMultipleSelect
      End If
    Finally
      vCursor.Dispose()
    End Try
    Return vResult
  End Function

  Public Shared Function ShowCriteriaLists(ByVal pListManager As Boolean) As Integer
    Dim vForm As frmCriteriaLists
    Dim vResult As String = ""

    Dim vCursor As New BusyCursor
    Try
      Dim vMailingInfo As New MailingInfo()
      vMailingInfo.Init("GM", 0)
      vForm = New frmCriteriaLists(vMailingInfo, True)
      vForm.ShowDialog()
      Return vForm.CriteriaSet
    Finally
      vCursor.Dispose()
    End Try
  End Function

  Public Shared Sub CloseOpenBatch(ByVal pBatchNumber As Integer)
    Dim vList As New ArrayListEx
    vList.Add(pBatchNumber)
    CloseOpenBatch(vList)
  End Sub

  Public Shared Sub CloseOpenBatch(ByVal pBatchNumbers As ArrayListEx)
    Dim vBatchNumber As Integer
    Dim vBatchNumbers As ArrayListEx = Nothing
    Dim vBatchNo As Integer

    vBatchNumbers = pBatchNumbers
    If Not vBatchNumbers Is Nothing Then
      For Each vBatchNo In vBatchNumbers
        vBatchNumber = vBatchNo
        Dim vList As New ParameterList(True)
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_manual_paying_in_slips) AndAlso AppValues.ConfigurationValue(AppValues.ConfigurationValues.manual_pis_location, "CLOSEBATCH") = "CLOSEBATCH" Then
          Dim vBatchInfo As New BatchInfo(vBatchNumber)
          If vBatchInfo.BatchType = CareNetServices.BatchTypes.Cash OrElse vBatchInfo.BatchType = CareNetServices.BatchTypes.CashWithInvoice Then

            Dim vForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptPayingInSlipNumber, Nothing, Nothing)
            If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
              Dim vReturnList As New ParameterList
              vReturnList = vForm.ReturnList
              vList("PayingInSlipNumber") = vReturnList("PayingInSlipNumber")
            Else
              Continue For
            End If
          End If
        End If
        Try
          vList.IntegerValue("BatchNumber") = vBatchNumber
          If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_auto_batch_processing) Then
            If ShowQuestion(QuestionMessages.QmDoBatchProcessing, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
              vList("ProcessBatch") = "Y"
            End If
          End If
          DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoCloseOpenBatch, vList)
        Catch vEx As CareException
          If vEx.ErrorNumber = CareException.ErrorNumbers.enUnbalancedBatch OrElse
             vEx.ErrorNumber = CareException.ErrorNumbers.enJobErrors Then
            ShowInformationMessage(vEx.Message)
          Else
            DataHelper.HandleException(vEx)
          End If
        Catch vException As Exception
          DataHelper.HandleException(vException)
        End Try
      Next
    End If
  End Sub

  Public Shared Function ShowBatchFinder(ByRef pBatchNumber As Integer, Optional ByVal pList As ParameterList = Nothing, Optional ByVal pOwner As Form = Nothing, Optional ByVal pOpenBatches As Boolean = False, Optional ByRef pBatchNumbers As ArrayListEx = Nothing) As Boolean
    Dim vFinder As frmFinder
    Dim vResult As Boolean

    pBatchNumber = 0
    Dim vCursor As New BusyCursor
    Try
      If pOpenBatches Then
        vFinder = New frmFinder(CareServices.XMLDataFinderTypes.xdftOpenBatches, pList)
      Else
        vFinder = New frmFinder(CareServices.XMLDataFinderTypes.xdftBatches, pList)
      End If
      If pOwner Is Nothing Then
        MainHelper.SetMDIParent(vFinder)
        vFinder.Show()
      ElseIf pOpenBatches Then
        vFinder.ShowDialog(pOwner)
      Else
        If vFinder.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
          pBatchNumber = vFinder.SelectedItem
          vResult = True
        End If
        vFinder.Close()
      End If
    Finally
      vFinder = Nothing
      vCursor.Dispose()
    End Try
    Return vResult
  End Function

  Public Shared Function RunMailing(ByVal pType As CareServices.TaskJobTypes, Optional ByVal pDefaults As ParameterList = Nothing, Optional ByVal vRunMailing As Boolean = True) As RunMailingResult
    Dim vList As ParameterList = Nothing
    Dim vCampaignMailing As Boolean
    Dim vCampaignCount As Boolean
    Dim vMailing As Boolean
    If pDefaults IsNot Nothing AndAlso pDefaults.Contains("Campaign") = True AndAlso pDefaults.Contains("Appeal") AndAlso pDefaults.Contains("Segment") Then
      vCampaignMailing = True
      If pDefaults.Item("Mail") = "N" Then vCampaignCount = True
    ElseIf pDefaults IsNot Nothing AndAlso pDefaults.Contains("GeneralMailing") Then
      'vCampaignMailing = True
      vMailing = True
      'pDefaults.Remove("GeneralMailing")
      vList = pDefaults
    End If
    If vCampaignCount Then
      vList = New ParameterList(True, True)
    ElseIf Not vMailing Then
      If vRunMailing Then
        vList = FormHelper.ShowApplicationParameters(pType, pDefaults)
      Else
        vList = pDefaults 'just set the values without showing the form
      End If
    End If
    If vList IsNot Nothing AndAlso vList.Count > 0 Then
      If pType = CareNetServices.TaskJobTypes.tjtSelectMailing Then
        pType = CType([Enum].Parse(GetType(CareNetServices.TaskJobTypes), vList("MailingType")), CDBNETCL.CareNetServices.TaskJobTypes)
        vList.Remove("MailingType")
      End If
      Dim vContinue As Boolean = True
      Dim vPanelItems As PanelItems = Nothing

      If vCampaignMailing Then
        For Each vParam As DictionaryEntry In pDefaults    'Copy all parameters from pDefaults to vList
          If Not vList.ContainsKey(vParam.Key) Then
            vList.Add(vParam.Key, vParam.Value)
          End If
        Next
        'This next call will lock an appeal if we are doing a mailing for it
        Dim vDataSet As DataSet = DataHelper.GetCampaignCriteriaVariableControls(vList, pType)
        Dim vControls As New PanelItems("CriteriaVariables")    'Add variables controls to PanelItems if any
        If vDataSet IsNot Nothing Then
          Dim vRow As DataRow
          Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
          If vTable IsNot Nothing Then
            For Each vRow In vTable.Rows
              'BR17896 - Client side validation added. Avoid Server side errors.
              'BR18706 - Can contain letters, digits, hyphens, underscores, and periods.
              'BR20699 - Allow for BR16016 changes i.e. a control caption 'From ($TODAY)' 
              Dim vRegEx As Regex = New Regex("[a-zA-Z0-9_.-]") 'Only allow alphanumeric, _ - .
              If Left(vRow("ControlCaption").ToString, 12).ToUpper <> "FROM ($TODAY" Then  'i.e. not Caption generated from server side  
                Dim vInvalidCharacters As String = vRegEx.Replace((vRow("ControlCaption").ToString).Remove(0, 1), "")
                If (vInvalidCharacters.Length > 0) Then
                  Throw (New CareException(Utilities.GetInformationMessage(InformationMessages.ImInvalidVariableName, vInvalidCharacters), CareException.ErrorNumbers.enVariableNameContainsInvalidCharacters))
                End If
              End If
              Dim vPanelItem As New PanelItem(vRow)
              vControls.Add(vPanelItem)
            Next
          End If
          If vDataSet.Tables.Contains("Parameters") Then          'Add additional parameters for Organisation Selection if any
            vRow = vDataSet.Tables("Parameters").Rows(0)
            vList.Add("OrganisationCriteriaCount", vRow("OrganisationCriteriaCount").ToString)
            vList.Add("ContactCriteriaCount", vRow("ContactCriteriaCount").ToString)
            vList.Add("OrgMailTo", vRow("OrgMailTo").ToString)
            vList.Add("OrgMailWhere", vRow("OrgMailWhere").ToString)
            vList.Add("OrgRoles", vRow("OrgRoles").ToString)
            vList.Add("OrgAddressUsage", vRow("OrgAddressUsage").ToString)
            vList.Add("OrgLabelName", vRow("OrgLabelName").ToString)
          End If
        End If
        vPanelItems = vControls
      Else
        If vList.Contains("CriteriaSet") AndAlso vList.IntegerValue("CriteriaSet") > 0 AndAlso vMailing = False Then vPanelItems = DataHelper.GetCriteriaVariableControls(vList.IntegerValue("CriteriaSet"), pType)
      End If
      If vPanelItems IsNot Nothing AndAlso vPanelItems.Count > 0 Then
        If pDefaults IsNot Nothing Then
          Dim vPanelItem As PanelItem
          For Each vItem As DictionaryEntry In pDefaults
            vPanelItem = vPanelItems.ItemByProperName(vItem.Key.ToString)
            If vPanelItem IsNot Nothing Then vPanelItem.DefaultValue = vItem.Value.ToString
          Next
        End If
        Dim vVariableList As ParameterList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optCriteriaVariables, vPanelItems, Nothing, "")
        If vVariableList IsNot Nothing AndAlso vVariableList.Count > 0 Then
          For Each vValue As DictionaryEntry In vVariableList
            If Not vList.Contains(vValue.Key.ToString) Then vList.Add(vValue.Key.ToString, vValue.Value.ToString)
          Next
        Else
          vContinue = False
        End If
      End If
      If vContinue Then
        Dim vFileName As String = ""
        If Not vCampaignCount Then
          If vList.ContainsKey("ReportDestination") Then
            vFileName = vList("ReportDestination")
            vList.Remove("ReportDestination")
          End If
        End If
        If vList.ContainsKey("OrganisationCriteriaCount") AndAlso vList.IntegerValue("OrganisationCriteriaCount") > 0 Then
          Dim vOrgSelectList As ParameterList
          Dim vPassList As ParameterList = Nothing
          If vCampaignMailing Then
            vPassList = vList
          Else
            vPassList = New ParameterList(True)
          End If
          If vList.IntegerValue("ContactCriteriaCount") > 0 Then
            vOrgSelectList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optAddressSelectionOptions, Nothing, vPassList, "")
          Else
            If Not vCampaignMailing Then
              vPassList("MailingApplicationCode") = AppValues.MailingApplicationCode(pType)
              vPassList("ApplicationName") = vList("ApplicationName")
            End If
            vOrgSelectList = FormHelper.ShowApplicationParameters(EditPanelInfo.OtherPanelTypes.optOrganisationSelectionOptions, Nothing, vPassList, "")
          End If
          vContinue = AddOrgMailingParameters(vOrgSelectList, vList)
        End If
        If vContinue Then
          Dim vScheduleOnly As Boolean
          Dim vRunAsynch As Boolean = Not vCampaignCount AndAlso Not vMailing        'Make counts synchronous?
          If vRunMailing Then
            If vCampaignMailing Then
              If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciScheduleTasks) Then
                Dim vScheduledList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptScheduledJobDetails)
                If vScheduledList.Count = 0 Then
                  Return RunMailingResult.NoMailingRun
                ElseIf vScheduledList.Contains("Schedule") Then
                  If vScheduledList.Contains("JobProcessor") Then vList.Add("JobProcessor", vScheduledList("JobProcessor").ToString)
                  If vScheduledList.Contains("NotifyStatus") Then vList.Add("NotifyStatus", vScheduledList("NotifyStatus").ToString)
                  vList.Add("DueDate", vScheduledList("DueDate").ToString)
                  vList.Add("JobFrequency", vScheduledList("JobFrequency").ToString)
                  If BooleanValue(vScheduledList("UpdateJobParameterDates").ToString) Then
                    vList.Add("UpdateJobParameterDates", "A")
                  Else
                    vList.Add("UpdateJobParameterDates", "N")
                  End If
                  vRunAsynch = False        'Schedule only so make synchronous
                  vScheduleOnly = True
                End If
              End If
            End If
          Else
            vList.Add("CreateJob", "N") 'add extra parameter so job is not created
          End If
          ''In Membership Card Production when merging to a standard document we need to generate the mailing number on the client-side, passing this to the MembershipCards 'MCMM' report,  
          ''such that we have the mailing number for when Updating the Mailing Document File in ProcessMailingHandler
          'Dim vGenerateMailingNumber As Boolean = (pType = CareNetServices.TaskJobTypes.tjtMembCardMailing AndAlso pDefaults.Contains("StandardDocument"))
          'BR18224: The vGenerateMailingNumber parameter has been reinstated as it is required for Standard Documents
          'Dim vResults As MailingResults = DataHelper.GetMailingFile(vList, vFileName, vCampaignMailing, Not vRunAsynch)
          Dim vGenerateMailingNumber As Boolean = (pType = CareNetServices.TaskJobTypes.tjtMembCardMailing OrElse (pDefaults IsNot Nothing AndAlso pDefaults.Contains("StandardDocument")))
          Dim vResults As MailingResults = DataHelper.GetMailingFile(vList, vFileName, vCampaignMailing, Not vRunAsynch, vGenerateMailingNumber)
          If vRunMailing Then
            If vRunAsynch Then
              Return RunMailingResult.MailingRunAsych
            Else
              If vResults.Success Then
                Return RunMailingResult.MailingRunSynchSuccess
              Else
                Return RunMailingResult.MailingRunSynchFail
              End If
            End If
          End If
        End If
      End If
    End If
    Return RunMailingResult.NoMailingRun
  End Function

  Public Shared Function AddOrgMailingParameters(ByVal pOrgSelectList As ParameterList, ByVal pList As ParameterList) As Boolean
    If pOrgSelectList IsNot Nothing AndAlso pOrgSelectList.Count > 0 Then
      'add the items in pOrgSelectList to vLlist
      If pOrgSelectList.Contains("OptionMailTo") Then pList("OrgMailTo") = pOrgSelectList("OptionMailTo")
      If pOrgSelectList.Contains("OptionMailWhere") Then
        If pOrgSelectList.Item("OptionMailWhere") <> "U" Then pList.Item("OrgAddressUsage") = ""
        pList("OrgMailWhere") = pOrgSelectList("OptionMailWhere")
      End If
      If pOrgSelectList.Contains("AddressUsage") Then pList("OrgAddressUsage") = pOrgSelectList("AddressUsage")
      If pOrgSelectList.Contains("IncludeHistoricRoles") Then pList("OrgIncludeHistoricRoles") = pOrgSelectList("IncludeHistoricRoles")
      If pOrgSelectList.Contains("LabelName") Then pList("OrgLabelName") = pOrgSelectList("LabelName")
      If pOrgSelectList.Contains("IncludedRoles") Then pList("OrgRoles") = pOrgSelectList("IncludedRoles")
      Return True
    Else
      Return False
    End If
  End Function

  Public Shared Sub ProcessTask(ByVal pTaskJobType As CareServices.TaskJobTypes, ByVal pDefaults As ParameterList)
    ProcessTask(pTaskJobType, pDefaults, True, ProcessTaskScheduleType.ptsNone, False)
  End Sub
  Public Shared Sub ProcessTask(ByVal pTaskJobType As CareServices.TaskJobTypes, ByVal pDefaults As ParameterList, ByVal pParamsEntry As Boolean, ByVal pScheduleType As ProcessTaskScheduleType)
    ProcessTask(pTaskJobType, pDefaults, pParamsEntry, pScheduleType, False)
  End Sub
  Public Shared Sub ProcessTask(ByVal pTaskJobType As CareServices.TaskJobTypes, ByVal pDefaults As ParameterList, ByVal pParamsEntry As Boolean, ByVal pScheduleType As ProcessTaskScheduleType, ByVal pRunAsynchronously As Boolean)
    Dim vIsNetJob As Boolean = False

    If pDefaults Is Nothing Then pDefaults = New ParameterList

    Select Case pTaskJobType
      Case CareServices.TaskJobTypes.tjtAmendmentHistoryView
        pDefaults("Date") = DateTime.Today.AddDays(-7).ToString(AppValues.DateFormat)
        pDefaults("Date2") = DateTime.Today.AddDays(1).ToString(AppValues.DateFormat)
      Case CareServices.TaskJobTypes.tjtBatchPurge
        pDefaults("BatchDate") = DateTime.Today.AddYears(-7).ToString(AppValues.DateFormat)
      Case CareServices.TaskJobTypes.tjtBackOrderPurge
        pDefaults("TransactionDate") = DateTime.Today.AddYears(-7).ToString(AppValues.DateFormat)
      Case CareServices.TaskJobTypes.tjtPickingAndDespatchPurge
        pDefaults("ConfirmedOn") = DateTime.Today.AddYears(-7).ToString(AppValues.DateFormat)
    End Select
    Dim vParamList As ParameterList
    If pParamsEntry Then
      Select Case pTaskJobType
        Case CareServices.TaskJobTypes.tjtConfirmStockAllocation
          vParamList = FormHelper.ShowApplicationParameters(pTaskJobType, pDefaults)
          If vParamList IsNot Nothing AndAlso vParamList.Count > 0 AndAlso vParamList("Checkbox") = "N" Then
            ProcessShortfall(vParamList)
            Exit Sub
          End If
        Case Else
          vParamList = FormHelper.ShowApplicationParameters(pTaskJobType, pDefaults)
          If vParamList IsNot Nothing AndAlso vParamList.Count > 0 Then
            Dim vMailsort As Boolean = False
            If (pTaskJobType = CareServices.TaskJobTypes.tjtBallotPaperProduction And vParamList.Contains("Checkbox") AndAlso vParamList("Checkbox") = "Y") Or _
            (pTaskJobType = CareServices.TaskJobTypes.tjtRenewalsAndReminders And vParamList.Contains("Checkbox4") AndAlso vParamList("Checkbox4") = "Y") Then
              vMailsort = True
              If pTaskJobType = CareNetServices.TaskJobTypes.tjtRenewalsAndReminders Then
                If (vParamList.Contains("ReportDestination") AndAlso vParamList("ReportDestination").ToString = "None") AndAlso _
                  (vParamList.Contains("ReportDestination3") AndAlso vParamList("ReportDestination3").ToString = "None") Then
                  'If neither Report Destination nor Gift Member Pack Destination is being saved then don't generate Mail Sort output
                  vMailsort = False
                End If
              End If
              If vMailsort Then
                'Mailsort mailing
                Dim vMailingParams As New ParameterList
                vMailingParams.FillFromXMLString(vParamList.XMLParameterString)
                Dim vMailingParamsNew As New ParameterList
                Dim vIgnoredParameters() As String = {"Culture", "Database", "UserLogname"}
                For Each vItem As String In vMailingParams.Keys
                  Dim vAdd As Boolean = True
                  vAdd = Not vIgnoredParameters.Contains(vItem)
                  If vAdd Then
                    If vItem.StartsWith("ReportDestination") Then
                      If GetReportDestinationType(vMailingParams(vItem).ToString) <> "Save" Then
                        vAdd = False
                      End If
                    End If
                  End If
                  If vAdd Then
                    vMailingParamsNew.Add(vItem, vMailingParams(vItem).ToString)
                  End If
                Next
                vMailingParams = vMailingParamsNew
                If vMailingParams.ContainsKey("ReportDestination3") Then
                  vMailingParams("ReportDestination2") = vMailingParams("ReportDestination3")
                  vMailingParams.Remove("ReportDestination3")
                Else
                  vMailingParams("ReportDestination2") = String.Empty
                End If
                vMailingParams = FormHelper.ShowApplicationParameters(CareServices.TaskJobTypes.tjtMailingRun, vMailingParams)

                If vMailingParams.Count > 0 Then
                  Dim vTemp As String = ""
                  Dim vTemp1 As String = ""
                  Dim vNewReportDestParam As String = ""
                  Dim vGiftMemReportDest As String = ""
                  Dim vReportDestParamExists As Boolean = False
                  'Add the new parameters to the existing collection
                  For Each vParamName As String In vMailingParams.Keys

                    vTemp = vParamName
                    If IsNumeric(vTemp.Substring(vTemp.Length - 1, 1)) Then
                      vTemp = vTemp.Remove(vTemp.Length - 1)
                    End If
                    If vParamList.Contains(vTemp) Then
                      If vParamName = "ReportDestination" Then
                        vReportDestParamExists = True
                      End If
                      Dim vCount As Integer = 1
                      vTemp1 = vTemp
                      Do
                        vCount += 1
                        vTemp = vTemp1 & vCount
                        If vParamName = "ReportDestination" Then
                          vNewReportDestParam = vTemp
                        ElseIf vParamName = "ReportDestination2" Then
                          vGiftMemReportDest = vTemp
                        End If
                      Loop While vParamList.Contains(vTemp)
                    End If
                    vParamList.Add(vTemp, vMailingParams(vParamName).ToString)
                  Next

                  If vReportDestParamExists Then
                    'The existing ReportDestination may have been changed so update
                    vParamList("ReportDestination") = vParamList(vNewReportDestParam).ToString
                    If pTaskJobType = CareNetServices.TaskJobTypes.tjtRenewalsAndReminders Then
                      If vParamList.Contains("ReportDestination3") AndAlso vMailingParams.Contains("ReportDestination2") Then
                        vParamList("ReportDestination3") = vParamList(vGiftMemReportDest).ToString
                      End If
                    End If
                  End If
                Else
                  'User has press cancel on the mailing option dialog. Prevent mail sort.
                  vMailsort = False
                End If
              End If
              If Not vMailsort Then
                'Set mail sort option to null
                If pTaskJobType = CareServices.TaskJobTypes.tjtBallotPaperProduction Then vParamList("Checkbox") = "N"
                If pTaskJobType = CareServices.TaskJobTypes.tjtRenewalsAndReminders Then vParamList("Checkbox4") = "N"
              End If
            End If
            If pTaskJobType = CareServices.TaskJobTypes.tjtManualSOReconciliation And pDefaults.ContainsKey("TTYNumber") Then
              vParamList("TTYNumber") = pDefaults("TTYNumber").ToString
            End If
          End If
      End Select
    Else
      vParamList = pDefaults
    End If
    Dim vProcessNETJob As Boolean
    If vParamList IsNot Nothing AndAlso vParamList.Count > 0 Then
      Select Case pTaskJobType
        Case CareServices.TaskJobTypes.tjtPublicCollectionsFulfilment
          vParamList("FulfilmentType") = pDefaults("FulfilmentType")
        Case CareServices.TaskJobTypes.tjtCheetahMailTotals
          ' Indicate that we are doing all CM mailing total
          vParamList("AllCMTotals") = "Y"
          vIsNetJob = True
        Case CareServices.TaskJobTypes.tjtCheetahMailMetaData, CareServices.TaskJobTypes.tjtCheetahMailEventData, CareServices.TaskJobTypes.tjtEMailProcessor, _
        CareServices.TaskJobTypes.tjtSetBoxesArrived, CareNetServices.TaskJobTypes.tjtIssueResources
          vIsNetJob = True
        Case CareServices.TaskJobTypes.tjtPrintBoxLabels, CareServices.TaskJobTypes.tjtShipDistributionBoxes
          vProcessNETJob = True
        Case CareServices.TaskJobTypes.tjtGiftAidPotentialClaim, _
          CareServices.TaskJobTypes.tjtGiftAidClaim, CareServices.TaskJobTypes.tjtGASPotentialClaim,
          CareServices.TaskJobTypes.tjtGASTaxClaim, CareServices.TaskJobTypes.tjtIrishGiftAidPotentialClaim,
          CareServices.TaskJobTypes.tjtIrishGiftAidTaxClaim, CareServices.TaskJobTypes.tjtRenewalsAndReminders,
          CareServices.TaskJobTypes.tjtDirectDebitRun, CareServices.TaskJobTypes.tjtDatabaseUpgrade
          If pDefaults.ContainsKey("ShowTaskStatus") Then vParamList("ShowStatus") = pDefaults("ShowTaskStatus")
          If pDefaults.ContainsKey("DBUpgrade") Then vParamList("DBUpgrade") = pDefaults("DBUpgrade")
        Case CareNetServices.TaskJobTypes.tjtCreateJournalFiles
          If vParamList("ReportDestination") = "Print" OrElse vParamList("ReportDestination") = "Preview" OrElse _
             vParamList("ReportDestination2") = "Print" OrElse vParamList("ReportDestination2") = "Preview" Then
            pScheduleType = ProcessTaskScheduleType.ptsAlwaysRun
          End If
      End Select
      'Case CareServices.TaskJobTypes.tjtBulkContactDeletion, CareServices.TaskJobTypes.tjtChequeProduction
      Dim vResult As DialogResult = DialogResult.Cancel
      If pScheduleType = ProcessTaskScheduleType.ptsAskToSchedule Then
        vResult = ScheduleTask(vParamList)
      ElseIf pScheduleType = ProcessTaskScheduleType.ptsAlwaysRun Then
        vResult = DialogResult.No
      ElseIf pScheduleType = ProcessTaskScheduleType.ptsAlwaysSchedule Then
        vResult = DialogResult.Yes
        For Each vKey As String In pDefaults.Keys     'Get Schedule parameters from pDefaults as ScheduleTask has already been called
          If vParamList.Contains(vKey) = False Then vParamList(vKey) = pDefaults(vKey)
        Next
      End If

      'Always add DestinationType parameters when running a job
      If vResult = DialogResult.No AndAlso vParamList.Contains("ReportDestination") Then
        vParamList("DestinationType") = GetReportDestinationType(vParamList("ReportDestination"))
        If vParamList.Item("DestinationType") <> "None" Then PrintHandler.GetDefaultPrinterParameters(vParamList)
      End If

      vParamList("JobName") = GetTaskJobTypeName(pTaskJobType)

      If (vResult = DialogResult.No AndAlso pRunAsynchronously) OrElse (vIsNetJob AndAlso vResult <> DialogResult.Cancel) Then
        Dim vProcessor As New AsyncProcessHandler(pTaskJobType, vParamList)
        AddHandler vProcessor.ProcessCompleted, AddressOf ProcessJobCompleted
        vProcessor.ProcessJob()
        MainHelper.SetTaskNotificationTimer()

        If pTaskJobType = CareNetServices.TaskJobTypes.tjtDatabaseUpgrade OrElse
          (vParamList.ContainsKey("ShowTaskStatus") AndAlso vParamList("ShowTaskStatus") = "Y") Then
          'Remove key as it is not required at the server side          
          Dim vTaskInfo As New frmTaskInfo(pTaskJobType)
          vTaskInfo.Show()
        End If
        Exit Sub
      End If
      'If we have already process this job above( in .NET) then exit SUB and don't process it again (in VB6)
      If vResult = DialogResult.Cancel AndAlso pScheduleType <> ProcessTaskScheduleType.ptsNone Then Exit Sub
      ProcessJob(pTaskJobType, vParamList)
    End If
  End Sub

  Private Shared Sub ProcessJob(ByVal pTaskJobType As CareServices.TaskJobTypes, ByVal pList As ParameterList)
    If Not pList.ContainsKey("JobName") Then pList("JobName") = GetTaskJobTypeName(pTaskJobType)
    Dim vDataSet As DataSet = DataHelper.ProcessJob(pTaskJobType, pList)
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(vDataSet)
    ProcessTaskResults(pTaskJobType, pList, vTable, vDataSet)
  End Sub

  Private Shared Sub ProcessJobCompleted(ByVal pJob As AsyncProcessHandler)
    ProcessTaskResults(pJob.TaskJobType, pJob.ParameterList, pJob.ResultTable, pJob.ResultDataSet)
  End Sub

  Public Shared Sub ProcessTaskResults(ByVal pTaskJobType As CareServices.TaskJobTypes, ByVal pParam As ParameterList, ByVal pResultTable As DataTable, ByVal pResultDataSet As DataSet)
    If pResultTable IsNot Nothing AndAlso pResultTable.Rows.Count > 0 Then
      Dim vRow As DataRow = pResultTable.Rows(0)
      For Each vColumn As DataColumn In pResultTable.Columns
        If vColumn.ColumnName.StartsWith("ReportDestination") OrElse vColumn.ColumnName.StartsWith("MailingFileName") Then
          'If the job ran a report or generated an output file then we may need to print it or show in print preview
          If pResultDataSet.Tables.Contains(vColumn.ColumnName) Then
            'If we have a table with the same name as the result column then we need to show it or print it
            Dim vList As New ParameterList(True)
            If vRow(vColumn.ColumnName).ToString = "Print" Then vList("DestinationType") = "PrintXML"
            Call (New PrintHandler).PrintReport(vList, pResultDataSet.Tables(vColumn.ColumnName))
          End If
        End If
      Next
      MainHelper.RefreshData(pTaskJobType)
      If pParam.Contains("ShowResultMessage") = False Then pParam.Add("ShowResultMessage", "Y")
      If BooleanValue(pParam("ShowResultMessage")) Then ShowInformationMessage(vRow("ResultStatus").ToString)
      'BR16745 hardcoded case where the .csv has been recreated but with no data so need the process to stop
      If pTaskJobType = CareNetServices.TaskJobTypes.tjtCreditStatementGeneration And vRow("ResultStatus").ToString().Contains("No credit statements match") Then Exit Sub
      If pTaskJobType = CareServices.TaskJobTypes.tjtBulkContactDeletion Then MainHelper.RefreshHistoryData(HistoryEntityTypes.hetSelectionSets, IntegerValue(pParam("SelectionSet")))
      If pTaskJobType = CareNetServices.TaskJobTypes.tjtCreditStatementGeneration Then
        If My.Computer.FileSystem.GetFileInfo(pParam("ReportDestination")).Length > 0 Then
          If pParam.Contains("StandardDocument") AndAlso pParam("StandardDocument").Length > 0 Then
            Dim vLookupList As New ParameterList(True)
            vLookupList("StandardDocument") = pParam("StandardDocument").ToString
            Dim vDataRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vLookupList).Rows(0)
            Dim vApplication As ExternalApplication = GetDocumentApplication(vDataRow.Item("DocfileExtension").ToString)
            vApplication.MergeStandardDocument(vDataRow.Item("StandardDocument").ToString, vDataRow.Item("DocfileExtension").ToString, pParam("ReportDestination"), False)
          End If
        End If
      End If
    End If
  End Sub

  Public Shared Function ShowApplicationParameters(ByVal pPanelType As EditPanelInfo.OtherPanelTypes, ByVal pPanelItems As PanelItems, ByVal pDefaults As ParameterList, ByVal pCaption As String) As ParameterList
    Return ShowApplicationParameters(pPanelType, pPanelItems, pDefaults, pCaption, CareNetServices.TaskJobTypes.tjtNone)
  End Function


  Public Shared Function ShowApplicationParameters(ByVal pPanelType As EditPanelInfo.OtherPanelTypes, ByVal pPanelItems As PanelItems, ByVal pDefaults As ParameterList, ByVal pCaption As String, ByVal pTaskJobType As CareNetServices.TaskJobTypes) As ParameterList
    Dim vCursor As New BusyCursor
    Dim vAppParameters As frmApplicationParameters
    Dim vReturnList As ParameterList = Nothing

    Try
      vAppParameters = New frmApplicationParameters(pPanelType, pPanelItems, pDefaults, pCaption, pTaskJobType)
      If vAppParameters.ShowDialog(CurrentMainForm) = DialogResult.OK Then
        vReturnList = vAppParameters.ReturnList
      End If
      Return vReturnList
    Catch vException As Exception
      DataHelper.HandleException(vException)
      Return Nothing
    Finally
      vCursor.Dispose()
    End Try
  End Function
  Public Shared Function ShowApplicationParameters(ByVal pTaskJobType As CareServices.TaskJobTypes, Optional ByVal pDefaults As ParameterList = Nothing) As ParameterList
    Dim vCursor As New BusyCursor
    Dim vAppParameters As frmApplicationParameters
    Dim vReturnList As New ParameterList

    Try
      vAppParameters = New frmApplicationParameters(pTaskJobType, pDefaults)
      Select Case pTaskJobType
        Case CareServices.TaskJobTypes.tjtRenewalsAndReminders, CareServices.TaskJobTypes.tjtDirectDebitRun, CareServices.TaskJobTypes.tjtBulkGiftAidUpdate, CareNetServices.TaskJobTypes.tjtIssueResources
          AddHandler vAppParameters.ProcessCountTask, AddressOf ProcessJob
        Case CareServices.TaskJobTypes.tjtDirectDebitMailing, CareServices.TaskJobTypes.tjtStandingOrderMailing, _
            CareServices.TaskJobTypes.tjtMemberMailing, CareServices.TaskJobTypes.tjtPayerMailing, _
            CareServices.TaskJobTypes.tjtSubscriptionMailing, CareServices.TaskJobTypes.tjtSelectionManagerMailing, _
            CareServices.TaskJobTypes.tjtMembCardMailing, CareServices.TaskJobTypes.tjtPayrollPledgeMailing, _
            CareServices.TaskJobTypes.tjtMemberFulfilment, CareServices.TaskJobTypes.tjtStandingOrderCancellation, _
            CareServices.TaskJobTypes.tjtIrishGiftAidMailing, CareServices.TaskJobTypes.tjtSelectionTester
          AddHandler vAppParameters.ProcessMailingCriteria, AddressOf ProcessMailingCriteriaFromAppParam
      End Select
      If vAppParameters.HasControls AndAlso vAppParameters.ShowDialog(CurrentMainForm) = DialogResult.OK Then
        vReturnList = vAppParameters.ReturnList
      End If
      Return vReturnList

    Catch vException As Exception
      DataHelper.HandleException(vException)
      Return Nothing
    Finally
      vCursor.Dispose()
    End Try

  End Function
  Public Shared Function ShowApplicationParameters(ByVal pFunctionParameterType As CareServices.FunctionParameterTypes, Optional ByVal pDefaults As ParameterList = Nothing, Optional ByVal pRenameList As ParameterList = Nothing, Optional ByVal pParent As MaintenanceParentForm = Nothing) As ParameterList
    Dim vCursor As New BusyCursor
    Dim vAppParameters As frmApplicationParameters
    Dim vReturnList As New ParameterList
    Dim vForm As Form
    Try
      vAppParameters = New frmApplicationParameters(pFunctionParameterType, pDefaults, pRenameList)
      If pParent IsNot Nothing Then
        vForm = pParent
      Else
        vForm = CurrentMainForm
      End If
      If Not vAppParameters.HasControls AndAlso pFunctionParameterType = CareServices.FunctionParameterTypes.fptScheduledJobDetails Then
        'Handle the scenario where the controls for scheduling have been deleted 
        vReturnList("RunNow") = "R"
      Else
        If vAppParameters.IsValid Then
          If vAppParameters.ShowDialog(vForm) = DialogResult.OK Then
            vReturnList = vAppParameters.ReturnList
          End If
        End If
      End If
      Return vReturnList

    Catch vException As Exception
      DataHelper.HandleException(vException)
      Return Nothing
    Finally
      vCursor.Dispose()
    End Try

  End Function

  Public Shared Sub ShowExamIndex()
    Dim vCursor As New BusyCursor
    Try
      Dim vFound As Boolean
      Dim vExamSet As frmExams
      For Each vForm As Form In MainHelper.Forms
        If TypeName(vForm) = GetType(frmExams).Name Then
          vExamSet = DirectCast(vForm, frmExams)
          vExamSet.BringToFront()
          vFound = True
        End If
      Next
      If Not vFound Then
        'Dim vCampaignItem As New CampaignItem(pCampaign, "", "")
        vExamSet = New frmExams()
        With vExamSet
          .MdiParent = MDIForm
          .Show()
          DirectCast(vExamSet, IPanelVisibility).PanelHasFocus = True
          InitWindowForViewType(vExamSet)
          .BringToFront()
        End With
      End If
    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enExamMaintenanceIncorrectSetup Then
        ShowInformationMessage(vCareEX.Message)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub ShowExamIndex(ByVal pUnitLinkId As Integer, ByVal pType As String)
    Dim vCursor As New BusyCursor
    Try
      Dim vParams As New ParameterList(True)
      Dim vExamUnitLinkId As Integer
      vParams.Add("ExamCentreUnitId", pUnitLinkId)
      Dim vDataTable As DataTable = ExamsDataHelper.SelectExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnits, vParams)
      If vDataTable IsNot Nothing Then
        vExamUnitLinkId = If(vDataTable.Rows(0)("ExamUnitLinkId").ToString.Length > 0, CInt(vDataTable.Rows(0)("ExamUnitLinkId").ToString), pUnitLinkId)
      End If

      Dim vFound As Boolean
      Dim vExamSet As frmExams = Nothing
      For Each vForm As Form In MainHelper.Forms
        If TypeName(vForm) = GetType(frmExams).Name Then
          vExamSet = DirectCast(vForm, frmExams)
          vExamSet.BringToFront()
          InitWindowForViewType(vExamSet)
          vFound = True
        End If
      Next
      If Not vFound Then
        vExamSet = New frmExams()
        With vExamSet
          .MdiParent = MDIForm
          .Show()
          DirectCast(vExamSet, IPanelVisibility).PanelHasFocus = True
          InitWindowForViewType(vExamSet)
          .BringToFront()
        End With
      End If

      Select Case pType
        Case "U"
          vExamSet.SelectExamNode(pUnitLinkId, ExamsAccess.XMLExamDataSelectionTypes.ExamUnits, pType)
        Case "X"
          vExamSet.SelectExamNode(pUnitLinkId, ExamsAccess.XMLExamDataSelectionTypes.ExamCentreUnitDetails, pType)
        Case "N"
          vExamSet.SelectExamNode(pUnitLinkId, ExamsAccess.XMLExamDataSelectionTypes.ExamCentres, pType)
        Case Else
      End Select

    Catch vCareEX As CareException
      If vCareEX.ErrorNumber = CareException.ErrorNumbers.enExamMaintenanceIncorrectSetup Then
        ShowInformationMessage(vCareEX.Message)
      Else
        DataHelper.HandleException(vCareEX)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub


  Public Shared Sub ShowCampaignIndex(ByVal pCampaign As String, ByVal pParentForm As MaintenanceParentForm, ByVal pRestrictions As ParameterList)
    Dim vCursor As New BusyCursor
    Try
      Dim vFound As Boolean
      Dim vCampaignSet As frmCampaignSet
      For Each vForm As Form In MainHelper.Forms
        If TypeName(vForm) = GetType(frmCampaignSet).Name Then
          vCampaignSet = DirectCast(vForm, frmCampaignSet)
          vCampaignSet.BringToFront()
          vFound = True
        End If
      Next
      If Not vFound Then
        Dim vCampaignItem As New CampaignItem(pCampaign, "", "")
        vCampaignSet = New frmCampaignSet(vCampaignItem, pParentForm, pRestrictions)
        With vCampaignSet
          .MdiParent = MDIForm
          InitWindowForViewType(vCampaignSet)
          .Show()
          DirectCast(vCampaignSet, IPanelVisibility).PanelHasFocus = True
          .BringToFront()
        End With
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub ShowEventIndex(ByVal pEventNumber As Integer, Optional ByVal pEventGroup As String = "", Optional ByVal pParentForm As MaintenanceParentForm = Nothing)
    Dim vCursor As New BusyCursor
    Dim vForm As Form
    Dim vEventSet As frmEventSet
    Dim vFound As Boolean
    Try
      Dim vEventInfo As New CareEventInfo(pEventNumber, pEventGroup)
      For Each vForm In MainHelper.Forms
        If TypeName(vForm) = GetType(frmEventSet).Name Then
          vEventSet = DirectCast(vForm, frmEventSet)
          If vEventSet.CareEventInfo.EventGroup = vEventInfo.EventGroup Then
            vEventSet.Init(CType(-1, CareServices.XMLEventDataSelectionTypes), vEventInfo, True)
            vEventSet.BringToFront()
            vFound = True
          End If
        End If
      Next
      If Not vFound Then
        vEventSet = New frmEventSet(pParentForm)
        With vEventSet
          vEventSet.Init(CType(-1, CareServices.XMLEventDataSelectionTypes), vEventInfo, True)
          .MdiParent = MDIForm
          InitWindowForViewType(vEventSet)
          .Show()
          DirectCast(vEventSet, IPanelVisibility).PanelHasFocus = True
          .BringToFront()
        End With
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
        ShowInformationMessage(InformationMessages.ImCannotFindEvent)
      Else
        Throw vException
      End If
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub ShowContactCardIndex(ByVal pContactNumber As Integer)
    ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, pContactNumber, True, False)
  End Sub

  Public Shared Sub ShowContactCardIndex(ByVal pContactNumber As Integer, ByVal pNewWindow As Boolean)
    ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtNone, pContactNumber, True, pNewWindow)
  End Sub

  Public Shared Function ShowCardIndex(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Integer, ByVal pRetainPage As Boolean) As Form
    Return ShowCardIndex(pType, pContactNumber, pRetainPage, False)
  End Function

  Public Shared Function ShowCardIndex(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Integer, ByVal pRetainPage As Boolean, ByVal pNewWindow As Boolean) As Form
    'Figure out what contact group the contact is for 
    'Find the appropriate card index if open and repopulate
    'else create the new card index and populate it
    Dim vCardSet As frmCardSet = Nothing
    Dim vCursor As New BusyCursor
    Try
      Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, pContactNumber)
      Dim vContactInfo As New ContactInfo(vDataSet.Tables("DataRow").Rows(0))
      If CheckAccessRights(vContactInfo) Then
        Dim vFound As Boolean
        Dim vCount As Integer
        Dim vMaxWindows As Integer = IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.maximum_entity_windows, "5"))
        For Each vForm As Form In MainHelper.Forms
          If TypeName(vForm) = GetType(frmCardSet).Name Then
            vCardSet = DirectCast(vForm, frmCardSet)
            If Not (vCardSet.ContactInfo Is Nothing) AndAlso vCardSet.ContactInfo.ContactGroup = vContactInfo.ContactGroup Then
              If pNewWindow Then
                vCount += 1
                If vCount = vMaxWindows Then
                  ShowInformationMessage(InformationMessages.ImCannotOpenInNewWindow, DataHelper.ContactAndOrganisationGroups(vContactInfo.ContactGroup).GroupName)
                  Return Nothing
                End If
              Else
                If vCardSet.Enabled = False Then
                  For Each vMaintForm As Form In MainHelper.Forms
                    If TypeOf (vMaintForm) Is frmCardMaintenance AndAlso DirectCast(vMaintForm, frmCardMaintenance).MaintenanceParentForm Is vCardSet Then
                      vMaintForm.Close()
                      If vCardSet.Enabled = False Then
                        vMaintForm.BringToFront()
                        Beep()
                        Return Nothing
                      End If
                    End If
                  Next
                End If
                vCardSet.Init(vDataSet, pType, vContactInfo, pRetainPage)
                vCardSet.BringToFront()
                vFound = True
                Exit For
              End If
            End If
          End If
        Next
        If Not vFound Then
          vCardSet = New frmCardSet
          Try
            vCardSet.SuspendLayout()
            vCardSet.MdiParent = MDIForm
            vCardSet.Init(vDataSet, pType, vContactInfo)
            InitWindowForViewType(vCardSet)
          Finally
            vCardSet.ResumeLayout()
          End Try
          vCardSet.Show()
          DirectCast(vCardSet, IPanelVisibility).PanelHasFocus = True
          vCardSet.BringToFront()
          vCardSet.Focus()
        End If
        ' BR12216 - need to see if we are showing the sticky notes
        vContactInfo.ShowLastStickyNote()
        vContactInfo.ShowAlerts()
      End If

    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enSpecifiedDataNotFound Then
        ShowInformationMessage(InformationMessages.ImCannotFindContact)
      Else
        Throw vException
      End If
    Finally
      vCursor.Dispose()
    End Try
    Return vCardSet
  End Function

  Public Shared Function CheckAccessRights(ByVal pContactInfo As ContactInfo) As Boolean
    If pContactInfo.OwnershipAccessLevel <= ContactInfo.OwnershipAccessLevels.oalBrowse Then
      ShowWarningMessage(pContactInfo.ViewAccessMessage)
    Else
      If pContactInfo.OwnershipAccessLevel = ContactInfo.OwnershipAccessLevels.oalRead Then
        Dim vMessage As String = pContactInfo.ReadAccessMessage
        If vMessage.Length > 0 Then ShowInformationMessage(vMessage)
      End If
      Return True
    End If
  End Function

  Public Shared Sub ShowCardDisplay(ByVal pType As CareServices.XMLContactDataSelectionTypes, ByVal pContactNumber As Integer)
    'Figure out what contact group the contact is for 
    'Find the appropriate card index if open and repopulate
    'else create the new card index and populate it
    Dim vForm As Form
    Dim vFound As Boolean
    Dim vCardSet As frmCardSet

    Dim vCursor As New BusyCursor
    Try
      Dim vDataSet As DataSet = Nothing
      Dim vContactInfo As ContactInfo = Nothing
      Dim vShowCard As Boolean
      If pType = CareServices.XMLContactDataSelectionTypes.xcdtContactJournals And pContactNumber = 0 Then
        vContactInfo = New ContactInfo(ContactInfo.ContactTypes.ctContact, EntityGroup.DefaultContactGroupCode)
        vShowCard = True
      Else
        vDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, pContactNumber)
        vContactInfo = New ContactInfo(vDataSet.Tables("DataRow").Rows(0))
        vShowCard = CheckAccessRights(vContactInfo)
      End If
      If vShowCard Then
        For Each vForm In MainHelper.Forms
          If TypeName(vForm) = GetType(frmCardDisplay).Name Then
            vCardSet = DirectCast(vForm, frmCardDisplay)
            If vCardSet.ContactInfo.ContactGroup = vContactInfo.ContactGroup Then
              vCardSet.Init(vDataSet, pType, vContactInfo)
              vCardSet.BringToFront()
              vFound = True
            End If
          End If
        Next

        If Not vFound Then
          vCardSet = New frmCardDisplay
          With vCardSet
            .MdiParent = MDIForm
            .Init(vDataSet, pType, vContactInfo)
            FormHelper.InitWindowForViewType(vCardSet)
            .Show()
            .BringToFront()
          End With
        End If
      End If
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Public Shared Sub CreateNewSelectionSet(pContactNumbers As String)
    Dim vBusyCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      vList("SelectionSetDesc") = String.Format("Drillout Selection {0}", Date.Now.ToString(AppValues.DateTimeFormat))
      Dim vResult As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctSelectionSet, vList)
      Dim vSSNo As Integer = vResult.IntegerValue("SelectionSetNumber")
      UserHistory.AddOtherHistoryNode(HistoryEntityTypes.hetSelectionSets, vSSNo, vList("SelectionSetDesc"))
      vResult = DataHelper.AddSelectionSetContacts(vSSNo, pContactNumbers)
      Dim vItemsNotFound As Integer = vResult.IntegerValue("ItemsNotFound")
      ShowSelectionSet(vSSNo, vList("SelectionSetDesc"))
      If vItemsNotFound > 0 Then ShowInformationMessage("{0} Item(s) missing from the selection set. Data could not be found", vItemsNotFound.ToString)
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    Finally
      vBusyCursor.Dispose()
    End Try
  End Sub

  Public Shared Function GetCommunicationsNumber() As Integer
    Dim vMaintenanceForm As frmCardMaintenance
    For Each vForm As Form In MainHelper.Forms
      If TypeName(vForm) = GetType(frmCardMaintenance).Name Then
        vMaintenanceForm = DirectCast(vForm, frmCardMaintenance)
        Return vMaintenanceForm.CommunicationsNumber
      End If
    Next
  End Function

  Public Shared Function ClipboardContainsCampaignData(ByVal pType As CampaignCopyInfo.CampaignCopyTypes, ByVal pCampaignItem As CampaignItem) As Boolean
    If Clipboard.ContainsData(GetType(CampaignCopyInfo).FullName) Then
      Dim vAppealInfo As CampaignCopyInfo = DirectCast(Clipboard.GetData(GetType(CampaignCopyInfo).FullName), CampaignCopyInfo)
      If vAppealInfo.CampaignCopyType = pType Then
        If pCampaignItem Is Nothing Then
          Return True
        Else
          With vAppealInfo
            If .Campaign = pCampaignItem.Campaign AndAlso .Appeal = pCampaignItem.Appeal _
            AndAlso .Segment = pCampaignItem.Segment AndAlso .CollectionNumber = pCampaignItem.CollectionNumber Then
              Return True
            Else
              Return False
            End If
          End With
        End If
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Shared Function ClipboardContainsCampaignData(ByVal pType As CampaignCopyInfo.CampaignCopyTypes) As Boolean
    Return ClipboardContainsCampaignData(pType, Nothing)
  End Function

  Public Shared Function CreateNewTraderBatch(ByVal pTraderApplication As TraderApplication, Optional ByVal pOwner As Form = Nothing) As Integer
    Dim vForm As frmCardMaintenance
    Dim vResult As Integer

    Dim vCursor As New BusyCursor
    Try
      Dim vList As New ParameterList(True)
      With pTraderApplication
        vList.IntegerValue("TraderApplication") = .ApplicationNumber
        vList("BatchType") = .BatchTypeCode
        If .BatchTypeCode = "CA" Then vList("Provisional") = IIf(.ProvisionalCashBatch = True, "Y", "N").ToString
        If .SourceCode.Length > 0 Then vList("Source") = .SourceCode
        If .ProductCode.Length > 0 Then vList("Product") = .ProductCode
        If .RateCode.Length > 0 Then vList("Rate") = .RateCode
        If .BatchCategory.Length > 0 Then vList("BatchCategory") = .BatchCategory
        If .BatchAnalysisCode.Length > 0 Then vList("BatchAnalysisCode") = .BatchAnalysisCode
        If Not String.IsNullOrEmpty(.BatchPaymentMethod) Then
          vList("BatchPaymentMethod") = .BatchPaymentMethod
        End If
      End With
      vForm = New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctBatches, vList)
      If pOwner Is Nothing Then
        vForm.Show()
      Else
        If vForm.ShowDialog(pOwner) = System.Windows.Forms.DialogResult.OK Then
          vResult = vForm.ReturnList.IntegerValue("BatchNumber")
        End If
      End If
    Finally
      vCursor.Dispose()
    End Try
    Return vResult
  End Function

  Public Shared Function GetIndividualsFromJointContact(ByVal pJointContactNumber As Integer) As DataTable
    Dim vDataSet As DataSet
    Dim vList As New ParameterList(True, True)
    Dim vRelationships As String

    vRelationships = AppValues.ControlValue(AppValues.ControlValues.real_to_joint_relationship)
    If vRelationships.Length > 0 Then vRelationships = "'" & vRelationships & "'"
    If AppValues.ControlValue(AppValues.ControlValues.derived_to_joint_relationship).Length > 0 Then
      vRelationships &= ",'" & AppValues.ControlValue(AppValues.ControlValues.derived_to_joint_relationship) & "'"
    End If
    vList("Relationships") = vRelationships
    vDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, pJointContactNumber, vList)
    Return DataHelper.GetTableFromDataSet(vDataSet)

  End Function

  Public Shared Function IsTraderLoaded(ByRef pTraderCaption As String) As Boolean
    Dim vForm As Form = Nothing
    Dim vFound As Boolean

    For Each vForm In MainHelper.Forms
      If TypeOf (vForm) Is frmTrader Then vFound = True
      If vFound Then Exit For
    Next
    If vFound Then pTraderCaption = vForm.Text
    Return vFound
  End Function

  Public Shared Sub RunTraderApplication(ByVal pTraderApplication As TraderApplication)
    RunTraderApplication(pTraderApplication, Nothing, Nothing, BatchInfo.AdjustmentTypes.None)
  End Sub
  Public Shared Sub RunTraderApplication(ByVal pTraderApplication As TraderApplication, ByVal pList As ParameterList)
    RunTraderApplication(pTraderApplication, pList, Nothing, BatchInfo.AdjustmentTypes.None)
  End Sub
  Public Shared Sub RunTraderApplication(ByVal pTraderApplication As TraderApplication, ByVal pList As ParameterList, ByVal pFrmTraderTransactions As frmTraderTransactions)
    RunTraderApplication(pTraderApplication, pList, pFrmTraderTransactions, BatchInfo.AdjustmentTypes.None)
  End Sub
  Public Shared Sub RunTraderApplication(ByVal pTraderApplication As TraderApplication, ByVal pAdjustmentType As BatchInfo.AdjustmentTypes)
    RunTraderApplication(pTraderApplication, Nothing, Nothing, pAdjustmentType)
  End Sub
  Public Shared Sub RunTraderApplication(ByVal pTraderApplication As TraderApplication, ByVal pList As ParameterList, ByVal pFrmTraderTransactions As frmTraderTransactions, ByVal pAdjustmentType As BatchInfo.AdjustmentTypes)
    If pList IsNot Nothing Then
      pTraderApplication.SetPopupMenuDetails(pList)
    End If
    pTraderApplication.FinancialAdjustment = pAdjustmentType
    Dim vFrmTrader As frmTrader
    If pFrmTraderTransactions Is Nothing Then
      vFrmTrader = New frmTrader(pTraderApplication)
    Else
      vFrmTrader = New frmTrader(pTraderApplication, pFrmTraderTransactions)
    End If
    Dim vShowDialog As Boolean = False

    If pAdjustmentType = BatchInfo.AdjustmentTypes.Adjustment Then
      vShowDialog = True
    ElseIf pList IsNot Nothing Then
      If pList.ContainsKey("EditFromBatchDetails") Then
        vShowDialog = BooleanValue(pList("EditFromBatchDetails"))
        pList.Remove("EditFromBatchDetails")
      End If
    End If

    If vShowDialog = True OrElse FormView = FormViews.Modern Then
      vFrmTrader.Parent = Nothing
      vFrmTrader.TopLevel = True
      If FormView = FormViews.Modern Then vFrmTrader.Owner = MainHelper.MainForm
    End If

    'Checking if the form is disposed 
    'Credit List Recon closes the form if there are no unreconciled transactions 
    If Not vFrmTrader.IsDisposed Then
      If vShowDialog Then
        vFrmTrader.ShowDialog()
      Else
        If Not vFrmTrader.IsDisposed Then vFrmTrader.Show()
      End If
    End If
  End Sub
  Public Shared Sub RunFinancialAdjustments(ByVal pAdjustmentType As CareServices.AdjustmentTypes, ByVal pList As ParameterList, ByVal pTransactionDate As String, ByVal pTransactionSign As String, ByVal pStock As Boolean)
    RunFinancialAdjustments(pAdjustmentType, pList, pTransactionDate, pTransactionSign, pStock, Nothing)
  End Sub
  Public Shared Sub RunFinancialAdjustments(ByVal pAdjustmentType As CareServices.AdjustmentTypes, ByVal pList As ParameterList, ByVal pTransactionDate As String, ByVal pTransactionSign As String, ByVal pStock As Boolean, ByVal pSelectedTrans As ArrayListEx)
    '
    Try
      Dim vAdjustDate As String = AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_adjust_transaction_date)
      If vAdjustDate.Length > 0 Then
        Select Case vAdjustDate
          Case "original", "today"
            'OK
          Case Else
            Throw New CareException(CareException.ErrorNumbers.enAdjustTransDateNotSet)
        End Select
      Else
        Throw New CareException(CareException.ErrorNumbers.enAdjustTransDateNotSet)
      End If
      Dim vCanAdjust As Boolean = True
      Dim vHasPaymentPlan As Boolean = False
      Dim vContainsSalesLedgerItems As Boolean = False
      If pList.ContainsKey("ContainsSalesLedgerItems") Then
        vContainsSalesLedgerItems = BooleanValue(pList("ContainsSalesLedgerItems"))
        pList.Remove("ContainsSalesLedgerItems")
      End If
      Dim vCanPartRefund As Boolean = False
      If pList.ContainsKey("CanPartRefund") Then
        vCanPartRefund = BooleanValue(pList("CanPartRefund"))
        pList.Remove("CanPartRefund")
      End If
      Dim vLineTotal As Double = 0
      If IntegerValue(pList.ValueIfSet("LineNumber")) > 0 AndAlso pList.ContainsKey("LineTotal") Then
        vLineTotal = DoubleValue(pList("LineTotal"))
        pList.Remove("LineTotal")
      End If
      Dim vPaymentMethod As String = String.Empty
      If pList.ContainsKey("PaymentMethodCode") Then
        vPaymentMethod = pList("PaymentMethodCode")
        pList.Remove("PaymentMethodCode")
      End If

      Dim vTable As DataTable
      If pSelectedTrans IsNot Nothing Then pList("BatchNumbers") = pSelectedTrans.CSList
      pList.IntegerValue("AdjustmentType") = CInt(pAdjustmentType)
      vTable = DataHelper.GetTableFromDataSet(DataHelper.CheckFinancialAdjustmentAllowed(pList))
      vCanAdjust = BooleanValue(vTable.Rows(0).Item("CanAdjust").ToString)
      vHasPaymentPlan = BooleanValue(vTable.Rows(0).Item("HasPaymentPlan").ToString)
      If vCanAdjust Then
RunFinancialAdjustmentsContinue:
        'Prompt user to confirm the adjustment
        Dim vQuestion As String = String.Empty
        Select Case pAdjustmentType
          Case CareNetServices.AdjustmentTypes.atAdjustment
            vQuestion = If(vContainsSalesLedgerItems = True, QuestionMessages.QmConfirmChangeSLAnalysis, QuestionMessages.QmConfirmChangeAnalysis)
          Case CareNetServices.AdjustmentTypes.atMove
            vQuestion = QuestionMessages.QmConfirmChangePayer
          Case CareNetServices.AdjustmentTypes.atRefund
            If String.IsNullOrWhiteSpace(vPaymentMethod) = False AndAlso vPaymentMethod.Equals(AppValues.ControlValue(AppValues.ControlTables.credit_sales_controls, AppValues.ControlValues.payment_method)) Then
              If vHasPaymentPlan Then
                vQuestion = QuestionMessages.QmConfirmRefundSLPPTransaction   'BR21357 An Invoice that has no allocations and is used to pay a payment plan. Refund just removes the invoice from the payment plan.
              Else
                vQuestion = QuestionMessages.QmConfirmRefundSLInvoice
              End If
            ElseIf vContainsSalesLedgerItems Then
              vQuestion = QuestionMessages.QmConfirmRefundSLTransaction
            Else
              vQuestion = QuestionMessages.QmConfirmRefundTransaction
            End If
          Case CareNetServices.AdjustmentTypes.atReverse
            vQuestion = If(vContainsSalesLedgerItems = True, QuestionMessages.QmConfirmReverseSLTransaction, QuestionMessages.QmConfirmReverseTransaction)
        End Select

        If IntegerValue(pList.ValueIfSet("LineNumber")) > 0 _
        AndAlso (pAdjustmentType = CareNetServices.AdjustmentTypes.atReverse OrElse pAdjustmentType = CareNetServices.AdjustmentTypes.atRefund) Then
          'Line-level adjustment
          If pAdjustmentType = CareNetServices.AdjustmentTypes.atRefund Then
            vQuestion = If(vContainsSalesLedgerItems = True, QuestionMessages.QmConfirmRefundSLLine, QuestionMessages.QmConfirmRefundLine)
          Else
            vQuestion = If(vContainsSalesLedgerItems = True, QuestionMessages.QmConfirmReverseSLLine, QuestionMessages.QmConfirmReverseLine)
          End If
          If vCanPartRefund Then
            Dim vRL As New ParameterList
            vRL("RunType") = "F"
            vRL("BatchNumber") = pList("BatchNumber")
            vRL = ShowApplicationParameters(CareServices.FunctionParameterTypes.fptFAReverseRefundOptions, vRL)
            If vRL.Count = 0 Then
              vCanAdjust = False
            ElseIf vRL("RunType") = "P" Then
              pAdjustmentType = CareServices.AdjustmentTypes.atPartRefund
              vQuestion = If(vContainsSalesLedgerItems, QuestionMessages.QmConfirmRefundPartSLLine, QuestionMessages.QmConfirmRefundPartLine)
            End If
          End If
        End If

        If Not (pList.ContainsKey("PromptUser") = False OrElse BooleanValue(pList("PromptUser")) = True) Then
          'If user was already asked (via an error from the server) do not need to ask them again
          vQuestion = String.Empty
          pList.Remove("PromptUser")
        End If

        If vCanAdjust = True AndAlso Not (String.IsNullOrWhiteSpace(vQuestion)) Then
          'Default button for this question is 'No'
          vCanAdjust = (ShowQuestion(vQuestion, MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2) = DialogResult.Yes)
        End If
      End If

      If vCanAdjust Then
        Select Case pAdjustmentType
          Case CareServices.AdjustmentTypes.atReverse, CareServices.AdjustmentTypes.atMove, CareServices.AdjustmentTypes.atAdjustment,
               CareServices.AdjustmentTypes.atRefund, CareServices.AdjustmentTypes.atNone, CareServices.AdjustmentTypes.atEventAdjustment,
               CareServices.AdjustmentTypes.atPartRefund

            'atNone is used for Reverse/Refund InAdvance only.
            If pAdjustmentType = CareServices.AdjustmentTypes.atNone AndAlso pList.Contains("PaymentPlanNumber") = False Then Exit Sub

            Dim vDefaults As New ParameterList
            vDefaults("TransactionDate") = pTransactionDate
            If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fp_adjust_transaction_date) = "today" Then vDefaults("TransactionDate") = AppValues.TodaysDate
            If (pAdjustmentType = CareServices.AdjustmentTypes.atAdjustment OrElse pAdjustmentType = CareServices.AdjustmentTypes.atEventAdjustment) Then
              vDefaults("TransactionSign") = pTransactionSign.ToUpper
              vDefaults("TransactionType") = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.adjustment_transaction_type)
            Else
              If pTransactionSign.ToUpper = "C" Then
                vDefaults("TransactionSign") = "D"
              Else
                vDefaults("TransactionSign") = "C"
              End If
              vDefaults("TransactionType") = AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.reversal_transaction_type)
            End If
            vDefaults("AdjustStockLevels") = CBoolYN(pStock AndAlso (pAdjustmentType = CareServices.AdjustmentTypes.atReverse OrElse
                                            pAdjustmentType = CareServices.AdjustmentTypes.atRefund OrElse pAdjustmentType = CareServices.AdjustmentTypes.atPartRefund))
            vDefaults("PostToCashBook") = "Y"
            vDefaults("AdjustOriginalProductCost") = IIf(pAdjustmentType = CareServices.AdjustmentTypes.atMove, "N", AppValues.ControlValue(AppValues.ControlTables.financial_controls, AppValues.ControlValues.adjust_original_product_cost)).ToString

            Dim vFrmApplicationParameters As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptFinancialAdjustment, vDefaults, Nothing)
            If vFrmApplicationParameters.ShowDialog = System.Windows.Forms.DialogResult.OK Then
              Dim vReturnList As ParameterList = vFrmApplicationParameters.ReturnList
              vReturnList("BatchNumber") = pList("BatchNumber")
              vReturnList("TransactionNumber") = pList("TransactionNumber")
              If pList.ContainsKey("LineNumber") Then vReturnList("LineNumber") = pList("LineNumber")
              If pList.ContainsKey("PaymentPlanNumber") Then vReturnList("PaymentPlanNumber") = pList("PaymentPlanNumber")
              vReturnList("AllocationsChecked") = "Y"
              vReturnList.IntegerValue("AdjustmentType") = pAdjustmentType
              Dim vDT As DataTable = Nothing
              Select Case pAdjustmentType
                Case CareServices.AdjustmentTypes.atReverse, CareServices.AdjustmentTypes.atRefund, CareServices.AdjustmentTypes.atPartRefund,
                     CareServices.AdjustmentTypes.atNone
                  'Part Refund/Reverse parameters
                  If pAdjustmentType = CareServices.AdjustmentTypes.atPartRefund Then
                    Dim vRL As ParameterList = Nothing
                    If vContainsSalesLedgerItems Then
                      Dim vSLList As New ParameterList(True, True)
                      vSLList.IntegerValue("BatchNumber") = pList.IntegerValue("BatchNumber")
                      vSLList.IntegerValue("TransactionNumber") = pList.IntegerValue("TransactionNumber")
                      vSLList.IntegerValue("LineNumber") = pList.IntegerValue("LineNumber")
                      Dim vSLDT As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtSalesLedgerAnalysis, vSLList))
                      If vSLDT IsNot Nothing AndAlso vSLDT.Rows.Count > 0 Then
                        Dim vSLRow As DataRow = vSLDT.Rows(0)
                        vSLList("InvoiceAmount") = vSLRow.Item("InvoicePaymentAmount").ToString
                        vSLList("UnallocatedAmount") = vSLRow.Item("UnallocatedAmount").ToString
                        vSLList("LineTotal") = vLineTotal.ToString("0.00")
                        vSLList("RefundAmount") = "0.00"
                      End If
                      vRL = ShowApplicationParameters(CareServices.FunctionParameterTypes.fptFASLPartRefund, vSLList)
                      If vRL Is Nothing OrElse vRL.Count = 0 Then Exit Sub
                      vReturnList("SalesLedgerPartRefund") = "Y"
                      vReturnList("InvoicePaymentAmount") = vRL("InvoiceAmount").ToString
                      vReturnList("UnallocatedAmount") = vRL("UnallocatedAmount").ToString
                      vReturnList("RefundAmount") = vRL("RefundAmount")
                    Else
                      vRL = ShowApplicationParameters(CareServices.FunctionParameterTypes.fptFAPartRefund, pList)
                      If vRL Is Nothing OrElse vRL.Count = 0 Then Exit Sub
                      vReturnList("Quantity") = vRL("Quantity")
                      vReturnList("Issued") = vRL("Issued")
                    End If
                  End If
                  If pAdjustmentType = CareNetServices.AdjustmentTypes.atRefund OrElse pAdjustmentType = CareNetServices.AdjustmentTypes.atPartRefund Then
                    Dim vProcess As New AsyncProcessHandler(AsyncProcessHandler.AsyncProcessHandlerTypes.ReverseTransaction, vReturnList)
                    vDT = DataHelper.GetTableFromDataSet(vProcess.GetDataSetFromResult)
                  Else
                    vDT = DataHelper.GetTableFromDataSet(DataHelper.ReverseTransaction(vReturnList))
                  End If
                Case CareServices.AdjustmentTypes.atMove
                  vReturnList("SmartClient") = pList("SmartClient")
                  vReturnList.IntegerValue("ContactNumber") = ShowFinder(CareServices.XMLDataFinderTypes.xdftContacts, CurrentMainForm, True)
                  If vReturnList.IntegerValue("ContactNumber") > 0 Then vDT = DataHelper.GetTableFromDataSet(DataHelper.ChangeTransactionPayer(vReturnList))
                Case CareServices.AdjustmentTypes.atAdjustment
                  Dim vTA As New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_fa)), pList.IntegerValue("BatchNumber"), , pList.IntegerValue("TransactionNumber"), CareServices.AdjustmentTypes.atAdjustment)
                  vTA.BatchNumber = pList.IntegerValue("BatchNumber")
                  vTA.TransactionNumber = pList.IntegerValue("TransactionNumber")
                  If vReturnList.ContainsKey("BatchDate") Then vTA.BatchDate = vReturnList("BatchDate")
                  If vReturnList.ContainsKey("PostToCashBook") Then vTA.PostToCashBook = vReturnList("PostToCashBook")
                  vTA.FATransactionType = vReturnList("TransactionType")
                  vTA.TransactionDate = vReturnList("TransactionDate")
                  If vReturnList.ContainsKey("Notes") Then vTA.TransactionNote = vReturnList("Notes")
                  If pSelectedTrans IsNot Nothing Then vTA.BatchNumbers = pList("BatchNumbers")
                  RunTraderApplication(vTA, BatchInfo.AdjustmentTypes.Adjustment)
                Case CareServices.AdjustmentTypes.atEventAdjustment
                  Dim vTA As New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_fa)))
                  vTA.BatchNumber = pList.IntegerValue("BatchNumber")
                  vTA.TransactionNumber = pList.IntegerValue("TransactionNumber")
                  vTA.FATransactionType = vReturnList("TransactionType")
                  If pSelectedTrans IsNot Nothing Then vTA.BatchNumbers = pList("BatchNumbers")
                  RunTraderApplication(vTA, BatchInfo.AdjustmentTypes.EventAdjustment)
              End Select
              If vDT IsNot Nothing Then
                Dim vRow As DataRow = vDT.Rows(0)
                Dim vString As New StringBuilder
                vString.AppendLine(vRow.Item("Message").ToString)
                If vDT.Columns.Contains("NewBatchNumber") Then
                  vString.AppendLine(String.Format(InformationMessages.ImFinancialAdjustmentReference, vRow.Item("NewBatchNumber").ToString, vRow.Item("NewTransactionNumber").ToString))
                End If
                ShowInformationMessage(vString.ToString)
              End If
            End If
        End Select
      End If

    Catch vCareException As CareException
      Select Case vCareException.ErrorNumber
        Case CareException.ErrorNumbers.enAdjustTransDateNotSet, CareException.ErrorNumbers.enCannotAdjustPaymentStatus,
             CareException.ErrorNumbers.enCannotReverseOrderAnalysis, CareException.ErrorNumbers.enCannotAdjustZeroBalancePP,
             CareException.ErrorNumbers.enAdjustmentError, CareException.ErrorNumbers.enOriginalBatchOrTransPurged, CareException.ErrorNumbers.enMultiBatchTypes,
             CareException.ErrorNumbers.enOriginalPaymentPartProcessed, CareException.ErrorNumbers.enCCAuthorisationFailed, CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout,
             CareException.ErrorNumbers.enInvalidPaymentMethod, CareException.ErrorNumbers.enCannotAdjustSLAllocation, CareException.ErrorNumbers.enInvoiceAllocationsRemovalRequired
          ShowWarningMessage(vCareException.Message)
        Case CareException.ErrorNumbers.enInvoiceAllocationError
          If pAdjustmentType = CareServices.AdjustmentTypes.atEventAdjustment Then
            GoTo RunFinancialAdjustmentsContinue
          ElseIf ShowQuestion(vCareException.Message, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            pList("PromptUser") = "N"
            GoTo RunFinancialAdjustmentsContinue
          End If
        Case Else
          DataHelper.HandleException(vCareException)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Shared Function ScheduleTask(ByVal pList As ParameterList) As DialogResult
    Return ScheduleTask(pList, ProcessTaskScheduleType.ptsAskToSchedule)
  End Function
  Public Shared Function ScheduleTask(ByVal pList As ParameterList, ByVal pScheduleType As ProcessTaskScheduleType) As DialogResult
    If AppValues.HasItemAccessRights(AppValues.AccessControlItems.aciScheduleTasks) Then
      Dim vNewList As New ParameterList(True)
      If pScheduleType = ProcessTaskScheduleType.ptsAlwaysSchedule Then
        vNewList("ScheduleOnly") = "Y"
      ElseIf pScheduleType = ProcessTaskScheduleType.ptsAlwaysRun Then
        vNewList("RunOnly") = "Y"
      End If
      If pList.ContainsKey("ShowStatus") AndAlso pList("ShowStatus").Length > 0 Then vNewList("ShowStatus") = pList("ShowStatus")
      If pList.ContainsKey("DBUpgrade") AndAlso pList("DBUpgrade").Length > 0 Then vNewList("DBUpgrade") = pList("DBUpgrade")
      Dim vScheduledList As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptScheduledJobDetails, vNewList)
      If vScheduledList.Count = 0 Then
        Return DialogResult.Cancel
      ElseIf vScheduledList.Contains("Schedule") Then
        If vScheduledList.Contains("JobProcessor") Then pList.Add("JobProcessor", vScheduledList("JobProcessor").ToString)
        If vScheduledList.Contains("NotifyStatus") Then pList.Add("NotifyStatus", vScheduledList("NotifyStatus").ToString)
        pList.Add("DueDate", vScheduledList("DueDate").ToString)
        pList.Add("JobFrequency", vScheduledList("JobFrequency").ToString)
        'pList.Add("UpdateJobParameterDates", vScheduledList("UpdateJobParameterDates").ToString)
        If BooleanValue(vScheduledList("UpdateJobParameterDates").ToString) Then
          pList.Add("UpdateJobParameterDates", "A")
        Else
          pList.Add("UpdateJobParameterDates", "N")
        End If
        Return DialogResult.Yes
      ElseIf vScheduledList.Contains("RunNow") Then
        If vScheduledList.ContainsKey("ShowTaskStatus") Then pList.Add("ShowTaskStatus", vScheduledList("ShowTaskStatus"))
        Return DialogResult.No
      End If
    End If
    Return DialogResult.No
  End Function

  Public Shared Sub DoBulkContactDeletion(ByVal pSetNumber As Integer)
    If ShowQuestion(QuestionMessages.QmDeleteAllContacts, MessageBoxButtons.YesNo) = DialogResult.Yes Then
      Dim vList As New ParameterList(True)
      vList.Add("SelectionSet", pSetNumber)
      ProcessTask(CareServices.TaskJobTypes.tjtBulkContactDeletion, vList, False, ProcessTaskScheduleType.ptsAskToSchedule, True)
    End If
  End Sub

  Public Shared Sub DoContactMerge(ByVal pOrganisations As Boolean)
    Dim vCursor As New BusyCursor
    Dim vAppParameters As frmApplicationParameters
    Dim vReturnList As ParameterList = Nothing

    If pOrganisations Then
      vAppParameters = New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optOrganisationMerge, Nothing, Nothing, "Organisation Merge")
    Else
      vAppParameters = New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optContactMerge, Nothing, Nothing, " Contact Merge")
    End If
    AddHandler vAppParameters.ProcessApplication, AddressOf ProcessContactMerge
    vAppParameters.ShowDialog(CurrentMainForm)
  End Sub

  Public Shared Sub DoAddressMerge()
    Dim vCursor As New BusyCursor
    Dim vAppParameters As frmApplicationParameters
    Dim vReturnList As ParameterList = Nothing

    vAppParameters = New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optAddressMerge, Nothing, Nothing, "Address Merge")
    AddHandler vAppParameters.ProcessApplication, AddressOf ProcessAddressMerge
    vAppParameters.ShowDialog(CurrentMainForm)
  End Sub

  Public Shared Sub DoDuplicateMeeting(ByVal pMeetingNumber As Integer)
    Dim vList As New ParameterList(True)
    vList.IntegerValue("MeetingNumber") = pMeetingNumber
    Dim vDataSet As DataSet = DataHelper.GetTableData(CareNetServices.XMLTableDataSelectionTypes.xtdstMeetings, vList)
    Dim vDefaults As New ParameterList
    Dim vDataTable As DataTable = vDataSet.Tables("DataRow")
    vDefaults("Description") = vDataTable.Rows(0).Item("MeetingDesc").ToString()
    vDefaults("MeetingNumber") = pMeetingNumber.ToString
    'date should be today's date but with time of original meeting
    Dim vDate As String = Date.Now.Date.ToString
    vDate = vDate.Substring(0, 10)
    Dim vTime As String = CDate(vDataTable.Rows(0).Item("MeetingDate")).TimeOfDay.ToString
    Dim vMeetingDateTime As DateTime = CDate(vDate + " " + vTime)
    vDefaults("Date") = vMeetingDateTime.ToString
    Dim vList1 As ParameterList = FormHelper.ShowApplicationParameters(CareServices.FunctionParameterTypes.fptDuplicateMeeting, vDefaults)
  End Sub

  Private Shared Sub ProcessContactMerge(ByVal pList As ParameterList, ByRef pReset As Nullable(Of Boolean))
    If pList IsNot Nothing AndAlso pList.Count > 0 Then
      Dim vRetry As Boolean
      pReset = True
      Dim vOrganisationMerge As Boolean
      Dim vList As New ParameterList(True)
      If pList.ContainsKey("OrganisationNumber") Then
        vList("ContactNumber") = pList("OrganisationNumber")
        vList("DuplicateContactNumber") = pList("OrganisationNumber2")
        vOrganisationMerge = True
      Else
        vList("ContactNumber") = pList("ContactNumber")
        vList("DuplicateContactNumber") = pList("ContactNumber2")
      End If
      If pList.ContainsKey("Notes") Then vList("Notes") = pList("Notes")
      If pList.ContainsKey("Queue") Then vList("Queue") = pList("Queue")
      Do
        Try
          vRetry = False
          DataHelper.MergeContact(vList)
          UserHistory.RemoveContactHistoryNode(vList.IntegerValue("DuplicateContactNumber"))
          For Each vForm As Form In MainHelper.Forms
            If TypeName(vForm) = GetType(frmCardSet).Name Then
              If DirectCast(vForm, frmCardSet).ContactInfo.ContactNumber = vList.IntegerValue("DuplicateContactNumber") Then vForm.Close()
            End If
          Next
          If pList.ContainsKey("Queue") Then
            If vOrganisationMerge Then
              ShowInformationMessage(InformationMessages.ImMergeQueuedOrganisations)
            Else
              ShowInformationMessage(InformationMessages.ImMergeQueuedContacts)
            End If
          Else
            If vOrganisationMerge Then
              ShowInformationMessage(InformationMessages.ImMergeCompleteOrganisations)
            Else
              ShowInformationMessage(InformationMessages.ImMergeCompleteContacts)
            End If
          End If
        Catch vEx As CareException
          Select Case vEx.ErrorNumber
            Case CareException.ErrorNumbers.enMergeError
              ShowInformationMessage(vEx.Message)
              pReset = False
            Case CareException.ErrorNumbers.enMergeJointToJoint
              If ShowQuestion(QuestionMessages.QmConfirmMergeJoint, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmJointMerge") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case CareException.ErrorNumbers.enMergeConfirmCreditCustomers
              If ShowQuestion(QuestionMessages.QmConfirmMergeCreditCustomers, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmMergeCreditCustomers") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case CareException.ErrorNumbers.enMergeQueueConfirm
              Dim vMsg As String
              If vOrganisationMerge Then
                vMsg = QuestionMessages.QmConfirmMergeQueueOrganisations
              Else
                vMsg = QuestionMessages.QmConfirmMergeQueue
              End If
              If ShowQuestion(vMsg, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("Confirm") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case CareException.ErrorNumbers.enMergeConfirm
              If ShowQuestion(QuestionMessages.QmConfirmMerge, MessageBoxButtons.YesNo, (IIf(vOrganisationMerge = True, "Organisations", "Contacts")).ToString) = DialogResult.Yes Then
                If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_merge_run_audit_report, True) Then
                  'vReportFilename = gvEnv.GetAuditFileName("C_" & mvDuplicateContact & ".txt")
                  Dim vReportList As New ParameterList(True)
                  If vOrganisationMerge Then
                    vReportList("ReportCode") = "ORGAUD"
                  Else
                    vReportList("ReportCode") = "CONAUD"
                  End If
                  vReportList.IntegerValue("RP1") = vList.IntegerValue("DuplicateContactNumber")
                  vReportList.IntegerValue("RPmerge_info") = vList.IntegerValue("ContactNumber")
                  Dim vCancelled As Boolean
                  Dim vPrintHandler As New PrintHandler
                  vPrintHandler.OutputType = PrintHandler.OutputTypes.Audit
                  vPrintHandler.PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave, vCancelled)
                  'Allow the merge even if cancelled the report  If vCancelled Then Exit Do
                End If
                vList("Confirm") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case CareException.ErrorNumbers.enMergeDeleteAddress
              If Not vList.Contains("DeleteAddress") Then 'BR21369
                If ShowQuestion(QuestionMessages.QmConfirmMergeDeleteAddress, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                  vList("DeleteAddress") = "Y"
                Else
                  vList("DeleteAddress") = "N"
                End If
              Else
                If vList("DeleteAddress") = "Y" Then
                  ShowInformationMessage(InformationMessages.ImCannotDeleteDupDefaultAddr)
                  vList("DeleteAddress") = "N"
                End If
              End If
              vRetry = True
            Case CareException.ErrorNumbers.enMergeObjectException
              ShowErrorMessage(vEx.Message)
              pReset = False
            Case Else
              DataHelper.HandleException(vEx)
              pReset = False
          End Select
        End Try
      Loop While vRetry
    End If
  End Sub

  Private Shared Sub ProcessAddressMerge(ByVal pList As ParameterList, ByRef pReset As Nullable(Of Boolean))

    If pList IsNot Nothing AndAlso pList.Count > 0 Then
      Dim vRetry As Boolean
      pReset = True
      Dim vList As New ParameterList(True)
      vList("ContactNumber") = pList("ContactNumber")
      vList("DuplicateContactNumber") = pList("ContactNumber2")
      vList("AddressNumber") = pList("AddressNumber")
      vList("DuplicateAddressNumber") = pList("AddressNumber2")
      If pList.ContainsKey("Notes") Then vList("Notes") = pList("Notes")
      If pList.ContainsKey("Queue") Then vList("Queue") = pList("Queue")
      Do
        Try
          vRetry = False
          DataHelper.MergeAddress(vList)
          UserHistory.RemoveContactHistoryNode(vList.IntegerValue("DuplicateContactNumber"))
          For Each vForm As Form In MainHelper.Forms
            If TypeName(vForm) = GetType(frmCardSet).Name Then
              If DirectCast(vForm, frmCardSet).ContactInfo.ContactNumber = vList.IntegerValue("DuplicateContactNumber") Then vForm.Close()
            End If
          Next
        Catch vEx As CareException
          Select Case vEx.ErrorNumber
            Case CareException.ErrorNumbers.enMergeError
              ShowInformationMessage(vEx.Message)
              pReset = False
            Case CareException.ErrorNumbers.enConfirmUpdateAddress
              If ShowQuestion(QuestionMessages.QmConfirmUpdateAddress, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmUpdateDates") = "Y"
              Else
                vList("ConfirmUpdateDates") = "N"
              End If
              vRetry = True
            Case CareException.ErrorNumbers.enConfirmDeleteDuplicateAddress
              If ShowQuestion(QuestionMessages.QmConfirmDeleteDuplicateAddress, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmDeleteDuplicateAddress") = "Y"
              Else
                vList("ConfirmDeleteDuplicateAddress") = "N"
              End If
              vRetry = True
            Case CareException.ErrorNumbers.enConfirmDuplicateAddressHistoric
              If ShowQuestion(QuestionMessages.QmConfirmDuplicateAddressHistoric, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmDuplicateAddressHistoric") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case CareException.ErrorNumbers.enConfirmOneOrBothAddress
              If ShowQuestion(QuestionMessages.QmConfirmOneOrBothAddress, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("OneOrBothAddress") = "Y"
                vRetry = True
              Else
                vList("OneOrBothAddress") = "N"
                pReset = False
              End If
            Case CareException.ErrorNumbers.enConfirmQueueAddress
              If ShowQuestion(QuestionMessages.QmConfirmQueueAddress, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                vList("ConfirmQueueAddress") = "Y"
                vRetry = True
              Else
                vList("ConfirmQueueAddress") = "N"
                pReset = False
              End If
            Case Else
              DataHelper.HandleException(vEx)
              pReset = False
          End Select
        End Try
      Loop While vRetry
    End If
  End Sub

  Public Shared Sub DoAmalgamateOrganisation()
    Dim vCursor As New BusyCursor
    Dim vAppParameters As New frmApplicationParameters(EditPanelInfo.OtherPanelTypes.optAmalgamateOrganisation, Nothing, Nothing, "Amalgamate Organisations")
    AddHandler vAppParameters.ProcessApplication, AddressOf ProcessAmalgamateOrganisaion
    vAppParameters.ShowDialog(MDIForm)
  End Sub

  Private Shared Sub ProcessAmalgamateOrganisaion(ByVal pList As ParameterList, ByRef pReset As Nullable(Of Boolean))
    If pList IsNot Nothing AndAlso pList.Count > 0 Then
      Dim vRetry As Boolean
      pReset = True
      Dim vList As New ParameterList(True)
      vList("OrganisationNumber") = pList("OrganisationNumber")
      vList("AmalgamateOrganisationNumber") = pList("OrganisationNumber2")
      If pList.ContainsKey("Status2") Then vList("Status") = pList("Status2")
      If pList.ContainsKey("OwnershipGroup2") Then vList("OwnershipGroup") = pList("OwnershipGroup2")
      If pList.ContainsKey("Notes") Then vList("Notes") = pList("Notes")
      'If pList.ContainsKey("Queue") Then vList("Queue") = pList("Queue")
      Do
        Try
          vRetry = False
          DataHelper.AmalgamateOrganisation(vList)
          UserHistory.RemoveContactHistoryNode(vList.IntegerValue("AmalgamateOrganisationNumber"))
          For Each vForm As Form In MainHelper.Forms
            If TypeName(vForm) = GetType(frmCardSet).Name Then
              If DirectCast(vForm, frmCardSet).ContactInfo.ContactNumber = vList.IntegerValue("AmalgamateOrganisationNumber") Then vForm.Close()
            End If
          Next
          'If pList.ContainsKey("Queue") Then
          '
        Catch vEx As CareException
          Select Case vEx.ErrorNumber
            Case CareException.ErrorNumbers.enMergeError
              ShowInformationMessage(vEx.Message)
              pReset = False
              'Case CareException.ErrorNumbers.enMergeConfirmCreditCustomers
              '  If ShowQuestion(QuestionMessages.QmConfirmMergeCreditCustomers, MessageBoxButtons.YesNo) = DialogResult.Yes Then
              '    vList("ConfirmMergeCreditCustomers") = "Y"
              '    vRetry = True
              '  Else
              '    pCancel = True
              '  End If
              'Case CareException.ErrorNumbers.enMergeQueueConfirm
              '  If ShowQuestion(QuestionMessages.QmConfirmMergeQueue, MessageBoxButtons.YesNo) = DialogResult.Yes Then
              '    vList("Confirm") = "Y"
              '    vRetry = True
              '  Else
              '    pCancel = True
              '  End If
            Case CareException.ErrorNumbers.enAmalgamationConfirm
              If ShowQuestion(QuestionMessages.QmConfirmAmalgamate, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_merge_run_audit_report, True) Then
                  'vReportFilename = gvEnv.GetAuditFileName("C_" & mvDuplicateContact & ".txt")
                  Dim vReportList As New ParameterList(True)
                  vReportList("ReportCode") = "ORGAUD"
                  vReportList.IntegerValue("RP1") = vList.IntegerValue("AmalgamateOrganisationNumber")
                  vReportList.IntegerValue("RPmerge_info") = vList.IntegerValue("OrganisationNumber")
                  Dim vCancelled As Boolean
                  Call (New PrintHandler).PrintReport(vReportList, PrintHandler.PrintReportOutputOptions.AllowSave, vCancelled)
                  'Allow the merge even if user cancelled the report  If vCancelled Then Exit Do
                End If
                vList("Confirm") = "Y"
                vRetry = True
              Else
                pReset = False
              End If
            Case Else
              DataHelper.HandleException(vEx)
              pReset = False
          End Select
        End Try
      Loop While vRetry
    End If
  End Sub

  Public Shared Function GetTaskJobTypeName(ByVal pTaskJobType As CareServices.TaskJobTypes) As String
    Dim vName As String
    Select Case pTaskJobType
      Case CareServices.TaskJobTypes.tjtNone
        vName = "TestApp"
      Case CareServices.TaskJobTypes.tjtExpirePaymentPlans
        vName = "ExpirePaymentPlans"
      Case CareServices.TaskJobTypes.tjtMailsortUpdate
        vName = "MailsortUpdate"
      Case CareServices.TaskJobTypes.tjtPriceChange
        vName = "PriceChange"
      Case CareServices.TaskJobTypes.tjtRemoveArrears
        vName = "RemoveArrears"
      Case CareServices.TaskJobTypes.tjtDirectDebitRun
        vName = "DirectDebitRun"
      Case CareServices.TaskJobTypes.tjtDDClaimFile
        vName = "DirectDebitClaimFile"
      Case CareServices.TaskJobTypes.tjtMembershipSuspension
        vName = "MembershipSuspension"
      Case CareServices.TaskJobTypes.tjtAutoSOReconciliation
        vName = "AutoSOReconciliation"
      Case CareServices.TaskJobTypes.tjtStatementLoader
        vName = "StatementLoader"
      Case CareServices.TaskJobTypes.tjtDDMandateFile
        vName = "DirectDebitMandateFile"
      Case CareServices.TaskJobTypes.tjtDDCreditFile
        vName = "DirectDebitCreditFile"
      Case CareServices.TaskJobTypes.tjtCCClaimFile
        vName = "CreditCardAuthorityClaim"
      Case CareServices.TaskJobTypes.tjtCCClaimReport
        vName = "CreditCardAuthorityReport"
      Case CareServices.TaskJobTypes.tjtCardSalesFile
        vName = "CardSalesFile"
      Case CareServices.TaskJobTypes.tjtCardSalesReport
        vName = "CardSalesReport"
      Case CareServices.TaskJobTypes.tjtCreditCardRun
        vName = "CreditCardRun"
      Case CareServices.TaskJobTypes.tjtThankYouLetters
        vName = "ThankYouLetterProduction"
      Case CareServices.TaskJobTypes.tjtRenewalsAndReminders
        vName = "RenewalsAndReminders"
      Case CareServices.TaskJobTypes.tjtPayingInSlips
        vName = "PayingInSlipProduction"
      Case CareServices.TaskJobTypes.tjtCashBookPosting
        vName = "CashBookPosting"
      Case CareServices.TaskJobTypes.tjtBatchUpdate
        vName = "BatchUpdate"
      Case CareServices.TaskJobTypes.tjtGiftAidClaim
        vName = "GiftAidClaim"
      Case CareServices.TaskJobTypes.tjtMailingCount
        vName = "MailingCount"
      Case CareServices.TaskJobTypes.tjtMailingRun
        vName = "MailingRun"
      Case CareServices.TaskJobTypes.tjtGenerateMarketingData
        vName = "GenerateMarketingData"
      Case CareServices.TaskJobTypes.tjtGenerateAddressGeoRegions
        vName = "UpdateRegionalData"
      Case CareServices.TaskJobTypes.tjtContactDeDuplication
        vName = "ContactDeDuplication"
      Case CareServices.TaskJobTypes.tjtBulkMerge
        vName = "BulkMerge"
      Case CareServices.TaskJobTypes.tjtFutureMembershipChanges
        vName = "FutureMembershipChanges"
      Case CareServices.TaskJobTypes.tjtCreateJournalFiles
        vName = "CreateJournalFiles"
      Case CareServices.TaskJobTypes.tjtPickingList
        vName = "PickingListProduction"
      Case CareServices.TaskJobTypes.tjtConfirmStockAllocation
        vName = "ConfirmStockAllocation"
      Case CareServices.TaskJobTypes.tjtBackOrderAllocation
        vName = "BackOrderAllocation"
      Case CareServices.TaskJobTypes.tjtDespatchNotes
        vName = "DespatchNotes"
      Case CareServices.TaskJobTypes.tjtInvoiceTransfer
        vName = "InvoiceTransfer"
      Case CareServices.TaskJobTypes.tjtBatchPurge
        vName = "BatchPurge"
      Case CareServices.TaskJobTypes.tjtBackOrderPurge
        vName = "BackOrderPurge"
      Case CareServices.TaskJobTypes.tjtPickingAndDespatchPurge
        vName = "PickingAndDespatchPurge"
      Case CareServices.TaskJobTypes.tjtGiftAidPotentialClaim
        vName = "GiftAidPotentialClaim"
      Case CareServices.TaskJobTypes.tjtDataImport
        vName = "DataImport"
      Case CareServices.TaskJobTypes.tjtManualSOReconciliation
        vName = "ManualSOReconciliation"
      Case CareServices.TaskJobTypes.tjtAmendmentHistoryView
        vName = "AmendmentHistoryView"
      Case CareServices.TaskJobTypes.tjtSetPostDatedContacts
        vName = "SetPostDatedContacts"
      Case CareServices.TaskJobTypes.tjtGAYEPaymentLoader
        vName = "PayrollGivingPaymentLoader"
      Case CareServices.TaskJobTypes.tjtGAYEReconciliation
        vName = "PayrollGivingReconciliation"
      Case CareServices.TaskJobTypes.tjtBankDataLoad
        vName = "LoadBankData"
      Case CareServices.TaskJobTypes.tjtCAFProvisionalBatchClaim
        vName = "CAFVoucherClaimReport"
      Case CareServices.TaskJobTypes.tjtCAFCardSalesReport
        vName = "CAFCardSalesReport"
      Case CareServices.TaskJobTypes.tjtCAFExpectedPaymentsReport
        vName = "CAFExpectedPaymentsReport"
      Case CareServices.TaskJobTypes.tjtCAFPaymentLoader
        vName = "CAFPaymentLoader"
      Case CareServices.TaskJobTypes.tjtMailingDocumentProduction
        vName = "MailingDocumentProduction"
      Case CareServices.TaskJobTypes.tjtBankTransactionsReport
        vName = "BankTransactionsReport"
      Case CareServices.TaskJobTypes.tjtPurchasedProductReport
        vName = "PurchasedProductReport"
      Case CareServices.TaskJobTypes.tjtBranchDonationsReport
        vName = "BranchDonationsReport"
      Case CareServices.TaskJobTypes.tjtJuniorMembershipAnalysisReport
        vName = "JuniorMemberAnalysisReport"
      Case CareServices.TaskJobTypes.tjtOutstandingBatchesReport
        vName = "OutstandingBatchesReport"
      Case CareServices.TaskJobTypes.tjtCAFPaymentReconciliation
        vName = "CAFPaymentReconciliation"
      Case CareServices.TaskJobTypes.tjtBranchIncomeReport
        vName = "BranchIncomeReport"
      Case CareServices.TaskJobTypes.tjtConvertManualDirectDebits
        vName = "ConvertManualDirectDebits"
      Case CareServices.TaskJobTypes.tjtBACSRejections
        vName = "BACSMessaging"
      Case CareServices.TaskJobTypes.tjtBallotPaperProduction
        vName = "BallotPaperProduction"
      Case CareServices.TaskJobTypes.tjtAssumedVotingRights
        vName = "AssumedVotingRightsReport"
      Case CareServices.TaskJobTypes.tjtPeriodStatsGenerateData
        vName = "GeneratePeriodStatistics"
      Case CareServices.TaskJobTypes.tjtPeriodStatsReport
        vName = "PeriodStatisticsReport"
      Case CareServices.TaskJobTypes.tjtSelectionTester
        vName = "SelectionTester"
      Case CareServices.TaskJobTypes.tjtUpdateActionStatus
        vName = "UpdateActionStatus"
      Case CareServices.TaskJobTypes.tjtPostPayments
        vName = "TransferPayments"
      Case CareServices.TaskJobTypes.tjtListAllContacts
        vName = "ListAllContacts"
      Case CareServices.TaskJobTypes.tjtPurgePrizeDrawBatches
        vName = "PurgePrizeDrawBatches"
      Case CareServices.TaskJobTypes.tjtPayrollPledgeCancellation
        vName = "GAYEPledgesBulkCancellation"
      Case CareServices.TaskJobTypes.tjtStandingOrderCancellation
        vName = "StandingOrderCancellation"
      Case CareServices.TaskJobTypes.tjtDirectDebitMailing
        vName = "DirectDebitMailing"
      Case CareServices.TaskJobTypes.tjtStandingOrderMailing
        vName = "StandingOrderMailing"
      Case CareServices.TaskJobTypes.tjtSubscriptionMailing
        vName = "SubscriptionMailing"
      Case CareServices.TaskJobTypes.tjtMemberMailing
        vName = "MemberMailing"
      Case CareServices.TaskJobTypes.tjtMembCardMailing
        vName = "MembershipCardMailing"
      Case CareServices.TaskJobTypes.tjtPayerMailing
        vName = "PayerMailing"
      Case CareServices.TaskJobTypes.tjtSelectionManagerMailing
        vName = "SelectionManagerMailing"
      Case CareServices.TaskJobTypes.tjtCustomerTransfer
        vName = "CustomerTransfer"
      Case CareServices.TaskJobTypes.tjtUpdateSearchNames
        vName = "UpdateSearchNames"
      Case CareServices.TaskJobTypes.tjtStockExport
        vName = "StockMovementExport"
      Case CareServices.TaskJobTypes.tjtEventTotalsUpdate
        vName = "EventTotalsUpdate"
      Case CareServices.TaskJobTypes.tjtGADConfirmation
        vName = "GADConfirmation"
      Case CareServices.TaskJobTypes.tjtGASPotentialClaim
        vName = "GASponsorshipPotentialClaim"
      Case CareServices.TaskJobTypes.tjtGASTaxClaim
        vName = "GASponsoredEventTaxClaim"
      Case CareServices.TaskJobTypes.tjtPOTransferSuppliers
        vName = "TransferSuppliers"
      Case CareServices.TaskJobTypes.tjtUpdatePaymentSchedule
        vName = "UpdatePaymentSchedule"
      Case CareServices.TaskJobTypes.tjtUpdateGovernmentRegions
        vName = "UpdateGovernmentRegions"
      Case CareServices.TaskJobTypes.tjtPayrollPledgeMailing
        vName = "PayrollGivingPledgeMailing"
      Case CareServices.TaskJobTypes.tjtCreditCardAuthorisationReport
        vName = "CCAuthorisationReport"
      Case CareServices.TaskJobTypes.tjtBulkAddressMerge
        vName = "BulkAddressMerge"
      Case CareServices.TaskJobTypes.tjtUpdatePaymentPlanProducts
        vName = "UpdatePaymentPlanProducts"
      Case CareServices.TaskJobTypes.tjtCheckPaymentPlans
        vName = "CheckPaymentPlans"
      Case CareServices.TaskJobTypes.tjtMemberFulfilment
        vName = "MemberFulfilment"
      Case CareServices.TaskJobTypes.tjtDormantContactDeletion
        vName = "DormantContactDeletion"
      Case CareServices.TaskJobTypes.tjtScheduledReport
        vName = "ScheduledReport"
      Case CareServices.TaskJobTypes.tjtPostTaxPGReconciliation
        vName = "PostTaxPGReconciliation"
      Case CareServices.TaskJobTypes.tjtPISStatementLoader
        vName = "PISStatementLoader"
      Case CareServices.TaskJobTypes.tjtPISReconciliation
        vName = "PISReconciliation"
      Case CareServices.TaskJobTypes.tjtPublicCollectionsFulfilment
        vName = "CollectionsFulfilment"
      Case CareServices.TaskJobTypes.tjtEventBookerMailing
        vName = "EventBookerMailing"
      Case CareServices.TaskJobTypes.tjtEventDelegateMailing
        vName = "EventDelegateMailing"
      Case CareServices.TaskJobTypes.tjtEventPersonnelMailing
        vName = "EventPersonnelMailing"
      Case CareServices.TaskJobTypes.tjtEventSponsorMailing
        vName = "EventSponsorMailing"
      Case CareServices.TaskJobTypes.tjtDutchElectronicPaymentsLoader
        vName = "DutchPaymentsLoader"
      Case CareServices.TaskJobTypes.tjtDutchElectronicPaymentsReconciliation
        vName = "DutchPaymentsReconciliation"
      Case CareServices.TaskJobTypes.tjtBulkGiftAidUpdate
        vName = "BulkGiftAidUpdate"
      Case CareServices.TaskJobTypes.tjtIrishGiftAidMailing
        vName = "IrishGiftAidMailing"
      Case CareServices.TaskJobTypes.tjtIrishGiftAidPotentialClaim
        vName = "IrishGiftAidPotentialClaim"
      Case CareServices.TaskJobTypes.tjtIrishGiftAidTaxClaim
        vName = "IrishGiftAidTaxClaim"
      Case CareServices.TaskJobTypes.tjtBulkContactDeletion
        vName = "BulkContactDeletion"
      Case CareServices.TaskJobTypes.tjtProcessAddressChanges
        vName = "ProcessAddressChanges"
      Case CareServices.TaskJobTypes.tjtBulkOrganisationMerge
        vName = "BulkOrganisationMerge"
      Case CareServices.TaskJobTypes.tjtAllocateDonationToBox
        vName = "AllocateDonations"
      Case CareServices.TaskJobTypes.tjtPrintBoxLabels
        vName = "CreateUnallocatedBoxes"
      Case CareServices.TaskJobTypes.tjtShipDistributionBoxes
        vName = "SetShippingInformation"
      Case CareServices.TaskJobTypes.tjtSetBoxesArrived
        vName = "SetArrivalInformation"
      Case CareServices.TaskJobTypes.tjtChequeProduction
        vName = "ChequeProduction"
      Case CareServices.TaskJobTypes.tjtGenerateRollOfHonour
        vName = "GenerateRollOfHonour"
      Case CareServices.TaskJobTypes.tjtCheetahMailMetaData
        vName = "CheetahMailMetaData"
      Case CareServices.TaskJobTypes.tjtCheetahMailEventData
        vName = "CheetahMailEventData"
      Case CareServices.TaskJobTypes.tjtCheetahMailTotals
        vName = "CheetahMailTotals"
      Case CareServices.TaskJobTypes.tjtBulkMailerStatistics
        vName = "BulkMailerStatistics"
      Case CareServices.TaskJobTypes.tjtEMailProcessor
        vName = "EmailProcessor"
      Case CareServices.TaskJobTypes.tjtDistributionBoxReports
        vName = "DistributionBoxReports"
      Case CType(CareNetServices.TaskJobTypes.tjtEventBlockBooking, CareServices.TaskJobTypes)
        vName = "EventBlockBooking"
      Case CareNetServices.TaskJobTypes.tjtPurchaseOrderGeneration
        vName = "AutoGeneratePurchaseOrder"
      Case CareNetServices.TaskJobTypes.tjtPurchaseOrderPrint
        vName = "PrintPurchaseOrder"
      Case CareNetServices.TaskJobTypes.tjtIssueResources
        vName = "IssueResources"
      Case CareNetServices.TaskJobTypes.tjtCancelEvent
        vName = "CancellationBlockMove"
      Case CareNetServices.TaskJobTypes.tjtDatabaseUpgrade
        vName = "DatabaseUpgrade"
      Case CareNetServices.TaskJobTypes.tjtApplyCPDPoints
        vName = "ApplyPoints"
      Case CareNetServices.TaskJobTypes.tjtRegisterSurvey
        vName = "RegisterSurvey"
      Case CareNetServices.TaskJobTypes.tjtCreditStatementGeneration
        vName = "CreditStatementGeneration"
      Case CareNetServices.TaskJobTypes.tjtCancelProvisionalBookings
        vName = "CancelProvisionalBookings"
      Case CareNetServices.TaskJobTypes.tjtPostcodeValidation
        vName = "PostcodeValidation"
      Case CareNetServices.TaskJobTypes.tjtProcessPurchaseOrderPayments
        vName = "ProcessPurchaseOrderPayments"
      Case CareNetServices.TaskJobTypes.tjtUploadBACSMessagingData
        vName = "UploadBACSMessagingData"
      Case CareNetServices.TaskJobTypes.tjtApplyPaymentPlanSurcharges
        vName = "ApplyPaymentPlanSurcharges"
      Case CareNetServices.TaskJobTypes.tjtRecalculateLoanInterest
        vName = "ReCalculateLoanInterest"
      Case CareNetServices.TaskJobTypes.tjtExamAllocateCandidateNumbers
        vName = "ExamAllocateCandidateNumbers"
      Case CareNetServices.TaskJobTypes.tjtExamAllocateMarkers
        vName = "ExamAllocateMarkers"
      Case CareNetServices.TaskJobTypes.tjtExamApplyGrading
        vName = "ExamApplyGrading"
      Case CareNetServices.TaskJobTypes.tjtExamGenerateExemptionInvoices
        vName = "ExamGenerateExemptionInvoices"
      Case CareNetServices.TaskJobTypes.tjtDataUpdates
        vName = "DataUpdates"
      Case CareNetServices.TaskJobTypes.tjtExamLoadCsvResults
        vName = "ExamLoadCsvResults"
      Case CareNetServices.TaskJobTypes.tjtUpdateLoanInterestRates
        vName = "UpdateLoanInterestRates"
      Case CareNetServices.TaskJobTypes.tjtCheckNonCoreTables
        vName = "CheckNonCoreTables"
      Case CareNetServices.TaskJobTypes.tjtGenerateTableCreationFiles
        vName = "GenerateTableCreationFiles"
      Case CareNetServices.TaskJobTypes.tjtGetConfigNameData
        vName = "GetConfigNameData"
      Case CareNetServices.TaskJobTypes.tjtGetReportData
        vName = "GetReportData"
      Case CareNetServices.TaskJobTypes.tjtBulkUpdateActivity
        vName = "BulkUpdateActivity"
      Case CareNetServices.TaskJobTypes.tjtRegenerateMessageQueue
        vName = "RegenerateMessageQueue"
      Case CareNetServices.TaskJobTypes.tjtTransferPaymentPlanChanges
        vName = "TransferPaymentPlanChanges"
      Case CareNetServices.TaskJobTypes.tjtExamCertificateReprints
        vName = "CertificateReprints"
      Case CareNetServices.TaskJobTypes.tjtUpdatePrincipalUser
        vName = "UpdatePrincipalUser"
      Case CareNetServices.TaskJobTypes.tjtUpdateFutureMembershipType
        vName = "UpdateFutureMembershipType"
      Case CareNetServices.TaskJobTypes.tjtProcessCertificateData
        vName = "ProcessCertificateData"
      Case CareNetServices.TaskJobTypes.tjtPurgeStickyNotes
        vName = "PurgeStickyNotes"
      Case Else
        vName = "UnknownTask"
    End Select
    Return vName
  End Function
  Private Shared Sub ProcessShortfall(ByVal pList As ParameterList)
    'We are confirming stock and have beed notified of a shortfall
    Dim vForm As New frmStockShortfall(pList)
    If vForm.ShowDialog = DialogResult.OK Then
      'Now flag the stock as confirmed and run the task again
      pList("Checkbox") = "Y"
      ProcessTask(CareServices.TaskJobTypes.tjtConfirmStockAllocation, pList, False, ProcessTaskScheduleType.ptsAlwaysRun)
    End If

  End Sub
  Private Shared Function GetReportDestinationType(ByVal pReportDestination As String) As String
    Dim vResult As String
    Select Case pReportDestination
      Case "Print"
        vResult = "PrintXML"
      Case "Preview"
        vResult = "PreviewXML"
      Case "None"
        vResult = "None"
      Case Else
        vResult = "Save"
    End Select
    Return vResult
  End Function

  Friend Shared Sub RunFastDataEntry(ByVal pFDEPageNumber As Integer)
    Dim vFrmFDE As New frmFastDataEntry(pFDEPageNumber, False)
    vFrmFDE.MdiParent = MainHelper.MainForm
    vFrmFDE.Show()
  End Sub

  Friend Shared Function ShowAmendEventBookingForm(ByVal pParentForm As Form, ByVal pDGR As DisplayGrid, ByVal pEventInfo As CareEventInfo) As Boolean
    Dim vBookingSaved As Boolean = False
    Try
      Dim vBookingNumber As Integer = IntegerValue(pDGR.GetValue(pDGR.CurrentRow, "BookingNumber"))
      If vBookingNumber > 0 Then
        Dim vList As New ParameterList
        vList.IntegerValue("BookingNumber") = vBookingNumber
        vList.IntegerValue("EventNumber") = pEventInfo.EventNumber
        With pDGR
          Dim vOptionNumber As Integer = IntegerValue(.GetValue(.CurrentRow, "OptionNumber"))
          If vOptionNumber = 0 Then vOptionNumber = IntegerValue(.GetValue(.CurrentRow, "BookingOptionNumber"))
          vList.IntegerValue("OptionNumber") = vOptionNumber
          vList("Product") = .GetValue(.CurrentRow, "Product")
          vList("Rate") = .GetValue(.CurrentRow, "Rate")
          vList.IntegerValue("ContactNumber") = IntegerValue(.GetValue(.CurrentRow, "ContactNumber"))
          vList.IntegerValue("Quantity") = IntegerValue(.GetValue(.CurrentRow, "Quantity"))
          If .GetValue(.CurrentRow, "AdultQuantity").Length > 0 Then vList.IntegerValue("AdultQuantity") = IntegerValue(.GetValue(.CurrentRow, "AdultQuantity"))
          If .GetValue(.CurrentRow, "ChildQuantity").Length > 0 Then vList.IntegerValue("ChildQuantity") = IntegerValue(.GetValue(.CurrentRow, "ChildQuantity"))
          vList("StartTime") = .GetValue(.CurrentRow, "StartTime")
          vList("EndTime") = .GetValue(.CurrentRow, "EndTime")
          vList("Amount") = .GetValue(.CurrentRow, "BookingAmount")
          vList("CreditSale") = .GetValue(.CurrentRow, "CreditSale")
          vList("InvoicePrinted") = .GetValue(.CurrentRow, "InvoicePrinted")
          If .GetValue(.CurrentRow, "InvoiceAllocationAmount").Length > 0 Then vList.IntegerValue("InvoiceAllocationAmount") = IntegerValue(.GetValue(.CurrentRow, "InvoiceAllocationAmount"))
        End With
        Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctAmendBooking, vList)
        If vForm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
          vBookingSaved = True
          Dim vReturnList As ParameterList = vForm.ReturnList
          If vReturnList.ContainsKey("ShowDelegates") = True AndAlso BooleanValue(vReturnList("ShowDelegates")) = True Then
            If vReturnList.ContainsKey("NewBookingNumber") Then vBookingNumber = IntegerValue(vReturnList("NewBookingNumber"))
            Dim vContactInfo As New ContactInfo(vList.IntegerValue("ContactNumber"))
            pEventInfo.SetBookingInfo(vBookingNumber, vReturnList.IntegerValue("Quantity"), vContactInfo)
            Dim vDelegateForm As New frmEventSet(pParentForm, pEventInfo, CareServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates)
            vDelegateForm.ShowDialog()
          End If
          MainHelper.RefreshEventData(CareServices.XMLEventDataSelectionTypes.xedtEventBookings, pEventInfo.EventNumber)
          MainHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactEventBookings, vList.IntegerValue("ContactNumber"))
        End If
      End If
    Catch vCareEx As CareException
      If vCareEx.ErrorNumber = CareException.ErrorNumbers.enOriginalPaymentPartProcessed Then
        ShowErrorMessage(vCareEx.Message)
      Else
        DataHelper.HandleException(vCareEx)
      End If
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
    Return vBookingSaved
  End Function

  Private Shared Sub ProcessMailingCriteriaFromAppParam(ByVal pCriteriaSet As Integer, ByRef pList As ParameterList, ByVal pTaskType As CareServices.TaskJobTypes, ByRef pSuccess As Boolean)
    Dim vGeneralMailing As GeneralMailing = New GeneralMailing(Nothing, pTaskType)
    vGeneralMailing.ProcessMailingCriteria(pCriteriaSet, pList, pSuccess)
  End Sub

  Public Shared Function ActionRights(ByVal pActionNumber As Integer) As DataHelper.DocumentAccessRights
    Return ActionsHelper.GetActionRights(pActionNumber)
  End Function

  Public Shared Sub InitWindowForViewType(vForm As Form)
    'sets whatever is required on the form to make sure it's aware of the parent MDI's view, i.e. classic or explorer
    With vForm
      If .MdiParent IsNot Nothing AndAlso TypeOf (.MdiParent) Is frmMain Then
        If AppHelper.FormView = FormViews.Modern Then
          .WindowState = FormWindowState.Maximized
        End If
      End If
    End With

  End Sub

  Shared Sub ShowExplorerMenuViewer(dataContext As Object)
    Dim vOpenForm As frmModernMenuViewer = Nothing
    For Each vItem As Form In MainHelper.Forms
      If TypeOf (vItem) Is frmModernMenuViewer Then
        vOpenForm = DirectCast(vItem, frmModernMenuViewer)
        Exit For
      End If
    Next
    If vOpenForm Is Nothing Then
      vOpenForm = New frmModernMenuViewer()
      vOpenForm.DataContext = dataContext
      vOpenForm.MdiParent = MDIForm
    End If
    InitWindowForViewType(vOpenForm)
    vOpenForm.WindowState = FormWindowState.Maximized
    vOpenForm.Show()
  End Sub

  Shared Function ShowWorkstreamIndex(ByVal pWorkstreamGroup As String) As PagedTreeViewDataContext
    Return ShowWorkstreamIndex(pWorkstreamGroup, -1)
  End Function

  Shared Function ShowWorkstreamIndex(ByVal pWorkstreamGroup As String, ByVal pWorkstreamId As Integer) As PagedTreeViewDataContext
    If pWorkstreamId <= 0 Then pWorkstreamId = -1
    Dim vOpenForms As List(Of frmPagedTreeView) = MainHelper.FindForms(Of frmPagedTreeView)()
    Dim vWorkstreamForm As frmPagedTreeView = Nothing
    Dim vDC As WorkstreamContainerDataContext = Nothing
    If vOpenForms.Count > 0 Then
      For Each vForm In vOpenForms
        If vForm.PagedTreeView.DataContext IsNot Nothing AndAlso TypeOf vForm.PagedTreeView.DataContext Is WorkstreamContainerDataContext Then
          If DirectCast(vForm.PagedTreeView.DataContext, WorkstreamContainerDataContext).WorkstreamGroup = pWorkstreamGroup Then
            vWorkstreamForm = vForm
            vDC = TryCast(vWorkstreamForm.DataContext, WorkstreamContainerDataContext)
            If pWorkstreamId > 0 Then
              Try
                vDC.WorkstreamId = pWorkstreamId
              Catch vEx As CareException
                If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidWorkstreamGroup OrElse
                   vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamGroupReadOnly OrElse
                   vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamNotFound Then
                  ShowErrorMessage(vEx.Message)
                Else
                  Throw
                End If
              End Try
            End If
            Exit For
          End If
        End If
      Next
    End If
    Try
      If vWorkstreamForm Is Nothing Then
        vDC = New WorkstreamContainerDataContext(pWorkstreamGroup, pWorkstreamId)
        vWorkstreamForm = New frmPagedTreeView()
        vWorkstreamForm.MdiParent = MainHelper.MainForm
        InitWindowForViewType(vWorkstreamForm)
        vWorkstreamForm.Show()
        vWorkstreamForm.Init(vDC)
      End If
      vWorkstreamForm.Show()
      vWorkstreamForm.BringToFront()
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidWorkstreamGroup OrElse
         vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamGroupReadOnly OrElse
         vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamNotFound Then
        ShowErrorMessage(vEx.Message)
      Else
        Throw
      End If
    Catch vAccEx As TypeAccessException
      ShowErrorMessage(vAccEx.Message)
    End Try
    Return If(vWorkstreamForm IsNot Nothing, vWorkstreamForm.DataContext, Nothing)
  End Function

  Shared Function ShowWorkstreamIndex(ByVal pWorkstreamGroup As String, ByVal pWorkstreamIds As IList(Of Integer)) As PagedTreeViewDataContext
    If pWorkstreamIds.Count < 1 Then
      Throw New ArgumentException("At least one workstream ID must be given", "pWorkstreamIds")
    End If
    Dim vWorkstreamForm As frmPagedTreeView = Nothing
    Dim vDC As WorkstreamContainerDataContext = Nothing
    Try
      If vWorkstreamForm Is Nothing Then
        vDC = New WorkstreamContainerDataContext(pWorkstreamGroup, pWorkstreamIds)
        vWorkstreamForm = New frmPagedTreeView()
        vWorkstreamForm.MdiParent = MainHelper.MainForm
        InitWindowForViewType(vWorkstreamForm)
        vWorkstreamForm.Show()
        vWorkstreamForm.Init(vDC)
      End If
      vWorkstreamForm.Show()
      vWorkstreamForm.BringToFront()
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidWorkstreamGroup OrElse
         vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamGroupReadOnly OrElse
         vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamNotFound Then
        ShowErrorMessage(vEx.Message)
      Else
        Throw
      End If
    Catch vAccEx As TypeAccessException
      ShowErrorMessage(vAccEx.Message)
    End Try
    Return If(vWorkstreamForm IsNot Nothing, vWorkstreamForm.DataContext, Nothing)
  End Function

  Shared Function ShowWorkstreamIndex(ByVal pWorkstreamId As Integer) As PagedTreeViewDataContext
    Dim vRtn As PagedTreeViewDataContext = Nothing
    Try
      Dim vWorkstreamGroup As String = String.Empty
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("WorkstreamId") = pWorkstreamId
      Dim vDT As DataTable = WorkstreamDataHelper.SelectWorkstreamData("", WorkstreamService.XMLDataSelectionTypes.WorkstreamDetails, vList)
      If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
        vWorkstreamGroup = vDT.Rows(0).Item("WorkstreamGroup").ToString
      End If
      vRtn = ShowWorkstreamIndex(vWorkstreamGroup, pWorkstreamId)
    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
    Return vRtn
  End Function

  Shared Sub NewWorkstream(ByVal pWorkstreamGroupCode As String)
    Dim vTreeViewDataContext As PagedTreeViewDataContext = ShowWorkstreamIndex(pWorkstreamGroupCode)
    If vTreeViewDataContext IsNot Nothing Then
      Try
        vTreeViewDataContext.NewRow()
      Catch vEx As CareException
        If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidWorkstreamGroup OrElse
           vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamGroupReadOnly OrElse
           vEx.ErrorNumber = CareException.ErrorNumbers.enWorkstreamNotFound Then
          ShowErrorMessage(vEx.Message)
        Else
          DataHelper.HandleException(vEx)
        End If
      Catch vEx As Exception
        DataHelper.HandleException(vEx)
      End Try
    End If
  End Sub

  Public Shared Sub ShowLegacy(ByVal pLegacyNumber As Integer)
    Try
      Dim vList As New ParameterList(True)
      vList.Add("LegacyNumber", pLegacyNumber)

      Dim vContactNumber As Integer
      Dim vDataSet As New DataSet
      vDataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftLegacies, vList)
      If vDataSet IsNot Nothing Then
        If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Columns.Contains("LegacyNumber") Then
          Dim vRow As DataRow = vDataSet.Tables("DataRow").Rows(0)
          vContactNumber = IntegerValue(vRow.Item("ContactNumber").ToString)
        End If
      End If
      ShowContactLegacy(vContactNumber)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public Shared Sub ShowContactLegacy(ByVal pContactNumber As Integer)
    FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactLegacy, pContactNumber, False)
  End Sub

  Public Shared Sub ShowContactPosition(ByVal pContactPositionNumber As Integer, ByVal pShowNewWindow As Boolean)
    Try
      'All we have is the ContactPositionNumber so find the ContactNumber
      Dim vList As New ParameterList(True, True)
      vList.IntegerValue("ContactPositionNumber") = pContactPositionNumber

      Dim vContactNumber As Integer = 0
      Dim vDataSet As DataSet = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
      If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Count > 0 AndAlso vDataSet.Tables.Contains("DataRow") Then
        Dim vRow As DataRow = vDataSet.Tables("DataRow").Rows(0)
        If vRow IsNot Nothing Then vContactNumber = IntegerValue(vRow.Item("ContactNumber").ToString)
      End If

      If vContactNumber > 0 Then
        'Got the ContactNumber so open the card and navigate to the required Position
        Dim vForm As Form = FormHelper.ShowCardIndex(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vContactNumber, False, pShowNewWindow)
        If vForm IsNot Nothing AndAlso TypeOf (vForm) Is frmCardSet Then
          Dim vCardSet As frmCardSet = CType(vForm, frmCardSet)
          vCardSet.SelectRowItem("ContactPositionNumber", pContactPositionNumber)
        End If
      End If

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Public Shared Sub ShowViewBatchDetails(ByVal pBatchNumber As Integer)
    Try
      Dim vList As New ParameterList(True)
      vList.IntegerValue("BatchNumber") = pBatchNumber
      DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoLockBatch, vList)
      Dim vForm As New frmCardMaintenance(CareServices.XMLMaintenanceControlTypes.xmctBatchDetails, DataHelper.GetRowFromDataSet(DataHelper.GetFinancialProcessingData(CareNetServices.XMLFinancialProcessingDataSelectionTypes.xbdstBatchInformation, vList)))
      RemoveHandler vForm.EditBatchDetails, AddressOf EditBatchDetails
      RemoveHandler vForm.BatchEditComplete, AddressOf BatchEditComplete
      AddHandler vForm.EditBatchDetails, AddressOf EditBatchDetails
      AddHandler vForm.BatchEditComplete, AddressOf BatchEditComplete
      vForm.TopMost = False
      vForm.Show(MainHelper.MainForm)

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

  Friend Shared Sub EditBatchDetails(ByVal sender As Object, ByVal pBatchInfo As BatchInfo, ByVal pTransactionNumber As Integer)
    Dim vTA As New TraderApplication(IntegerValue(AppValues.ConfigurationValue(AppValues.ConfigurationValues.trader_application_edit_trans)), pBatchInfo.BatchNumber, False, pTransactionNumber)
    vTA.BatchNumber = pBatchInfo.BatchNumber
    vTA.TransactionNumber = pTransactionNumber
    vTA.BatchInfo = pBatchInfo
    vTA.BatchLocked = True
    Dim vList As New ParameterList()
    vList("EditFromBatchDetails") = "Y"
    RunTraderApplication(vTA, vList, Nothing, BatchInfo.AdjustmentTypes.None)
  End Sub

  Friend Shared Sub BatchEditComplete(ByVal sender As Object, ByVal pBatchNumber As Integer)
    'Batch editing from View Batch Details is complete so unlock the batch
    Dim vList As New ParameterList(True)
    vList.IntegerValue("BatchNumber") = pBatchNumber
    DataHelper.UpdateBatch(CareServices.UpdateBatchOptions.buoUnlockBatch, vList)
  End Sub

End Class
