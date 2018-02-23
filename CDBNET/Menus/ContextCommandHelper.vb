
Imports CareServices = CDBNETCL.CareNetServices
Imports CDBNETCL.My.Resources
Public Class ContextCommandHelper
  'This provides help for creating and processing menu items for associated objects.
  'For example, when in Contact Meeting node and if Links is to a document, menu items should be for the Meeting Link
  ' and for the associated Document. This helper will assist in creating the menu items for the Document and their navigation.  
  Public Enum AssociatedObject
    cciCommunicationsLog
  End Enum
  Private mvAssociatedObject As AssociatedObject
  Private mvContextCommandItems As CollectionList(Of MenuToolbarCommand)
  Private mvAssociatedDataRow As DataRow
  Private mvAssociatedObjectNumber As Integer  'This is the primary key e.g. DocumentNumber
  Public ReadOnly Property ContextCommandItems As CollectionList(Of MenuToolbarCommand)
    Get
      Return mvContextCommandItems
    End Get
  End Property
  Public Sub New(ByVal pAssociatedObject As AssociatedObject, ByVal pNumber As Integer)
    mvAssociatedObject = pAssociatedObject
    mvContextCommandItems = New CollectionList(Of MenuToolbarCommand)

    Try
      Select Case mvAssociatedObject
        Case AssociatedObject.cciCommunicationsLog
          mvAssociatedObjectNumber = pNumber
          Dim vDataSet As DataSet = DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentInformation, pNumber)
          mvAssociatedDataRow = vDataSet.Tables(0).Rows(0)

          '(Following code is developed from BaseDocumentMenu New method and can be extended to include other menu items) 
          mvContextCommandItems.Add(BaseDocumentMenu.DocumentMenuItems.dmiEdit.ToString, New MenuToolbarCommand(BaseDocumentMenu.DocumentMenuItems.dmiEdit.ToString, ControlText.MnuDocumentEdit2, BaseDocumentMenu.DocumentMenuItems.dmiEdit, "CDDPUP"))
          mvContextCommandItems.Add(BaseDocumentMenu.DocumentMenuItems.dmiViewDocument.ToString, New MenuToolbarCommand(BaseDocumentMenu.DocumentMenuItems.dmiViewDocument.ToString, ControlText.MnuDocumentView, BaseDocumentMenu.DocumentMenuItems.dmiViewDocument, "CDDPVD"))
          mvContextCommandItems.Add(BaseDocumentMenu.DocumentMenuItems.dmiEditDocument.ToString, New MenuToolbarCommand(BaseDocumentMenu.DocumentMenuItems.dmiEditDocument.ToString, ControlText.MnuDocumentEditDocument, BaseDocumentMenu.DocumentMenuItems.dmiEditDocument, "CDDPED"))
          mvContextCommandItems.Add(BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument.ToString, New MenuToolbarCommand(BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument.ToString, ControlText.MnuDocumentPrintDocument, BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument, "CDDPPR"))
          mvContextCommandItems.Add(BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails.ToString, New MenuToolbarCommand(BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails.ToString, ControlText.MnuDocumentPrintDetails2, BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails, "CDDPPD"))
      End Select
    Catch ex As Exception
      'If an error do not add any commands to the collection
    End Try
  End Sub
  Public Sub SetVisibleItems(ByRef pMenuStrip As ContextMenuStrip)
    'This routine enables or disables Tool Strip Item 

    '(Following code is developed from BaseDocumentMenu SetVisibleItems method and can be extended to include other menu items) 
    Select Case mvAssociatedObject
      Case AssociatedObject.cciCommunicationsLog
        Dim vDocumentExtension As String
        Dim vKey As String
        If Not mvAssociatedDataRow Is Nothing Then
          Dim vRights As DataHelper.DocumentAccessRights = CType(mvAssociatedDataRow.Item("AccessRights"), DataHelper.DocumentAccessRights)
          vDocumentExtension = mvAssociatedDataRow.Item("ExternalApplicationExtension").ToString
          Dim vGotExtension As Boolean = vDocumentExtension.Length > 0

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiEdit.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = (vRights And DataHelper.DocumentAccessRights.darEditHeader) = DataHelper.DocumentAccessRights.darEditHeader
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = (vRights And DataHelper.DocumentAccessRights.darHeader) = DataHelper.DocumentAccessRights.darHeader
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiEditDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = vGotExtension AndAlso (vRights And DataHelper.DocumentAccessRights.darEdit) = DataHelper.DocumentAccessRights.darEdit AndAlso mvAssociatedDataRow.Item("DocumentStyle").ToString <> CStr(DataHelper.DocumentStyles.dsnScannedImage)
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiViewDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = vGotExtension AndAlso (vRights And DataHelper.DocumentAccessRights.darView) = DataHelper.DocumentAccessRights.darView
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = vGotExtension AndAlso (vRights And DataHelper.DocumentAccessRights.darPrint) = DataHelper.DocumentAccessRights.darPrint
          End If
        Else
          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiEdit.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = False
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = False
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiViewDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = False
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiEditDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = False
          End If

          vKey = "msi" + BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument.ToString
          If pMenuStrip.Items.ContainsKey(vKey) Then
            pMenuStrip.Items(vKey).Enabled = False
          End If
        End If
    End Select

  End Sub
  Public Sub MenuHandler(ByVal pContextCommand As MenuToolbarCommand, ByVal pParentForm As MaintenanceParentForm, ByVal e As System.EventArgs)

    '(Following code is developed from BaseDocumentMenu MenuHandler method and can be extended to include other menu items) 
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As BaseDocumentMenu.DocumentMenuItems = CType(pContextCommand.CommandID, BaseDocumentMenu.DocumentMenuItems)

      Dim vForm As frmCardMaintenance = Nothing
      Select Case vMenuItem
        Case BaseDocumentMenu.DocumentMenuItems.dmiEdit
          FormHelper.EditDocument(mvAssociatedObjectNumber, pParentForm)
        Case BaseDocumentMenu.DocumentMenuItems.dmiPrintDetails
          Dim vList As ParameterList = New ParameterList(True)
          vList("ReportCode") = "COMLOG"
          vList.IntegerValue("RP1") = mvAssociatedObjectNumber
          vList("RP2") = AppValues.Logname
          Call (New PrintHandler).PrintReport(vList, PrintHandler.PrintReportOutputOptions.AllowSave)
        Case Else
          Dim vDocumentExtension As String = mvAssociatedDataRow.Item("ExternalApplicationExtension").ToString
          If vDocumentExtension.ToUpper = ".PDF" Then
            Dim vApplication As ExternalApplication = GetDocumentApplication(vDocumentExtension)
            Select Case vMenuItem
              Case BaseDocumentMenu.DocumentMenuItems.dmiViewDocument
                vApplication.ViewDocument(mvAssociatedObjectNumber, vDocumentExtension)
              Case BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument
                vApplication.PrintDocument(mvAssociatedObjectNumber, vDocumentExtension)
            End Select
            vApplication = Nothing
            DocumentApplication = Nothing
          Else
            Dim vApplication As ExternalApplication = GetDocumentApplication(vDocumentExtension)
            AddHandler vApplication.ActionComplete, AddressOf ActionComplete
            Select Case vMenuItem
              Case BaseDocumentMenu.DocumentMenuItems.dmiViewDocument
                vApplication.ViewDocument(mvAssociatedObjectNumber, vDocumentExtension)
              Case BaseDocumentMenu.DocumentMenuItems.dmiPrintDocument
                vApplication.PrintDocument(mvAssociatedObjectNumber, vDocumentExtension)
              Case BaseDocumentMenu.DocumentMenuItems.dmiEditDocument
                vApplication.EditDocument(mvAssociatedObjectNumber, vDocumentExtension)
            End Select
          End If
      End Select
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enCannotDeleteDocument Then
        ShowInformationMessage(vEx.Message)
      Else
        DataHelper.HandleException(vEx)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try

  End Sub
  Private Sub ActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFilename As String)
    Select Case pAction
      Case ExternalApplication.DocumentActions.daEditing
        If ShowQuestion(QuestionMessages.QmConfirmSaveEditedDoc, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
          DataHelper.UpdateDocumentFile(mvAssociatedObjectNumber, pFilename)
          'Don't need to add the document history as the external application will have done this for us
          DocumentApplication.CleanupTemporaryObjects()
        End If
    End Select
    DocumentApplication = Nothing
  End Sub

End Class
