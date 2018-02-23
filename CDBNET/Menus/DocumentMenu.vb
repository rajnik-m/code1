Imports System.IO

Public Class DocumentMenu
  Inherits BaseDocumentMenu

  Public Sub New(ByVal pParent As MaintenanceParentForm)
    MyBase.New(pParent)
  End Sub

  Protected Overrides Sub DoNotify()
    If Not FormHelper.NotifyForm Is Nothing Then FormHelper.NotifyForm.DoRefresh()
  End Sub

  Protected Overrides Function DoRefresh() As Boolean
    If TypeOf mvParent Is frmEventSet Or TypeOf mvParent Is frmCampaignSet Then
      Return True
    Else
      Return False
    End If
  End Function

  Protected Overrides Sub HandleNewDocument(sender As Object, e As EventArgs)
    ProcessNewEditDocument(sender, e)
  End Sub

  Protected Overrides Sub HandleDocumentAction(sender As Object, e As EventArgs)
    ProcessNewEditDocument(sender, e)
  End Sub

  Protected Overrides Sub HandleEditDocument(sender As Object, e As EventArgs)
    ProcessNewEditDocument(sender, e)
  End Sub

  Protected Overrides Sub HandleDeleteDocument(sender As Object, e As EventArgs)
    ProcessNewEditDocument(sender, e)
  End Sub

  Protected Overrides Sub HandleRelatedDocument(sender As Object, e As EventArgs)
    ProcessNewEditDocument(sender, e)
  End Sub

  Private Sub ProcessNewEditDocument(ByVal sender As Object, ByVal e As EventArgs)
    Dim vMenuItem As DocumentMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, DocumentMenuItems)

    Dim vDocumentToEdit As Integer = 0
    If vMenuItem = DocumentMenuItems.dmiEdit Then vDocumentToEdit = DocumentNumber
    Dim vParams As ParameterList
    If TryCast(e, DocumentEventArgs) IsNot Nothing Then
      vParams = TryCast(e, DocumentEventArgs).DocumentParamters
    Else
      vParams = New ParameterList()
    End If

    If ExamCentreId > 0 Then
      vParams.IntegerValue("ExamCentreId") = ExamCentreId
      vParams.IntegerValue("ContactNumber") = If(ExamCentreContact > 0, ExamCentreContact, ExamCentreOrginisation)
    ElseIf ExamCentreUnitId > 0 Then
      vParams.IntegerValue("ExamCentreUnitId") = ExamCentreUnitId
      vParams.IntegerValue("ContactNumber") = If(ExamCentreContact > 0, ExamCentreContact, ExamCentreOrginisation)
    ElseIf ExamUnitLinkId > 0 Then
      vParams.IntegerValue("ExamUnitLinkId") = ExamUnitLinkId
    ElseIf DocumentType = DocumentTypes.CPDCycleDocuments AndAlso CPDPeriodNumber > 0 Then
      vParams.IntegerValue("ContactCpdPeriodNumber") = CPDPeriodNumber
    ElseIf DocumentType = DocumentTypes.CPDPointDocuments AndAlso CPDPointNumber > 0 Then
      vParams.IntegerValue("ContactCpdPointNumber") = CPDPointNumber
    ElseIf DocumentType = DocumentTypes.PositionDocuments AndAlso ContactPositionNumber > 0 Then
      vParams.IntegerValue("ContactPositionNumber") = ContactPositionNumber
    End If

    FormHelper.EditDocument(vDocumentToEdit, mvParent, Nothing, vParams)

  End Sub

  Protected Overrides Sub HandleNewDocumentLink(sender As Object, e As EventArgs)
    Try
      Dim vList As New ParameterList(True)
      vList("CPDDocumentFinder") = "Y"
      Dim vDocumentNumber As Integer = FormHelper.ShowFinder(CareNetServices.XMLDataFinderTypes.xdftDocuments, vList, mvParent)
      If vDocumentNumber > 0 Then
        vList = New ParameterList(True, True)
        vList.IntegerValue("DocumentNumber") = vDocumentNumber
        If CPDPeriodNumber > 0 Then
          vList.IntegerValue("ContactCpdPeriodNumber") = CPDPeriodNumber
        ElseIf CPDPointNumber > 0 Then
          vList.IntegerValue("ContactCpdPointNumber") = CPDPointNumber
        ElseIf ContactPositionNumber > 0 Then
          vList.IntegerValue("ContactPositionNumber") = ContactPositionNumber
        End If
        DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctDocumentLink, vList)
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

  Protected Overrides Sub HandleDeleteDocumentLink(sender As Object, e As EventArgs)
    Try
      If ConfirmDelete() Then
        Dim vList As New ParameterList(True, True)
        vList.IntegerValue("DocumentNumber") = DocumentNumber
        vList("DocumentLinkType") = "R"   'These are always 'related-to' links
        If DocumentType = BaseDocumentMenu.DocumentTypes.CPDCycleDocuments Then
          vList.IntegerValue("ContactCpdPeriodNumber") = CPDPeriodNumber
        ElseIf DocumentType = BaseDocumentMenu.DocumentTypes.CPDPointDocuments Then
          vList.IntegerValue("ContactCpdPointNumber") = CPDPointNumber
        ElseIf DocumentType = BaseDocumentMenu.DocumentTypes.PositionDocuments Then
          vList.IntegerValue("ContactPositionNumber") = ContactPositionNumber
        Else
          'Not supported
          Throw New NotImplementedException("Delete Document Link menu only available for CPD Document Links")
        End If

        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctDocumentLink, vList)
        If mvParent IsNot Nothing Then mvParent.RefreshData()
      End If

    Catch vEX As Exception
      DataHelper.HandleException(vEX)
    End Try
  End Sub

End Class
