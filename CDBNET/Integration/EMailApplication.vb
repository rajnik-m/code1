Imports System.IO
Imports System.Text
'Sourcesafe test shared file-Pooja 
Public Class EMailApplication

  Private Shared mvEMailInterface As EmailInterface
  Private Shared WithEvents mvOutlookEMail As OutlookEMail
  Private Shared WithEvents mvMAIPIEMail As MAPIEMail

  Public Shared ReadOnly Property EmailInterface() As EmailInterface
    Get
      If mvEMailInterface Is Nothing Then
        Select Case AppValues.ConfigurationValue(AppValues.ConfigurationValues.email_interface).ToUpper
          Case "OUTLOOK"
            mvOutlookEMail = New OutlookEMail
            mvEMailInterface = mvOutlookEMail
          Case Else
            mvMAIPIEMail = New MAPIEMail
            mvEMailInterface = mvMAIPIEMail
        End Select
      End If
      Return mvEMailInterface
    End Get
  End Property

  Public Shared Sub SendDocumentAsEMail(ByVal pForm As Form, ByVal pDocumentNumber As Integer, ByVal pTable As DataTable)
    Dim vFileName As String
    Dim vFileNames As New ArrayListEx
    Dim vOurReferences As New ArrayListEx
    Dim vSubjects As New ArrayListEx
    Dim vAttachments As New Collection
    Dim vNotes As New StringBuilder

        'Just a comment-changed by pooja

    If pTable IsNot Nothing Then
      For Each vAttachmentRow As DataRow In pTable.Rows
        If CBool(vAttachmentRow.Item("Select").ToString) Then
          If vAttachmentRow.Item("DocumentSource").ToString = "W" And Not vAttachmentRow.Item("WordProcessorDocument").ToString = "Y" Then
            'This is a precis only so save the precis as a temporary file
            vFileName = DataHelper.GetTempFile(".txt")
            My.Computer.FileSystem.WriteAllText(vFileName, vAttachmentRow.Item("Precis").ToString, False)
          Else
            vFileName = DataHelper.GetDocumentFile(CInt(vAttachmentRow.Item("DocumentNumber")), ".doc")
          End If
          vAttachments.Add(vFileName)
          Dim vFileInfo As New FileInfo(vFileName)
          vFileNames.Add(vFileInfo.Name)
          vOurReferences.Add(vAttachmentRow.Item("OurReference"))
          vSubjects.Add(vAttachmentRow.Item("Subject"))
        End If
      Next
    End If
    Dim vParams As New ParameterList(True)
    vParams("Filenames") = vFileNames.CRLFList
    vParams("OurReferences") = vOurReferences.CRLFList
    vParams("Subjects") = vSubjects.CRLFList
    Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetDocumentData(CareServices.XMLDocumentDataSelectionTypes.xddtDocumentEMailDetails, pDocumentNumber, vParams))
    Dim vOptions As EmailInterface.SendEmailOptions = CDBNET.EmailInterface.SendEmailOptions.seoAddressResolveUI Or CDBNET.EmailInterface.SendEmailOptions.seoReceiptRequired Or CDBNET.EmailInterface.SendEmailOptions.seoMultipleRecipients Or CDBNET.EmailInterface.SendEmailOptions.seoForceAddressBook
    If EmailInterface.SendMail(pForm, vOptions, vRow.Item("Subject").ToString, vRow.Item("BodyText").ToString, vRow.Item("AddressTo").ToString, vAttachments, vRow.Item("CCList").ToString) Then
      For Each vRecipient As String In EmailInterface.LastRecipientToList
        vNotes.Append("To: ")
        vNotes.AppendLine(vRecipient)
      Next
      For Each vRecipient As String In EmailInterface.LastRecipientCCList
        vNotes.Append("Cc: ")
        vNotes.AppendLine(vRecipient)
      Next
      DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaEMailed, pDocumentNumber, vNotes.ToString)
      Dim vAttachNotes As StringBuilder
      If pTable IsNot Nothing Then
        For Each vAttachmentRow As DataRow In pTable.Rows
          If CInt(vAttachmentRow.Item("DocumentNumber")) <> pDocumentNumber Then
            vAttachNotes = New StringBuilder
            vAttachNotes.Append("As attachment to document: ")
            vAttachNotes.AppendLine(pDocumentNumber.ToString)
            vAttachNotes.Append(vNotes)
            DataHelper.AddDocumentHistory(CareServices.XMLDocumentHistoryActions.xdhaEMailed, CInt(vAttachmentRow.Item("DocumentNumber")), vAttachNotes.ToString)
          End If
        Next
      End If
    End If
    For Each vFileName In vAttachments
      My.Computer.FileSystem.DeleteFile(vFileName)
    Next
  End Sub

  Private Shared Sub EMailInterface_RetrievedMessages(ByVal pCount As Long) Handles mvOutlookEMail.RetrievedMessages, mvMAIPIEMail.RetrievedMessages
    FormHelper.MainForm.SetStatusMessage(GetInformationMessage(InformationMessages.imProcessedInBox, pCount.ToString))
  End Sub
End Class
