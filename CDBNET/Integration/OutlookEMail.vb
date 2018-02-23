Option Strict Off
Imports System.IO
Imports System.Reflection

Public Class OutlookEMail
  Inherits EmailInterface

  Private mvInitialised As Boolean
  Private mvCanEmail As Boolean
  'Private mvOutlook As Outlook.Application
  'Private mvNS As Outlook.NameSpace
  'Private mvInboxFolder As Outlook.MAPIFolder
  'Private mvMailItem As Outlook.MailItem
  Private mvOutlook As Object
  Private mvNS As Object
  Private mvInboxFolder As Object
  Private WithEvents mvMailItem As Object
  Private mvMailSent As Boolean
  Private mvOSM As AddInExpress.OutlookSecurityManager

  Private Enum OutlookItemTypes
    oitUnknown
    oitMailItem
    oitReportItem
    oitMeetingItem
  End Enum

  Public Event RetrievedMessages(ByVal pCount As Long)

  Public Overrides Function CanEMail(Optional ByVal pDownloadMail As Boolean = False) As Boolean
    Dim vUserName As String
    Try
      If mvInitialised Then
        Return mvCanEmail
      Else
        'mvOutlook = New Outlook.Application
        mvOutlook = CreateObject("Outlook.Application")
        mvNS = mvOutlook.GetNamespace("MAPI")
        If UserEMailLogname.Length > 0 Then
          vUserName = UserEMailLogname
        Else
          vUserName = AppValues.Logname
        End If
        mvNS.Logon(vUserName, AppValues.Password, False, True)
        If UseOutlookSecurityManager Then
          Try
            mvOSM = New AddInExpress.OutlookSecurityManager
            mvOSM.ConnectTo(mvOutlook)
          Catch vException As Exception
            UseOutlookSecurityManager = False
            ShowInformationMessage(InformationMessages.imOutlookSecurityManagerFailed)
          End Try
        End If
        Dim vType As System.Type = mvOutlook.GetType
        Dim vEventInfo As EventInfo
        vEventInfo = vType.GetEvent("ApplicationEvents_Event_ItemSend")
        If vEventInfo Is Nothing Then vEventInfo = vType.GetEvent("ApplicationEvents10_Event_ItemSend")
        If vEventInfo Is Nothing Then vEventInfo = vType.GetEvent("ApplicationEvents11_Event_ItemSend")
        If vEventInfo Is Nothing Then
          'ShowInformationMessage(InformationMessages.imCannotAccessOutlookSendEvent)
        Else
          vEventInfo.AddEventHandler(mvOutlook, [Delegate].CreateDelegate(vEventInfo.EventHandlerType, Me, "ItemSendHandler"))
        End If
        mvCanEmail = True
        mvInitialised = True
        Return True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Function

  Private Sub ItemSendHandler(ByVal Item As Object, ByRef Cancel As Boolean)
    Dim vRecipient As Object

    For Each vRecipient In mvMailItem.Recipients
      Select Case vRecipient.Type
        Case 1                                        'Outlook.OlMailRecipientType.olTo = 1
          mvRecipientToList.Add(vRecipient.Name)
        Case 2                                        'Outlook.OlMailRecipientType.olCC = 2
          mvRecipientCCList.Add(vRecipient.Name)
      End Select
    Next
    mvMailSent = True
  End Sub

  Public Overrides Function GetAttachmentPathName(ByVal pMsgID As String, ByVal pIndex As Integer) As String
    'Dim vMailItem As Outlook.MailItem
    'Dim vAttachment As Outlook.Attachment
    Dim vMailItem As Object
    Dim vAttachment As Object
    Dim vPath As String
    Dim vIndex As Integer
    Dim vFileInfo As FileInfo

    For vIndex = 1 To mvInboxFolder.Items.Count
      If TypeName(mvInboxFolder.Items.Item(vIndex)) = "MailItem" Then
        'vMailItem = CType(mvInboxFolder.Items.Item(vIndex), Outlook.MailItem)
        vMailItem = mvInboxFolder.Items.Item(vIndex)
        If pMsgID = vMailItem.EntryID Then
          vAttachment = vMailItem.Attachments.Item(pIndex + 1)
          vFileInfo = New FileInfo(vAttachment.FileName)
          vPath = DataHelper.GetTempFile(vFileInfo.Extension)
          vAttachment.SaveAsFile(vPath)
          Return vPath
          Exit For
        End If
      End If
    Next
    Return ""
  End Function

  Public Overrides Function GetInBox() As System.Collections.ArrayList
    Dim vInBox As New ArrayList
    Dim vIndex As Integer
    Dim vUID As String = ""
    Dim vMsg As EMailMessage
    Dim vSubject As String = ""
    Dim vDeliveredDate As Date
    Dim vTOList As String = ""
    Dim vCCList As String = ""
    Dim vAttachments As ArrayList = Nothing
    Dim vAttachCount As Integer
    Dim vFromName As String = ""
    Dim vFromAddress As String = ""
    Dim vText As String = ""
    Dim vRead As Boolean
    Dim vItem As Object
    'Dim vRecipient As Outlook.Recipient
    'Dim vAttachment As Outlook.Attachment
    'Dim vMailItem As Outlook.MailItem
    'Dim vReportItem As Outlook.ReportItem
    'Dim vMeetingItem As Outlook.MeetingItem
    'Dim vItemType As OutlookItemTypes
    'Dim vOLRecipients As Outlook.Recipients = Nothing
    'Dim vOLAttachments As Outlook.Attachments = Nothing
    Dim vRecipient As Object
    Dim vAttachment As Object
    Dim vMailItem As Object
    Dim vReportItem As Object
    Dim vMeetingItem As Object
    Dim vItemType As OutlookItemTypes
    Dim vOLRecipients As Object = Nothing
    Dim vOLAttachments As Object = Nothing
    Dim vNewItem As Object

    RaiseEvent RetrievedMessages(0)
    If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = True
    mvInboxFolder = mvNS.GetDefaultFolder(6)          'Outlook.OlDefaultFolders.olFolderInbox = 6
    If mvInboxFolder.Items.Count > 0 Then
      For vIndex = 1 To mvInboxFolder.Items.Count
        RaiseEvent RetrievedMessages(vIndex)
        vItem = mvInboxFolder.Items.Item(vIndex)
        Select Case TypeName(vItem)
          Case "ReportItem"
            vReportItem = vItem
            'vReportItem = CType(vItem, Outlook.ReportItem)
            vItemType = OutlookItemTypes.oitReportItem
            With vReportItem
              vUID = .EntryID
              vSubject = .Subject
              vDeliveredDate = .CreationTime
              vText = .Body
              vRead = Not .UnRead
              vFromName = InformationMessages.imSystemMessage
              vFromAddress = vFromName
              vTOList = ""
              vCCList = ""
              vAttachCount = 0
              vAttachments = New ArrayList
            End With

          Case "MeetingItem"
            vMeetingItem = vItem
            'vMeetingItem = CType(vItem, Outlook.MeetingItem)
            vItemType = OutlookItemTypes.oitMeetingItem
            With vMeetingItem
              vUID = .EntryID
              vSubject = .Subject
              vDeliveredDate = .ReceivedTime
              vText = .Body
              vRead = Not .UnRead
              vFromName = .SenderName
              vFromAddress = .SenderName
              vOLRecipients = .Recipients
              vOLAttachments = .Attachments
            End With

          Case "MailItem"
            vMailItem = vItem
            'vMailItem = CType(vItem, Outlook.MailItem)
            vItemType = OutlookItemTypes.oitMailItem
            With vMailItem
              vUID = .EntryID
              vSubject = .Subject
              vDeliveredDate = .ReceivedTime
              vText = .Body
              vRead = Not .UnRead
              vFromName = .SenderName
              vFromAddress = .SenderName
              Try
                vFromAddress = .SenderEMailAddress  'This property is supported in Outlook 2003, but not in Outlook 2000, which is why Try & Catch is used
              Catch vException As Exception
                Try
                  If vFromAddress.IndexOf("@"c) < 0 Then
                    'Simulate a reply to the email...
                    vNewItem = vMailItem.Reply
                    '...in order to capture the sender's email address, which will be the recipient of the reply
                    vFromAddress = vNewItem.Recipients(1).Address
                    If vFromAddress Is Nothing Then vFromAddress = .SenderName 'For some mail items at this stage vFromAddress is Nothing, so set it back to SenderName
                  End If
                Catch vException1 As Exception
                  'Ignore the error
                End Try
              End Try
              vOLRecipients = .Recipients
              vOLAttachments = .Attachments
            End With

          Case Else
            vItemType = OutlookItemTypes.oitUnknown
        End Select

        If vItemType = OutlookItemTypes.oitMailItem Or vItemType = OutlookItemTypes.oitMeetingItem Then
          vTOList = ""
          vCCList = ""
          Dim vRIndex As Integer
          For vRIndex = 1 To vOLRecipients.Count
            vRecipient = vOLRecipients.Item(vRIndex)
            Select Case vRecipient.Type
              Case 1                                                      'Outlook.OlMailRecipientType.olTo = 1
                If vTOList.Length > 0 Then vTOList = vTOList & ", "
                vTOList = vTOList & vRecipient.Name
              Case 2                                                      'Outlook.OlMailRecipientType.olCC = 2
                If vCCList.Length > 0 Then vCCList = vCCList & ", "
                vCCList = vCCList & vRecipient.Name
              Case 0                                                      'Outlook.OlMailRecipientType.olOriginator
                If vFromName.Length > 0 Then vFromName = vFromName & ", "
                vFromName = vFromName & vRecipient.Name
                If vFromAddress.Length > 0 Then vFromAddress = vFromAddress & ", "
                vFromAddress = vFromAddress & vRecipient.Address
            End Select
          Next
          Dim vAIndex As Integer
          vAttachCount = vOLAttachments.Count
          vAttachments = New ArrayList
          For vAIndex = 1 To vOLAttachments.Count
            vAttachment = vOLAttachments.Item(vAIndex)
            vAttachments.Add(vAttachment.DisplayName)
          Next
        End If
        If vItemType <> OutlookItemTypes.oitUnknown Then
          vMsg = New EMailMessage
          vMsg.InitFromMessage(vUID, vSubject, vDeliveredDate.ToString("dd/MM/yyyy HH:mm"), vFromName, vFromAddress, vRead, vTOList, vCCList, vAttachCount, vAttachments)
          vMsg.NoteText = vText
          vInBox.Add(vMsg)
        End If
      Next
    End If
    If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = False
    Return vInBox
  End Function

  Public Overrides Sub MarkRead(ByVal pMsg As EMailMessage)
    Dim vItem As Object

    vItem = GetMailItem(pMsg)
    If vItem Is Nothing Then
      ShowInformationMessage(InformationMessages.imFindEMailFailed, pMsg.Subject)
    Else
      vItem.UnRead = False
    End If
  End Sub

  Public Overloads Overrides Function ProcessAction(ByVal pMsg As EMailMessage, ByVal pAction As EmailInterface.EMailActions) As Boolean
    Dim vMailItem As Object
    'Dim vNewItem As Outlook.MailItem
    Dim vNewItem As Object

    vMailItem = GetMailItem(pMsg)
    If Not vMailItem Is Nothing Then
      Select Case pAction
        Case EMailActions.emaDelete
          vMailItem.Delete()
          Return True
        Case EMailActions.emaReply
          vNewItem = vMailItem.Reply
          vNewItem.Display(True)
          Return True
        Case EMailActions.emaReplyAll
          vNewItem = vMailItem.ReplyAll
          vNewItem.Display(True)
          Return True
        Case EMailActions.emaForward
          vNewItem = vMailItem.Forward
          vNewItem.Display(True)
          Return True
        Case EMailActions.emaSave
          If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = True
          vNewItem = vMailItem.Reply
          pMsg.OrigAddress = vNewItem.Recipients.Item(1).Address
          'pMsg.OrigAddress = GetSenderID(vMailItem)
          If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = False
          Return True
      End Select
    End If
  End Function

  Public Overrides Function SendMail(ByVal pForm As System.Windows.Forms.Form, ByVal pOptions As EmailInterface.SendEmailOptions, ByVal pSubject As String, ByVal pMessage As String, ByVal pEmailAddress As String, Optional ByVal pAttachments As Microsoft.VisualBasic.Collection = Nothing, Optional ByVal pCCList As String = "") As Boolean
    Dim vItem As Object
    Dim vAddresses() As String
    Dim vNullAddress As Boolean
    'Dim vRecipient As Outlook.Recipient
    Dim vRecipient As Object
    Dim vIndex As Integer

    mvMailSent = False
    If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = True
    mvMailItem = mvOutlook.CreateItem(0)              'Outlook.OlItemType.olMailItem = 0
    'mvMailItem = DirectCast(mvOutlook.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
    With mvMailItem
      .Subject = pSubject
      .Body = pMessage
      .ReadReceiptRequested = (pOptions And EmailInterface.SendEmailOptions.seoReceiptRequired) = EmailInterface.SendEmailOptions.seoReceiptRequired
      If Not pAttachments Is Nothing Then
        For Each vItem In pAttachments
          .Attachments.Add(vItem)
        Next
      End If
      If (pOptions And EmailInterface.SendEmailOptions.seoMultipleRecipients) = EmailInterface.SendEmailOptions.seoMultipleRecipients Then
        vAddresses = pEmailAddress.Split(","c)
        If UBound(vAddresses) < 0 Then vNullAddress = True
        For vIndex = 0 To UBound(vAddresses)
          If vAddresses(vIndex).Length > 0 Then
            vRecipient = .Recipients.Add(vAddresses(vIndex))
            vRecipient.Type = 1               'Outlook.OlMailRecipientType.olTo = 1
          Else
            vNullAddress = True
          End If
        Next
        vAddresses = pCCList.Split(","c)
        For vIndex = 0 To UBound(vAddresses)
          If vAddresses(vIndex).Length > 0 Then
            vRecipient = .Recipients.Add(vAddresses(vIndex))
            vRecipient.Type = 2               'Outlook.OlMailRecipientType.olCC = 2
          End If
        Next
        If .Recipients.Count > 0 Then
          If .Recipients.ResolveAll() = False Then vNullAddress = True
        End If
        If vNullAddress Then
          ShowInformationMessage(InformationMessages.imResolveAddress)      'Could not resolve recipient addresses\r\n\r\nPlease select from address book
          pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail
        End If
      ElseIf InStr(pEmailAddress, "@") > 0 Then
        vRecipient = .Recipients.Add(pEmailAddress)
        vRecipient.Type = 1                                     'Outlook.OlMailRecipientType.olTo
        If (pOptions And EmailInterface.SendEmailOptions.seoAddressResolveUI) > 0 Then .Recipients.ResolveAll()
      ElseIf pEmailAddress.Length = 0 Then
        If (pOptions And EmailInterface.SendEmailOptions.seoShowAddressBook) = EmailInterface.SendEmailOptions.seoShowAddressBook Then
          'Don't know how to do this?
        End If
      End If
      'We don't have any way to show the address resolution UI so we must flag to always edit the mail
      If (pOptions And EmailInterface.SendEmailOptions.seoAddressResolveUI) > 0 Then pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail
      mvRecipientToList = New ArrayList
      mvRecipientCCList = New ArrayList
      If (.Recipients.Count > 0) And pMessage.Length > 0 And ((pOptions And EmailInterface.SendEmailOptions.seoAlwaysEditMail) = 0) Then
        .Send()
      Else
        Dim vInspector As Object = .GetInspector
        vInspector.Display(True)
        vInspector.Close(1)         'OLDiscard
        '.Display(True)
      End If
    End With
    If UseOutlookSecurityManager Then mvOSM.DisableOOMWarnings = False
    Return mvMailSent
  End Function

  Private Function GetMailItem(ByVal pMsg As EMailMessage) As Object
    Dim vItem As Object
    Dim vFound As Boolean

    Dim vCursor As New BusyCursor
    Try
      vItem = mvInboxFolder.Items.Find("[Subject] = '" & pMsg.Subject.Replace("'", "''") & "'")
      Do While Not vItem Is Nothing
        If vItem.EntryID = pMsg.ID Then
          Return vItem
          vFound = True
          Exit Do
        ElseIf vItem.Body = pMsg.NoteText And CDate(vItem.ReceivedTime) = CDate(pMsg.DateReceived) Then
          Return vItem
          vFound = True
          Exit Do
        End If
        vItem = mvInboxFolder.Items.FindNext
      Loop
      If Not vFound Then
        Dim vIndex As Integer
        For vIndex = 1 To mvInboxFolder.Items.Count
          vItem = mvInboxFolder.Items.Item(vIndex)
          If vItem.EntryID = pMsg.ID Then
            Return vItem
            Exit For
          End If
        Next
      End If
      Return Nothing
    Finally
      vCursor.Dispose()
    End Try
  End Function

End Class
