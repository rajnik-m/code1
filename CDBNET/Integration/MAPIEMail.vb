Option Strict Off
'Imports MSMAPI

Public Class MAPIEMail
  Inherits EmailInterface

  Private mvInitialised As Boolean
  Private mvCanEmail As Boolean
  'Private mvSession As MAPISession
  'Private mvMessages As MAPIMessagesClass
  Private mvSession As Object
  Private mvMessages As Object
  Private mvOSM As AddInExpress.OutlookSecurityManager

  Public Event RetrievedMessages(ByVal pCount As Long)

  Public Overrides Function CanEMail(Optional ByVal pDownload As Boolean = False) As Boolean
    Dim vUserName As String

    Try
      If mvInitialised Then
        Return mvCanEmail
      Else
        'mvSession = New MAPISession
        mvSession = CreateObject("MSMAPI.MAPISession")
        If UserEMailLogname.Length > 0 Then
          vUserName = UserEMailLogname
        Else
          vUserName = AppValues.Logname
        End If
        With mvSession
          .UserName = vUserName
          .Password = AppValues.Password
          .NewSession = False
          .DownLoadMail = pDownload
          .SignOn()
          mvInitialised = True
          If .SessionID > 0 Then
            If UseOutlookSecurityManager Then
              Try
                mvOSM = New AddInExpress.OutlookSecurityManager
              Catch vEx As Exception
                ShowInformationMessage(InformationMessages.imOutlookSecurityManagerFailed)
                UseOutlookSecurityManager = False
              End Try
            End If
            mvCanEmail = True
          Return True
          End If
        End With
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Function

  Public Overrides Function GetAttachmentPathName(ByVal pMsgID As String, ByVal pIndex As Integer) As String
    Dim vIndex As Integer

    With mvMessages
      For vIndex = 0 To .MsgCount - 1
        .MsgIndex = vIndex
        If .MsgID = pMsgID Then
          .AttachmentIndex = pIndex
          Return (.AttachmentPathName)
        End If
      Next
    End With
    Return ""
  End Function

  Public Overrides Function GetInBox() As System.Collections.ArrayList
    Dim vIndex As Integer
    Dim vMsg As EMailMessage
    Dim vInBox As New ArrayList
    Dim vRecipIndex As Integer
    Dim vTOList As String
    Dim vCCList As String
    Dim vRead As Boolean

    If mvCanEmail Then
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = True
      'mvMessages = New MAPIMessages
      mvMessages = CreateObject("MSMAPI.MAPIMessages")
      With mvMessages
        RaiseEvent RetrievedMessages(0)
        .SessionID = mvSession.SessionID
        .FetchSorted = False
        .FetchUnreadOnly = False
        .Fetch()
        For vIndex = 0 To .MsgCount - 1
          .MsgIndex = vIndex
          RaiseEvent RetrievedMessages(vIndex)
          vTOList = ""
          vCCList = ""
          vRead = .MsgRead
          For vRecipIndex = 0 To .RecipCount - 1
            .RecipIndex = vRecipIndex
            Select Case .RecipType
              Case CType(1, Short)             'RecipTypeConstants.mapToList
                If vTOList.Length > 0 Then vTOList = vTOList & "; "
                vTOList = vTOList & .RecipDisplayName
              Case CType(2, Short)             'RecipTypeConstants.mapCcList
                If vCCList.Length > 0 Then vCCList = vCCList & "; "
                vCCList = vCCList & .RecipDisplayName
            End Select
          Next
          'Don't read the attachment list here as it will create temporary files for all of them
          'Leave it until we need to get them
          vMsg = New EMailMessage
          vMsg.InitFromMessage(.MsgID, .MsgSubject, Date.Parse(.MsgDateReceived).ToString("dd/MM/yyyy HH:mm"), .MsgOrigDisplayName, .MsgOrigAddress, vRead, vTOList, vCCList, .AttachmentCount, Nothing)
          vInBox.Add(vMsg)
        Next
      End With
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = False
    End If
    GetInBox = vInBox

  End Function

  Public Overrides Sub MarkRead(ByVal pMsg As EMailMessage)
    Dim vIndex As Integer
    Dim vIndex2 As Integer
    Dim vList As New ArrayList

    Try
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = True
      With mvMessages
        For vIndex = 0 To .MsgCount - 1
          .MsgIndex = vIndex
          If .MsgID = pMsg.ID Then
            pMsg.NoteText = .MsgNoteText
            For vIndex2 = 0 To .AttachmentCount - 1
              .AttachmentIndex = vIndex2
              vList.Add(.AttachmentName)
            Next
            pMsg.AttachmentCollection = vList
            Exit For
          End If
        Next
      End With
    Finally
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = False
    End Try
  End Sub

  Public Overloads Overrides Function ProcessAction(ByVal pMsg As EMailMessage, ByVal pAction As EmailInterface.EMailActions) As Boolean
    Dim vIndex As Integer

    Try
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = True
      With mvMessages
        For vIndex = 0 To .MsgCount - 1
          .MsgIndex = vIndex
          If .MsgID = pMsg.ID Then
            Select Case pAction
              Case EmailInterface.EMailActions.emaDelete
                .Delete()
                Return True
              Case EmailInterface.EMailActions.emaReply
                .Reply()
                .Send(True)
                Return True
              Case EmailInterface.EMailActions.emaReplyAll
                .ReplyAll()
                .Send(True)
                Return True
              Case EmailInterface.EMailActions.emaForward
                .Forward()
                .Send(True)
                Return True
            End Select
            Exit For
          End If
        Next
      End With
    Catch vComEx As System.Runtime.InteropServices.COMException
      Select Case vComEx.ErrorCode And 32767
        Case 32001                'MAPIErrors.mapUserAbort
          'Do nothing
        Case Else
          DataHelper.HandleException(vComEx)
      End Select
    Finally
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = False
    End Try
  End Function

  Public Overrides Function SendMail(ByVal pForm As System.Windows.Forms.Form, ByVal pOptions As EmailInterface.SendEmailOptions, ByVal pSubject As String, ByVal pMessage As String, ByVal pEmailAddress As String, Optional ByVal pAttachments As Microsoft.VisualBasic.Collection = Nothing, Optional ByVal pCCList As String = "") As Boolean
    Dim vIndex As Integer
    Dim vItem As Object
    Dim vAddresses() As String
    Dim vNullAddress As Boolean
    Dim vPosition As Integer
    Dim vMsg As Object

    vMsg = CreateObject("MSMAPI.MAPIMessages")
    Try
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = True
      If mvCanEmail Then
        With vMsg
          .SessionID = mvSession.SessionID
          .Compose()
          .MsgSubject = pSubject
          .MsgReceiptRequested = (pOptions And EmailInterface.SendEmailOptions.seoReceiptRequired) = EmailInterface.SendEmailOptions.seoReceiptRequired

          If Not pAttachments Is Nothing Then
            .MsgNoteText = pMessage & vbCrLf & Space$(pAttachments.Count)
            vPosition = pMessage.Length + 2
            For Each vItem In pAttachments
              .AttachmentIndex = vIndex
              .AttachmentPosition = vPosition
              .AttachmentPathName = vItem.ToString
              vIndex = vIndex + 1
              vPosition = vPosition + 1
            Next
          Else
            .MsgNoteText = pMessage
          End If

          If (pOptions And EmailInterface.SendEmailOptions.seoMultipleRecipients) = EmailInterface.SendEmailOptions.seoMultipleRecipients Then
            vAddresses = pEmailAddress.Split(","c)
            If UBound(vAddresses) < 0 Then vNullAddress = True
            For vIndex = 0 To UBound(vAddresses)
              If vAddresses(vIndex).Length > 0 Then
                .RecipIndex = vIndex
                .RecipType = CType(1, Short)         'RecipTypeConstants.mapToList
                .RecipAddress = vAddresses(vIndex)
                .RecipDisplayName = .RecipAddress
              Else
                vNullAddress = True
              End If
            Next
            vAddresses = pCCList.Split(","c)
            For vIndex = 0 To UBound(vAddresses)
              If vAddresses(vIndex).Length > 0 Then
                .RecipIndex = vIndex
                .RecipType = CType(2, Short)         'RecipTypeConstants.mapCcList
                .RecipAddress = vAddresses(vIndex)
                .RecipDisplayName = .RecipAddress
              End If
            Next
            If .RecipCount > 0 Then
              For vIndex = 0 To .RecipCount - 1
                .RecipIndex = vIndex
                .AddressResolveUI = (pOptions And EmailInterface.SendEmailOptions.seoAddressResolveUI) > 0
                Try
                  .ResolveName()
                Catch vComEx As System.Runtime.InteropServices.COMException
                  If (vComEx.ErrorCode And 32767) = 32026 Then                                'MAPIErrors.mapNotSupported 
                    pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail  'Could not resolve so edit mail
                  Else
                    Throw vComEx
                  End If
                End Try
              Next
            End If
            If vNullAddress Then
              ShowInformationMessage(InformationMessages.imResolveAddress)      'Could not resolve recipient addresses\r\n\r\nPlease select from address book
              pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail
            End If
          ElseIf InStr(pEmailAddress, "@") > 0 Then
            .RecipIndex = 0
            .RecipType = CType(1, Short)                                                    'RecipTypeConstants.mapToList
            .RecipAddress = pEmailAddress
            .RecipDisplayName = pEmailAddress
            .AddressResolveUI = (pOptions And EmailInterface.SendEmailOptions.seoAddressResolveUI) > 0
            Try
              .ResolveName()
            Catch vComEx As System.Runtime.InteropServices.COMException
              If (vComEx.ErrorCode And 32767) = 32026 Then                                  'MAPIErrors.mapNotSupported 
                pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail    'Could not resolve so edit mail
              Else
                Throw vComEx
              End If
            End Try
          ElseIf pEmailAddress.Length = 0 Then
            If (pOptions And EmailInterface.SendEmailOptions.seoShowAddressBook) = EmailInterface.SendEmailOptions.seoShowAddressBook Then
              .AddressCaption = "Select Addresses For Notification"
              .AddressEditFieldCount = 2
              Try
                .Action = CType(11, Short)                                                  'MessagesActionConstants.mapShowAddressBook
              Catch vComEx As System.Runtime.InteropServices.COMException
                If (vComEx.ErrorCode And 32767) = 32026 Then                                'MAPIErrors.mapNotSupported 
                  pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail  'Could not resolve so edit mail
                Else
                  Throw vComEx
                End If
              End Try
            End If
          End If
          If (pOptions And EmailInterface.SendEmailOptions.seoForceAddressBook) = EmailInterface.SendEmailOptions.seoForceAddressBook Then
            .AddressCaption = "Select Addresses For Delivery"
            .AddressEditFieldCount = 2
            Try
              Do
                .Show(False)
              Loop While .RecipCount <= 0
            Catch vComEx As System.Runtime.InteropServices.COMException
              If (vComEx.ErrorCode And 32767) = 32026 Then                                  'MAPIErrors.mapNotSupported 
                pOptions = pOptions Or EmailInterface.SendEmailOptions.seoAlwaysEditMail    'Could not resolve so edit mail
              Else
                Throw vComEx
              End If
            End Try
          End If
          If (.RecipCount > 0) And pMessage.Length > 0 And ((pOptions And EmailInterface.SendEmailOptions.seoAlwaysEditMail) = 0) Then
            Try
              .Send()
            Catch vComEx As System.Runtime.InteropServices.COMException
              Select Case vComEx.ErrorCode And 32767
                Case 32014, 32021, 32025                                                         'MAPIErrors.mapUnknownRecipient, MAPIErrors.mapAmbiguousRecipient, MAPIErrors.mapInvalidRecips
                  ShowInformationMessage(InformationMessages.imResolveAddress)      'Could not resolve recipient addresses\r\n\r\nPlease select from address book
                  While .RecipCount > 0
                    .RecipIndex = 0
                    .Delete(1)                                'DeleteConstants.mapRecipientDelete
                  End While
                  .Send(True)
                Case 32002                                    'MAPIErrors.mapFailure
                  ShowInformationMessage(InformationMessages.imCheckEMailAddress)
                  .Send(True)
                Case Else
                  Throw vComEx
              End Select
            End Try
          Else
            .Send(True)
          End If
          mvRecipientToList = New ArrayList
          mvRecipientCCList = New ArrayList
          For vIndex = 0 To .RecipCount - 1
            .RecipIndex = vIndex
            If .RecipType = 1 Then                                'RecipTypeConstants.mapToList
              mvRecipientToList.Add(.RecipDisplayName)
            Else
              mvRecipientCCList.Add(.RecipDisplayName)
            End If
          Next
          Return True
        End With
      End If
    Catch vComEx As System.Runtime.InteropServices.COMException
      Select Case vComEx.ErrorCode And 32767
        Case 32001                                                      'MAPIErrors.mapUserAbort
          'Do nothing
        Case 32002, 32007                                               'MAPIErrors.mapFailure, MAPIErrors.mapGeneralFailure
          If pMessage.Length = 0 Then
            ShowInformationMessage(InformationMessages.imNoEMailText)
          Else
            DataHelper.HandleException(vComEx)
          End If
        Case Else
          DataHelper.HandleException(vComEx)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      If UseOutlookSecurityManager Then mvOSM.DisableSMAPIWarnings = False
    End Try
  End Function
End Class
