Namespace Access

  Partial Public Class CommunicationsLog

    Private mvWordProcessorDocument As Boolean
    Private mvMailMergeLetter As Boolean
    Private mvReadDocumentType As Boolean
    Private mvDocumentSource As ExternalApplication.DocumentSourceTypes

    Public Function GetRecordSetFieldsDetail() As String
      Dim vAttrs As String = "cl.communications_log_number,cl.document_type,created_by,precis,cl.document_class,direction,our_reference,their_reference,cl.department,standard_document,archiver,recipient,forwarded,completed,dated,received,cl.package,in_use_by,source," & mvEnv.Connection.DBSpecialCol("cl", "distributed") & ",subject,contact_number,address_number,call_duration,total_duration,email_to,email_from,email_cc,email_bcc,email_reply_to,email_body_is_html,cl.original_uri"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbTelemarketing) Then vAttrs &= ",cl.selection_set"
      Return vAttrs
    End Function

    Public Sub SetCommunicationsLogNumber()
      mvClassFields.SetControlNumber(mvEnv)
      mvClassFields.Item(CommunicationsLogFields.OurReference).Value = GetDocumentReference(CommunicationsLogNumber)
    End Sub

    Public Function GetDocumentReference(ByVal pDocNo As Integer) As String
      Dim vString As String
      Dim vChar As String

      Dim vDocRef As String = mvEnv.GetConfig("our_reference_format")
      If vDocRef.Length = 0 Then vDocRef = "CDB/initials/docno"
      Dim vInitials As String = ""
      If mvEnv.User.FullName.Length = 0 Then
        vInitials = "..."
      Else
        vString = CapitaliseWords(mvEnv.User.FullName, True, False)
        For vIndex As Integer = 0 To vString.Length - 1
          vChar = vString(vIndex)
          If vChar = vChar.ToUpper And vChar >= "A" And vChar <= "Z" Then vInitials = vInitials & vChar
        Next
      End If
      vDocRef = vDocRef.Replace("initials", vInitials)
      Return vDocRef.Replace("docno", pDocNo.ToString)
    End Function

    Public Overloads Sub Create(ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pDocumentType As String, ByRef pDocumentClass As String, ByRef pDirection As String, ByRef pOurReference As String, ByRef pDated As String, ByRef pSource As String, ByRef pSubject As String, ByRef pPrecis As String, Optional ByRef pStandardDocument As String = "", Optional ByRef pPackage As String = "", Optional ByRef pTheirReference As String = "", Optional ByRef pDocumentNumber As Integer = 0, Optional ByRef pDistributed As String = "")

      With mvClassFields
        If pDocumentNumber > 0 Then .Item(CommunicationsLogFields.CommunicationsLogNumber).IntegerValue = pDocumentNumber
        SetValid()
        .Item(CommunicationsLogFields.ContactNumber).IntegerValue = pContactNumber
        .Item(CommunicationsLogFields.AddressNumber).IntegerValue = pAddressNumber
        .Item(CommunicationsLogFields.DocumentType).Value = pDocumentType
        .Item(CommunicationsLogFields.DocumentClass).Value = pDocumentClass
        .Item(CommunicationsLogFields.Direction).Value = pDirection
        .Item(CommunicationsLogFields.Dated).Value = pDated
        .Item(CommunicationsLogFields.Source).Value = pSource
        .Item(CommunicationsLogFields.Subject).Value = pSubject
        .Item(CommunicationsLogFields.Precis).Value = pPrecis
        .Item(CommunicationsLogFields.Department).Value = mvEnv.User.Department
        .Item(CommunicationsLogFields.CreatedBy).Value = mvEnv.User.Logname
        If pOurReference.Length > 0 Then
          .Item(CommunicationsLogFields.OurReference).Value = pOurReference
        Else
          .Item(CommunicationsLogFields.OurReference).Value = GetDocumentReference(CommunicationsLogNumber)
        End If
        .Item(CommunicationsLogFields.StandardDocument).Value = pStandardDocument
        .Item(CommunicationsLogFields.Package).Value = pPackage
        .Item(CommunicationsLogFields.TheirReference).Value = pTheirReference
        .Item(CommunicationsLogFields.Distributed).Value = CStr(IIf(pDistributed = "", "N", pDistributed))
      End With
    End Sub


    Public Sub CreateCopy(ByVal pDated As String, Optional ByVal pReply As Boolean = False, Optional ByVal pContactNumber As Integer = 0, Optional ByVal pAddressNumber As Integer = 0)
      Dim vCommsLogNumber As Integer
      Dim vCLH As New CommunicationsLogHistory

      With mvClassFields
        vCommsLogNumber = CommunicationsLogNumber
        .Item(CommunicationsLogFields.CommunicationsLogNumber).IntegerValue = 0
        SetValid()
        If pContactNumber > 0 Then .Item(CommunicationsLogFields.ContactNumber).IntegerValue = pContactNumber
        If pAddressNumber > 0 Then .Item(CommunicationsLogFields.AddressNumber).IntegerValue = pAddressNumber
        If pReply Then .Item(CommunicationsLogFields.Direction).Value = "O"
        .Item(CommunicationsLogFields.Dated).Value = pDated
        .Item(CommunicationsLogFields.CreatedBy).Value = mvEnv.User.Logname
        .Item(CommunicationsLogFields.OurReference).Value = GetDocumentReference(CommunicationsLogNumber)
        .Item(CommunicationsLogFields.Distributed).Value = "N"
        .Item(CommunicationsLogFields.Department).Value = mvEnv.User.Department
        'Reset the SetValue property of the following fields to force the s/w to save this copy with the values of the document being copied
        'This has to be done since we reset mvExisting below and the ClassFields.Save method only 'saves' those classfields whose value has changed.
        'Resettting the SetValue property like this causes the ClassFields.Save method to think that all the fields have changed.
        .Item(CommunicationsLogFields.ContactNumber).SetValueOnly = ""
        .Item(CommunicationsLogFields.AddressNumber).SetValueOnly = ""
        .Item(CommunicationsLogFields.DocumentClass).SetValueOnly = ""
        .Item(CommunicationsLogFields.Direction).SetValueOnly = ""
        .Item(CommunicationsLogFields.Source).SetValueOnly = ""
        .Item(CommunicationsLogFields.Subject).SetValueOnly = ""
        .Item(CommunicationsLogFields.Precis).SetValueOnly = ""

        'Clear the standard document and package to make it a precis only document
        .Item(CommunicationsLogFields.StandardDocument).Value = ""
        .Item(CommunicationsLogFields.Package).Value = ""
        .Item(CommunicationsLogFields.DocumentType).Value = ""

        '.Item(clfDocumentType).SetValueOnly = ""
        '.Item(clfStandardDocument).SetValueOnly = ""
        '.Item(clfPackage).SetValueOnly = ""
        .Item(CommunicationsLogFields.TheirReference).SetValueOnly = ""
        .Item(CommunicationsLogFields.Dated).SetValueOnly = ""
        .Item(CommunicationsLogFields.CreatedBy).SetValueOnly = ""
        .Item(CommunicationsLogFields.Distributed).SetValueOnly = ""
        .Item(CommunicationsLogFields.Department).SetValueOnly = ""
      End With
      mvExisting = False
      mvEnv.Connection.StartTransaction()
      'Copy data from the original comms log to the copy
      CopyData("communications_log_links", vCommsLogNumber, mvClassFields.Item(CommunicationsLogFields.CommunicationsLogNumber).LongValue, pReply)
      CopyData("communications_log_doc_links", vCommsLogNumber, mvClassFields.Item(CommunicationsLogFields.CommunicationsLogNumber).LongValue, pReply)
      CopyData("communications_log_subjects", vCommsLogNumber, mvClassFields.Item(CommunicationsLogFields.CommunicationsLogNumber).LongValue, pReply)
      'Create history for the copy
      vCLH.Init(mvEnv)
      vCLH.Create((mvClassFields.Item(CommunicationsLogFields.CommunicationsLogNumber).LongValue), CommunicationsLogHistory.CommunicationsLogHistoryActions.clhaCreated)
      vCLH.Save()
      'Save the copy
      Save(mvEnv.User.Logname)
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub CopyData(ByVal pTableName As String, ByVal pOrigCommsLogNumber As Integer, ByVal pNewCommsLogNumber As Integer, Optional ByVal pReply As Boolean = False)
      Dim vSQL As String = ""
      Dim vColumnList As String = ""
      Dim vSelectList As String
      Dim vCLL As CommunicationsLogLink
      Dim vCLLCopy As CommunicationsLogLink
      Dim vAttr As String

      If Not (pReply And pTableName = "communications_log_doc_links") Then
        'Build two lists containing the names of the columns on the specified table
        vAttr = "communications_log_number"
        Select Case pTableName
          Case "communications_log_links"
            vColumnList = "communications_log_number,contact_number,address_number,link_type,processed,notified"
          Case "communications_log_doc_links"
            vAttr = "communications_log_number_1"
            vColumnList = "communications_log_number_1,communications_log_number_2,amended_by,amended_on"
          Case "communications_log_subjects"
            vColumnList = "communications_log_number,topic,sub_topic," & mvEnv.Connection.DBSpecialCol(pTableName, "primary") & ",amended_on,amended_by,quantity"
        End Select
        vSelectList = Mid(vColumnList, InStr(vColumnList, ",") + 1)
        vSelectList = Replace(vSelectList, "amended_by", "'" & mvEnv.User.Logname & "'")
        vSelectList = Replace(vSelectList, "amended_on", mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, "today"))
        'Construct the SQL statement
        vSQL = "INSERT INTO " & pTableName & " (" & vColumnList & ") SELECT " & pNewCommsLogNumber & ", " & vSelectList
        vSQL = vSQL & " FROM " & pTableName & " WHERE " & vAttr & " = " & pOrigCommsLogNumber
      End If
      'Execute the SQL statement
      Select Case pTableName
        Case "communications_log_links"
          If pReply Then
            'copy sender link as addressee link
            vCLL = New CommunicationsLogLink
            vCLL.Init(mvEnv, pOrigCommsLogNumber, CommunicationsLogLink.CommunicationLogLinkTypes.clltSender)
            If vCLL.Existing Then
              vCLLCopy = New CommunicationsLogLink
              vCLLCopy.Create(CommunicationsLogLink.DocumentLinkTypes.dltDocumentToContact, mvEnv, pNewCommsLogNumber, vCLL.ContactNumber, vCLL.AddressNumber, CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee, False, False)
            End If
            'copy addressee link as sender link
            vCLL = New CommunicationsLogLink
            vCLL.Init(mvEnv, pOrigCommsLogNumber, CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee)
            If vCLL.Existing Then
              vCLLCopy = New CommunicationsLogLink
              vCLLCopy.Create(CommunicationsLogLink.DocumentLinkTypes.dltDocumentToContact, mvEnv, pNewCommsLogNumber, vCLL.ContactNumber, vCLL.AddressNumber, CommunicationsLogLink.CommunicationLogLinkTypes.clltSender, False, False)
            End If
            'copy all other links
            mvEnv.Connection.ExecuteSQL(vSQL & " AND link_type NOT IN ('A','S')")
          Else
            'copy all links
            mvEnv.Connection.ExecuteSQL(vSQL)
          End If
        Case "communications_log_doc_links"
          If pReply Then
            'link the original document to the copy
            vCLL = New CommunicationsLogLink
            vCLL.Create(CommunicationsLogLink.DocumentLinkTypes.dltDocumentToDocument, mvEnv, pOrigCommsLogNumber, pNewCommsLogNumber, 0, CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated, False, False)
            'link the copy to the original document
            vCLL = New CommunicationsLogLink
            vCLL.Create(CommunicationsLogLink.DocumentLinkTypes.dltDocumentToDocument, mvEnv, pNewCommsLogNumber, pOrigCommsLogNumber, 0, CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated, False, False)
          Else
            'copy all document links
            mvEnv.Connection.ExecuteSQL(vSQL)
          End If
        Case "communications_log_subjects"
          'copy all records
          mvEnv.Connection.ExecuteSQL(vSQL)
      End Select
    End Sub

    Public Sub SetCallTimes(ByVal pCallDuration As String, ByVal pTotalDuration As String)
      mvClassFields.Item(CommunicationsLogFields.CallDuration).Value = pCallDuration.Replace(":", "")
      mvClassFields.Item(CommunicationsLogFields.TotalDuration).Value = pTotalDuration.Replace(":", "")
    End Sub

    Public Function EMailDetailsDataTable(ByVal pAutoReply As Boolean, ByVal pAttachments As CDBDataTable) As CDBDataTable
      Dim vAttachmentRow As CDBDataRow
      Dim vAddressTo As String = ""
      Dim vMsgText As String
      Dim vAddDocument As Boolean
      Dim vAttachmentText As String = ""
      Dim vFileName As String
      Dim vOurRef As String
      Dim vHeaderText As String = ""
      Dim vCopyToList As String = ""
      Dim vContact As Contact

      Dim vSender As Contact = Nothing
      Dim vAddressee As Contact = Nothing
      Dim vAAddressNo As Integer
      Dim vSAddressNo As Integer
      GetSenderAddressee(vSender, vSAddressNo, vAddressee, vAAddressNo)
      vSender.SetAddress(vSAddressNo)
      vAddressee.SetAddress(vAAddressNo)
      Dim vTable As New CDBDataTable
      vTable.AddColumnsFromList("DocumentNumber,Subject,AddressTo,BodyText,CCList")
      Dim vRow As CDBDataRow = vTable.AddRow
      With vRow
        .Item(1) = CStr(CommunicationsLogNumber)
        If Len(Subject) > 0 Then
          .Item(2) = Subject
        Else
          .Item(2) = String.Format(ProjectText.String16705, vSender.LabelName) 'Document From %s
        End If
        'Build the initial To list
        Dim vEmailAddresses() As String
        If vAddressee.Existing Then
          vEmailAddresses = Split(vAddressee.EmailAddresses, ",")
          For Each vEMailAddress As String In vEmailAddresses
            If InStr(vEMailAddress, "@") > 0 Or pAutoReply Then
              If Len(vAddressTo) > 0 Then vAddressTo = vAddressTo & ","
              vAddressTo = vAddressTo & vEMailAddress
            End If
          Next
        End If
        'Get the contacts to whom this Comms Log was either copied or distributed
        Dim vCopyTo As New List(Of Contact)
        Dim vDistributedTo As New List(Of Contact)
        GetDistributionCopyList(vCopyTo, vDistributedTo)
        'Add the 'distributed to' contacts to the To list
        For Each vContact In vDistributedTo
          vEmailAddresses = Split(vContact.EmailAddresses, ",")
          For Each vEMailAddress As String In vEmailAddresses
            If InStr(vEMailAddress, "@") > 0 Or pAutoReply Then
              If Len(vAddressTo) > 0 Then vAddressTo = vAddressTo & ","
              vAddressTo = vAddressTo & vEMailAddress
            End If
          Next
        Next vContact
        'Add the 'copy to' contacts to the CC list
        For Each vContact In vCopyTo
          vEmailAddresses = Split(vContact.EmailAddresses, ",")
          For Each vEMailAddress As String In vEmailAddresses
            If InStr(vEMailAddress, "@") > 0 Or pAutoReply Then
              If Len(vCopyToList) > 0 Then vCopyToList = vCopyToList & ","
              vCopyToList = vCopyToList & vEMailAddress
            End If
          Next
        Next vContact
        .Item(3) = vAddressTo
        'Build the default email text first
        If vAddDocument Then
          vMsgText = ProjectText.String16723 'The attached document has been sent to you from:
        Else
          vMsgText = ProjectText.String16724 'This document has been sent to you from:
        End If
        vMsgText = vMsgText & vbCrLf & vbCrLf & vSender.NameAndAddress & vbCrLf & vbCrLf
        If Len(Precis) > 0 Then
          vMsgText = vMsgText & ProjectText.String16708 & vbCrLf & vbCrLf & Precis & vbCrLf & vbCrLf 'Document Precis:
        End If
        If pAttachments.Rows.Count() > 0 Then
          For Each vAttachmentRow In pAttachments.Rows
            If Len(vAttachmentText) > 0 Then vAttachmentText = vAttachmentText & vbCrLf
            vFileName = Left(vAttachmentRow.Item("Filename") & Space(20), 20)
            vOurRef = Left(vAttachmentRow.Item("OurReference") & Space(20), 20)
            vAttachmentText = vAttachmentText & vFileName
            vAttachmentText = vAttachmentText & vbTab & vOurRef
            vAttachmentText = vAttachmentText & vbTab & vAttachmentRow.Item("Subject")
          Next vAttachmentRow
          vMsgText = Space(pAttachments.Rows.Count() + 1) & vbCrLf & vMsgText
          'Build the attachment list header
          vHeaderText = ProjectText.String16725 & Space(10) & vbTab & ProjectText.String16726 & Space(11) & vbTab & ProjectText.String16727 & vbCrLf & "----------" & Space(10) & vbTab & "---------" & Space(11) & vbTab & "-------" 'Attachment    'Reference    'Subject
          vMsgText = vMsgText & vbCrLf & vbCrLf & vHeaderText & vbCrLf & vAttachmentText
        End If
        If mvEnv.GetControlBool(CDBEnvironment.cdbControlConstants.cdbControlEmailUseHeaderTemplate) Then
          vMsgText = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEmailHeaderTemplate)
          vMsgText = MultiLine(vMsgText)
          vMsgText = Replace(vMsgText, "%SenderLabelName", vSender.LabelName, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%SenderOrganisation", vSender.OrganisationName, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%SenderPosition", vSender.Position, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%SenderAddress", vSender.Address.AddressMultiLine, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%SenderNameAndAddress", vSender.NameAndAddress, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%UserName", mvEnv.User.FullName, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%UserPosition", mvEnv.User.Position, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%Precis", Precis, 1, -1, CompareMethod.Text)
          vMsgText = Replace(vMsgText, "%AttachmentText", vHeaderText & vbCrLf & vAttachmentText, 1, -1, CompareMethod.Text)
        End If
        .Item(4) = vMsgText
        .Item(5) = vCopyToList
      End With
      EMailDetailsDataTable = vTable
    End Function

    Private Sub GetSenderAddressee(ByRef pSender As Contact, ByRef pSenderAddressNumber As Integer, ByRef pAddressee As Contact, ByRef pAddresseeAddressNumber As Integer)
      pSender = New Contact(mvEnv)
      pAddressee = New Contact(mvEnv)
      pSenderAddressNumber = 0
      pAddresseeAddressNumber = 0
      pSender.Init()
      pAddressee.Init()
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & pSender.GetRecordSetFieldsName & ",link_type,cll.address_number FROM communications_log_links cll, contacts c WHERE communications_log_number = " & CommunicationsLogNumber & " AND link_type IN ('A','S') AND cll.contact_number = c.contact_number")
      While vRecordSet.Fetch
        If vRecordSet.Fields("link_type").Value = "A" Then
          pAddressee.InitFromRecordSetName(vRecordSet)
          pAddresseeAddressNumber = vRecordSet.Fields("address_number").LongValue
        ElseIf vRecordSet.Fields("link_type").Value = "S" Then
          pSender.InitFromRecordSetName(vRecordSet)
          pSenderAddressNumber = vRecordSet.Fields("address_number").LongValue
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Private Sub GetDistributionCopyList(ByVal pCopyTo As List(Of Contact), ByVal pDistributed As List(Of Contact))
      Dim vContact As New Contact(mvEnv)
      vContact.Init()
      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtName) & ",link_type,cll.address_number FROM communications_log_links cll, contacts c WHERE communications_log_number = " & CommunicationsLogNumber & " AND link_type IN ('" & CommunicationsLogLink.GetLinkTypeCode(CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied) & "','" & CommunicationsLogLink.GetLinkTypeCode(CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed) & "') AND cll.contact_number = c.contact_number")
      While vRecordSet.Fetch
        vContact = New Contact(mvEnv)
        If vRecordSet.Fields("link_type").Value = CommunicationsLogLink.GetLinkTypeCode(CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied) Then
          vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtName)
          pCopyTo.Add(vContact)
        ElseIf vRecordSet.Fields("link_type").Value = CommunicationsLogLink.GetLinkTypeCode(CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed) Then
          vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtName)
          pDistributed.Add(vContact)
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Overrides ReadOnly Property DataTable() As CDBDataTable
      Get
        'This function is only used by WEB Services at present
        'Please let me know if you want to change it (SDT)
        Dim vTable As New CDBDataTable
        Dim vCLS As New CommunicationsLogSubject
        Dim vSender As Contact = Nothing
        Dim vAddressee As Contact = Nothing
        Dim vAccessRights As New AccessRights
        Dim vRights As AccessRights.DocumentAccessRights
        Dim vStyle As DocumentStyles
        Dim vExternalApp As New ExternalApplication(mvEnv)

        vCLS.Init(mvEnv, CommunicationsLogNumber)
        Dim vAAddressNo As Integer
        Dim vSAddressNo As Integer
        GetSenderAddressee(vSender, vSAddressNo, vAddressee, vAAddressNo)
        'BR14489 to default missing Sender or Address following Data Import
        If vSender.ContactNumber = 0 OrElse vAddressee.ContactNumber = 0 Then
          Dim vAnonTCR As Boolean = False
          If vSender.ContactNumber = 0 AndAlso vAddressee.ContactNumber = 0 Then
            'Could be anonymous TCR (BR15326)
            If Direction = "I" AndAlso mvEnv.GetConfigOption("phone_allow_anonymous", False) = True _
            AndAlso DocumentType.Length > 0 AndAlso DocumentType = mvEnv.GetConfig("phone_in_document_type") Then
              vAnonTCR = True
            End If
          End If
          Dim vUser As New CDBUser(mvEnv)
          Dim vWhereFields As New CDBFields
          vWhereFields.Add("logname", CreatedBy)
          vUser.InitWithPrimaryKey(vWhereFields)
          If vSender.ContactNumber = 0 AndAlso vAnonTCR = False Then
            vSender.Init(vUser.ContactNumber)
            vSAddressNo = vSender.AddressNumber
          Else
            vAddressee.Init(vUser.ContactNumber)
            vAAddressNo = vAddressee.AddressNumber
          End If
        End If
        vAccessRights.Init(mvEnv)
        vRights = vAccessRights.GetDocumentRights(CommunicationsLogNumber)
        GetDocumentTypeInfo()
        Select Case mvDocumentSource
          Case ExternalApplication.DocumentSourceTypes.dstOther
            vStyle = DocumentStyles.dsnOther
          Case ExternalApplication.DocumentSourceTypes.dstScanner
            vStyle = DocumentStyles.dsnScannedImage
          Case ExternalApplication.DocumentSourceTypes.dstEmail
            If Len(StandardDocument) > 0 Then
              vStyle = DocumentStyles.dsnStandardEmailWithMerge
            Else
              vStyle = DocumentStyles.dsnBlankEmail
            End If
          Case Else
            If Len(StandardDocument) > 0 Then
              If mvMailMergeLetter Then
                vStyle = DocumentStyles.dsnStandardDocumentWithMerge
              ElseIf mvWordProcessorDocument Then
                vStyle = DocumentStyles.dsnStandardDocumentTemplate
              Else
                vStyle = DocumentStyles.dsnStandardDocumentPrecis
              End If
            Else
              If mvWordProcessorDocument Then
                vStyle = DocumentStyles.dsnBlankDocument             'Cannot distinguish if document was top and tailed
              Else
                vStyle = DocumentStyles.dsnPrecisOnly
              End If
            End If
        End Select
        vExternalApp.Init(ExternalApplicationCode)

        vTable.AddColumnsFromList("DocumentNumber,AddresseeContactNumber,AddresseeContactName,AddresseeAddressNumber")
        vTable.AddColumnsFromList("SenderContactNumber,SenderContactName,SenderAddressNumber,Dated")
        vTable.AddColumnsFromList("Direction,DocumentType,Topic,SubTopic,DocumentClass,DocumentSubject,Precis,StandardDocument")
        vTable.AddColumnsFromList("OurReference,Source,SourceDesc,CreatedBy,TheirReference,Department,Distributed,ExternalApplicationCode")
        vTable.AddColumnsFromList("AccessRights,DocumentStyle,ExternalApplicationType,ExternalApplicationExtension,DocumentName,Quantity")
        vTable.AddColumnsFromList("Received,Archiver,Recipient,Forwarded,Completed,HasLinkedDocuments")
        vTable.AddColumnsFromList("EmailTo,EmailFrom,EmailCc,EmailBcc,EmailReplyTo,OriginalUri")
        Dim vRow As CDBDataRow = vTable.AddRow
        With vRow
          .Item(1) = CStr(CommunicationsLogNumber)
          If vAddressee.ContactNumber > 0 Then
            .Item(2) = vAddressee.ContactNumber.ToString
          Else
            .Item(2) = ""
          End If
          .Item(3) = vAddressee.Name
          .Item(4) = vAAddressNo.ToString
          If vSender.ContactNumber > 0 Then
            .Item(5) = vSender.ContactNumber.ToString
          Else
            .Item(5) = ""
          End If
          .Item(6) = vSender.Name
          .Item(7) = vSAddressNo.ToString
          .Item(8) = Dated
          .Item(9) = Direction
          .Item(10) = DocumentType
          .Item(11) = vCLS.Topic
          .Item(12) = vCLS.SubTopic
          .Item(13) = DocumentClass
          .Item(14) = Subject
          .Item(15) = Precis
          .Item(16) = StandardDocument
          .Item(17) = OurReference
          .Item(18) = Source
          .Item(19) = mvEnv.GetDescription("sources", "source", Source)
          .Item(20) = CreatedBy
          .Item(21) = TheirReference
          .Item(22) = Department
          .Item(23) = Distributed
          .Item(24) = ExternalApplicationCode
          .Item(25) = CStr(vRights)
          .Item(26) = CStr(vStyle)
          .Item(27) = vExternalApp.CommunicationType.ToString
          .Item(28) = vExternalApp.Extension
          .Item(29) = Name
          If vCLS.QuantitySet Then
            .Item(30) = vCLS.Quantity.ToString
          Else
            .Item(30) = ""
          End If
          .Item(31) = Received
          .Item(32) = Archiver
          .Item(33) = Recipient
          .Item(34) = Forwarded
          .Item(35) = Completed
          .Item(36) = BooleanString(mvEnv.Connection.GetCount("communications_log_doc_links", New CDBFields(New CDBField("communications_log_number_1", CommunicationsLogNumber))) > 0)
          .Item(37) = EmailTo
          .Item(38) = EmailFrom
          .Item(39) = EmailCc
          .Item(40) = EmailBcc
          .Item(41) = EmailReplyTo
          .Item(42) = OriginalUri
        End With
        Return vTable
      End Get
    End Property

    Private Sub GetDocumentTypeInfo()
      If Not mvReadDocumentType Then
        mvReadDocumentType = True
        If DocumentType.Length > 0 Then
          Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT document_source, word_processor_document, mail_merge_letter FROM document_types WHERE document_type = '" & DocumentType & "'")
          With vRS
            If .Fetch() Then
              Select Case .Fields(1).Value
                Case "W"
                  mvDocumentSource = ExternalApplication.DocumentSourceTypes.dstWordProcessor
                Case "S"
                  mvDocumentSource = ExternalApplication.DocumentSourceTypes.dstScanner
                Case "O"
                  mvDocumentSource = ExternalApplication.DocumentSourceTypes.dstOther
                Case "E"
                  mvDocumentSource = ExternalApplication.DocumentSourceTypes.dstEmail
                Case Else
                  mvDocumentSource = ExternalApplication.DocumentSourceTypes.dstNoSource
              End Select
              mvWordProcessorDocument = .Fields(2).Bool
              mvMailMergeLetter = .Fields(3).Bool
            End If
            .CloseRecordSet()
          End With
        End If
      End If
    End Sub

  End Class

End Namespace
