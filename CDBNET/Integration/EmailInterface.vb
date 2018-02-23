Imports System.IO

Public MustInherit Class EmailInterface

  'TODO Add option to save only attachments
  'TODO Handle deletion of document links
  'TODO Handle adding document links
  'TODO Handle drag drop of documents

  Public Enum SendEmailOptions
    seoNoOptions = 0
    seoAddressResolveUI = 1
    seoCloseFormAfterEMail = 2
    seoShowAddressBook = 4
    seoAlwaysEditMail = 8
    seoReceiptRequired = 16
    seoMultipleRecipients = 32
    seoForceAddressBook = 64
  End Enum

  Public Enum EMailActions
    emaNone
    emaReply
    emaReplyAll
    emaForward
    emaDelete
    emaSave
  End Enum

  Protected mvRecipientToList As ArrayList
  Protected mvRecipientCCList As ArrayList
  Private mvInBox As ArrayList
  Private mvMessageTable As DataTable
  Private mvUseSecurityManager As Boolean
  Private mvSecurityManagerChecked As Boolean

  Public MustOverride Function CanEMail(Optional ByVal pDownloadMail As Boolean = False) As Boolean
  Public MustOverride Function SendMail(ByVal pForm As Form, ByVal pOptions As SendEmailOptions, ByVal pSubject As String, ByVal pMessage As String, ByVal pEmailAddress As String, Optional ByVal pAttachments As Collection = Nothing, Optional ByVal pCCList As String = "") As Boolean
  Public MustOverride Function GetAttachmentPathName(ByVal pMsgID As String, ByVal pIndex As Integer) As String
  Public MustOverride Function GetInBox() As ArrayList
  Public MustOverride Sub MarkRead(ByVal pMsg As EMailMessage)
  Public MustOverride Function ProcessAction(ByVal pMsg As EMailMessage, ByVal pAction As EMailActions) As Boolean

  Public Function GetAttachmentPackage(ByVal pMessage As EMailMessage, ByVal pIndex As Integer) As String
    Dim vName As String = pMessage.AttachmentCollection(pIndex).ToString
    Dim vFileInfo As New FileInfo(vName)
    Dim vTable As DataTable = DataHelper.GetCachedLookupData(CareServices.XMLLookupDataTypes.xldtPackages)
    For Each vRow As DataRow In vTable.Rows
      If vFileInfo.Extension.ToUpper = vRow.Item("DocfileExtension").ToString.ToUpper Then
        'TODO Handle internal or external storage
        Return vRow.Item("Package").ToString
      End If
    Next
    Return ""
  End Function
  Public Function ProcessAction(ByVal pID As String, ByVal pAction As EMailActions) As Boolean
    Dim vMsg As EMailMessage = EmailMessageByID(pID)
    If Not vMsg Is Nothing Then Return ProcessAction(vMsg, pAction)
  End Function
  Public Sub ShowAttachment(ByVal pID As String, ByVal pIndex As Integer)
    Dim vFileName As String = GetAttachmentPathName(pID, pIndex)
    Dim vFileInfo As New FileInfo(vFileName)
    Dim vApplication As ExternalApplication = FormHelper.GetDocumentApplication(vFileInfo.Extension)
    vApplication.ViewDocument(vFileInfo)
  End Sub
  Public Function EmailMessageByID(ByVal pID As String) As EMailMessage
    Dim vIndex As Integer

    For vIndex = 0 To mvInBox.Count - 1
      If pID = CType(mvInBox(vIndex), EMailMessage).ID Then
        Return CType(mvInBox(vIndex), EMailMessage)
      End If
    Next
    Return Nothing
  End Function
  Public Function GetInBoxData() As DataSet
    Dim vRow As DataRow
    Dim vDataSet As New DataSet

    Dim vMsgTable As New DataTable("DataRow")
    vMsgTable.Columns.AddRange(New DataColumn() _
    { _
      New DataColumn("ID"), _
      New DataColumn("Read"), _
      New DataColumn("Attachments"), _
      New DataColumn("From"), _
      New DataColumn("Subject"), _
      New DataColumn("Received", Type.GetType("System.DateTime")), _
      New DataColumn("To"), _
      New DataColumn("CC") _
    })
    mvInBox = GetInBox()
    For Each vEMailMessage As EMailMessage In mvInBox
      vRow = vMsgTable.NewRow
      With vEMailMessage
        Dim vString() As String = {.ID, .Read.ToString, .AttachmentCount.ToString, .OrigDisplayName, .Subject, .DateReceived, .ToList, .CCList}
        vRow.ItemArray = vString
      End With
      vMsgTable.Rows.Add(vRow)
    Next
    Dim vPrimaryColumns(0) As DataColumn
    vPrimaryColumns(0) = vMsgTable.Columns("ID")
    vMsgTable.PrimaryKey = vPrimaryColumns

    vDataSet.Tables.Add(vMsgTable)
    vMsgTable.DefaultView.Sort = "Received DESC"
    mvMessageTable = vMsgTable

    Dim vColTable As DataTable = New DataTable("Column")
    vColTable.Columns.AddRange(New DataColumn() _
    { _
      New DataColumn("Name"), _
      New DataColumn("Visible"), _
      New DataColumn("Heading"), _
      New DataColumn("DataType"), _
      New DataColumn("Width") _
    })
    Dim vWidths() As String = {"100", "100", "25", "180", "450", "0", "0", "0"}
    For Each vCol As DataColumn In vMsgTable.Columns
      vRow = vColTable.NewRow
      vRow.Item(0) = vCol.ColumnName
      If vCol.ColumnName = "ID" Or vCol.ColumnName = "Read" Then
        vRow.Item(1) = "N"
      Else
        vRow.Item(1) = "Y"
      End If
      If vCol.ColumnName = "Attachments" Then
        vRow.Item(2) = ""
      Else
        vRow.Item(2) = vCol.ColumnName
      End If
      If vCol.ColumnName = "Received" Then
        vRow.Item(3) = "DateTime"
      Else
        vRow.Item(3) = "Char"
      End If
      vRow.Item(4) = vWidths(vColTable.Rows.Count)
      vColTable.Rows.Add(vRow)
    Next
    vDataSet.Tables.Add(vColTable)
    Return vDataSet
  End Function
  Public Property LastRecipientCCList() As System.Collections.ArrayList
    Get
      Return mvRecipientCCList
    End Get
    Set(ByVal Value As System.Collections.ArrayList)
      mvRecipientCCList = Value
    End Set
  End Property
  Public Property LastRecipientToList() As System.Collections.ArrayList
    Get
      Return mvRecipientToList
    End Get
    Set(ByVal Value As System.Collections.ArrayList)
      mvRecipientToList = Value
    End Set
  End Property
  Protected Property UseOutlookSecurityManager() As Boolean
    Get
      If Not mvSecurityManagerChecked Then
        mvUseSecurityManager = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.email_security_manager)
        mvSecurityManagerChecked = True
      End If
      Return mvUseSecurityManager
    End Get
    Set(ByVal pValue As Boolean)
      mvUseSecurityManager = pValue
    End Set
  End Property
  Protected ReadOnly Property UserEMailLogname() As String
    Get
      Return DataHelper.UserInfo.EMailLogin
    End Get
  End Property

End Class
