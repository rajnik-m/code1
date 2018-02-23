Imports System.IO
Imports System.Runtime.Serialization

''' <summary>
''' Document informnation required to run a mailing fulfilment.
''' </summary>
Public Class DocumentInfo

  Private mvEnv As CDBEnvironment = Nothing

  ''' <summary>
  ''' Initializes a new instance of the <see cref="DocumentInfo"/> class.
  ''' </summary>
  ''' <param name="pEnv">The environment to use.</param>
  ''' <param name="pMailingTemplate">The mailing template to filter on.</param>
  ''' <param name="pCreatedBy">The creators log name to filter on.</param>
  ''' <param name="pCreatedOn">The date created to filter on.</param>
  ''' <param name="pContactNumber">The contact number to filter on.</param>
  ''' <param name="pBatchNumber">The batch number to filter on.</param>
  ''' <param name="pTransactionNumber">The transaction number to filter on.</param>
  ''' <remarks>When instanciated, the fulfulment history for this run.</remarks>
  Private Sub New(pEnv As CDBEnvironment, pMailingTemplate As String, pCreatedBy As String, pCreatedOn As String, pContactNumber As Integer, pBatchNumber As Integer, pTransactionNumber As Integer, pDataFileName As String)
    mvEnv = pEnv
    RequestedMailingTemplate = pMailingTemplate
    CreatedBy = pCreatedBy
    CreatedOn = pCreatedOn
    ContactNumber = pContactNumber
    BatchNumber = pBatchNumber
    TransactionNumber = pTransactionNumber
    If Not String.IsNullOrWhiteSpace(pDataFileName) Then
      Me.FulfillmentHistory.SetDataFileName(pDataFileName, True)
    End If
    Dim vDatatable As New CDBDataTable(mvEnv, New SQLStatement(mvEnv.Connection, "mt.standard_document,package,mailmerge_header," & (New ContactMailingDocument(mvEnv)).GetRecordSetFields, "contact_mailing_documents cmd", SelectionFilter, "cmd.mailing_template,mt.standard_document,selected_paragraphs,mailing_document_number", New AnsiJoins({New AnsiJoin("mailing_templates mt", "cmd.mailing_template", "mt.mailing_template"), New AnsiJoin("standard_documents sd", "mt.standard_document", "sd.standard_document")}), True))
    If vDatatable.Rows.Count > 0 Then
      DataRow = vDatatable.Rows(0)
      FulfillmentHistory.AddToDocumentList(Me.ContactMailingDocument)
      If vDatatable.Rows.Count > 1 Then
        Dim vLastStandardDocument As String = DataRow.Item("standard_document").ToString
        Dim vLastMailingTemplate As String = Me.ContactMailingDocument.MailingTemplateCode
        Dim vSelectedParagraphs As String = Me.ContactMailingDocument.SelectedParagraphs
        Dim vIndex As Integer = 1
        Dim vContactMailingDocument As ContactMailingDocument = GetCmdFromRow(vDatatable.Rows(vIndex))
        While vIndex < vDatatable.Rows.Count AndAlso vLastStandardDocument = vDatatable.Rows(vIndex).Item("standard_document").ToString AndAlso vLastMailingTemplate = vContactMailingDocument.MailingTemplateCode AndAlso SelectedParagraphs = vContactMailingDocument.SelectedParagraphs AndAlso vContactMailingDocument.SelectedParagraphs.Length <> 0 AndAlso FulfillmentHistory.NumberOfDocuments < FulfillmentHistory.DocumentsPerBatch
          vLastMailingTemplate = vContactMailingDocument.MailingTemplateCode
          vLastStandardDocument = vDatatable.Rows(vIndex).Item("standard_document").ToString
          vSelectedParagraphs = vContactMailingDocument.SelectedParagraphs
          FulfillmentHistory.AddToDocumentList(vContactMailingDocument)
          vIndex += 1
          If vIndex < vDatatable.Rows.Count Then
            vContactMailingDocument = GetCmdFromRow(vDatatable.Rows(vIndex))
          End If
        End While
      End If
    End If
  End Sub

  ''' <summary>
  ''' Get a new instance of the <see cref="DocumentInfo"/> class.
  ''' </summary>
  ''' <param name="pEnv">The environment to use.</param>
  ''' <param name="pMailingTemplate">The mailing template to filter on.</param>
  ''' <param name="pCreatedBy">The creators log name to filter on.</param>
  ''' <param name="pCreatedOn">The date created to filter on.</param>
  ''' <param name="pContactNumber">The contact number to filter on.</param>
  ''' <param name="pBatchNumber">The batch number to filter on.</param>
  ''' <param name="pTransactionNumber">The transaction number to filter on.</param>
  ''' <remarks>This method is used to get a new <see cref="DocumentInfo"/> object so that the object is only
  ''' created if appropriate for the parameters given.  If no appropriate data is found then no object
  ''' reference is returned.</remarks>
  Public Shared Function GetInstance(pEnv As CDBEnvironment, pMailingTemplate As String, pCreatedBy As String, pCreatedOn As String, pContactNumber As Integer, pBatchNumber As Integer, pTransactionNumber As Integer, Optional ByVal pDataFileName As String = "") As DocumentInfo
    Dim vResult As New DocumentInfo(pEnv, pMailingTemplate, pCreatedBy, pCreatedOn, pContactNumber, pBatchNumber, pTransactionNumber, pDataFileName)
    Return If(vResult.DataRow IsNot Nothing, vResult, Nothing)
  End Function

  ''' <summary>
  ''' Gets the selection filter.
  ''' </summary>
  ''' <value>
  ''' The selection filter.
  ''' </value>
  Private ReadOnly Property SelectionFilter() As CDBFields
    Get
      Dim vResult As New CDBFields({New CDBField("fulfillment_number", CDBField.FieldTypes.cftLong),
                                    New CDBField("earliest_fulfilment_date", "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket),
                                    New CDBField("earliest_fulfilment_date#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoLessThanEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)})
      If Not String.IsNullOrWhiteSpace(RequestedMailingTemplate) Then
        vResult.Add("cmd.mailing_template", RequestedMailingTemplate)
      End If
      If Not String.IsNullOrWhiteSpace(CreatedBy) Then
        vResult.Add("created_by", CreatedBy)
      End If
      If Not String.IsNullOrWhiteSpace(CreatedOn) Then
        vResult.Add("created_on", CDBField.FieldTypes.cftDate, CreatedOn)
      End If
      If mvContactNumber > 0 Then
        vResult.Add("contact_number", mvContactNumber.ToString)
      End If
      If mvBatchNumber > 0 Then
        vResult.Add("batch_number", mvBatchNumber.ToString)
      End If
      If mvTransactionNumber > 0 Then
        vResult.Add("transaction_number", mvTransactionNumber.ToString)
      End If
      If IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber)) > 0 Then
        vResult.Add("cmd.contact_number", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber), CDBField.FieldWhereOperators.fwoNotEqual)
      End If
      If String.IsNullOrWhiteSpace(RequestedMailingTemplate) Then
        vResult.Add("explicit_selection", "N")
      End If
      Return vResult
    End Get
  End Property

  Private mvRequestedMailingTemplate As String = String.Empty
  ''' <summary>
  ''' Gets the requested mailing template.
  ''' </summary>
  ''' <value>
  ''' The requested mailing template.
  ''' </value>
  Public Property RequestedMailingTemplate As String
    Get
      Return mvRequestedMailingTemplate
    End Get
    Private Set(pValue As String)
      mvRequestedMailingTemplate = pValue
    End Set
  End Property

  Private mvCreatedBy As String = String.Empty
  ''' <summary>
  ''' Gets the creator filter.
  ''' </summary>
  ''' <value>
  ''' The created filter.
  ''' </value>
  Public Property CreatedBy As String
    Get
      Return mvCreatedBy
    End Get
    Private Set(pValue As String)
      mvCreatedBy = pValue
    End Set
  End Property

  Private mvCreatedOn As String = String.Empty
  ''' <summary>
  ''' Gets the created on filter.
  ''' </summary>
  ''' <value>
  ''' The created on filter.
  ''' </value>
  Public Property CreatedOn As String
    Get
      Return mvCreatedOn
    End Get
    Private Set(pValue As String)
      mvCreatedOn = pValue
    End Set
  End Property

  Private mvContactNumber As Integer = 0
  ''' <summary>
  ''' Gets the contact number filter.
  ''' </summary>
  ''' <value>
  ''' The contact number filter.
  ''' </value>
  Public Property ContactNumber As Integer
    Get
      Return mvContactNumber
    End Get
    Private Set(pValue As Integer)
      mvContactNumber = pValue
    End Set
  End Property

  Private mvBatchNumber As Integer = 0
  ''' <summary>
  ''' Gets the batch number filter.
  ''' </summary>
  ''' <value>
  ''' The batch number filter.
  ''' </value>
  Public Property BatchNumber As Integer
    Get
      Return mvBatchNumber
    End Get
    Private Set(pValue As Integer)
      mvBatchNumber = pValue
    End Set
  End Property

  Private mvTransactionNumber As Integer = 0
  ''' <summary>
  ''' Gets the transaction number filter.
  ''' </summary>
  ''' <value>
  ''' The transaction number filter.
  ''' </value>
  Public Property TransactionNumber As Integer
    Get
      Return mvTransactionNumber
    End Get
    Private Set(pValue As Integer)
      mvTransactionNumber = pValue
    End Set
  End Property

  ''' <summary>
  ''' Gets the mailing template code.
  ''' </summary>
  ''' <value>
  ''' The mailing template code.
  ''' </value>
  Public ReadOnly Property MailingTemplateCode As String
    Get
      Return ContactMailingDocument.MailingTemplateCode
    End Get
  End Property

  ''' <summary>
  ''' Gets the standard document.
  ''' </summary>
  ''' <value>
  ''' The standard document.
  ''' </value>
  Public ReadOnly Property StandardDocument As String
    Get
      Return DataRow.Item("standard_document").ToString
    End Get
  End Property

  ''' <summary>
  ''' Gets the package.
  ''' </summary>
  ''' <value>
  ''' The package.
  ''' </value>
  Public ReadOnly Property Package As String
    Get
      Return DataRow.Item("package").ToString
    End Get
  End Property

  ''' <summary>
  ''' Gets the extenstion.
  ''' </summary>
  ''' <value>
  ''' The extenstion.
  ''' </value>
  Public ReadOnly Property Extension As String
    Get
      Return ExternalApplication.Extension
    End Get
  End Property

  ''' <summary>
  ''' Gets the selected paragraphs.
  ''' </summary>
  ''' <value>
  ''' The selected paragraphs.
  ''' </value>
  Public ReadOnly Property SelectedParagraphs As String
    Get
      Return ContactMailingDocument.SelectedParagraphs
    End Get
  End Property

  ''' <summary>
  ''' Gets the mailing document number.
  ''' </summary>
  ''' <value>
  ''' The mailing document number.
  ''' </value>
  Public ReadOnly Property MailingDocumentNumber As Integer
    Get
      Return ContactMailingDocument.MailingDocumentNumber
    End Get
  End Property

  ''' <summary>
  ''' Gets the standard documents.
  ''' </summary>
  ''' <value>
  ''' The standard documents.
  ''' </value>
  Public ReadOnly Property StandardDocuments As IList(Of String)
    Get
      Dim vResult As New List(Of String)
      vResult.Add(MailingTemplate.StandardDocumentCode)
      For Each vMTD As MailingTemplateDocument In MailingTemplate.Documents
        vResult.Add(vMTD.StandardDocumentCode)
      Next vMTD
      Return vResult.AsReadOnly
    End Get
  End Property

  ''' <summary>
  ''' Gets the printers.
  ''' </summary>
  ''' <value>
  ''' The printers.
  ''' </value>
  Public ReadOnly Property Printers As IList(Of Integer)
    Get
      Dim vResult As New List(Of Integer)
      vResult.Add(MailingTemplate.PrinterNumber)
      For Each vMTD As MailingTemplateDocument In MailingTemplate.Documents
        vResult.Add(vMTD.PrinterNumber)
      Next vMTD
      Return vResult.AsReadOnly
    End Get
  End Property

  ''' <summary>
  ''' Gets the mailmerge headers.
  ''' </summary>
  ''' <value>
  ''' The mailmerge headers.
  ''' </value>
  Public ReadOnly Property MailmergeHeaders As IList(Of String)
    Get
      Dim vResult As New List(Of String)
      vResult.Add(DataRow.Item("mailmerge_header").ToString)
      For Each vMTD As MailingTemplateDocument In MailingTemplate.Documents
        vResult.Add(vMTD.StandardDocument.MailmergeHeader)
      Next vMTD
      Return vResult.AsReadOnly
    End Get
  End Property

  ''' <summary>
  ''' Gets the book marks.
  ''' </summary>
  ''' <value>
  ''' The book marks.
  ''' </value>
  Public ReadOnly Property BookMarks As IList(Of String)
    Get
      Dim vResult As New List(Of String)
      If Not String.IsNullOrWhiteSpace(ContactMailingDocument.SelectedParagraphs) Then
        MailingTemplate.SetIncludedParagraphs(ContactMailingDocument.SelectedParagraphs)
        For Each vMTP As MailingTemplateParagraph In MailingTemplate.Paragraphs
          If Not vMTP.Include Then
            vResult.Add(vMTP.BookmarkName)
          End If
        Next
      End If
      Return vResult.AsReadOnly
    End Get
  End Property

  ''' <summary>
  ''' Gets the document list.
  ''' </summary>
  ''' <value>
  ''' The document list.
  ''' </value>
  Public ReadOnly Property DocumentList As String
    Get
      Return FulfillmentHistory.DocumentList
    End Get
  End Property

  ''' <summary>
  ''' Gets the report code.
  ''' </summary>
  ''' <value>
  ''' The report code.
  ''' </value>
  Public ReadOnly Property ReportCode As String
    Get
      Return MailingTemplate.StandardDocument.MailmergeHeader()
    End Get
  End Property

  ''' <summary>
  ''' Gets the first report parameter.
  ''' </summary>
  ''' <value>
  ''' The first report parameter.
  ''' </value>
  Public ReadOnly Property RP1 As String
    Get
      Return If(Not String.IsNullOrWhiteSpace(ReportCode), ContactMailingDocument.MailingTemplateCode, String.Empty)
    End Get
  End Property

  ''' <summary>
  ''' Gets the secomd report parameter.
  ''' </summary>
  ''' <value>
  ''' The second report parameter.
  ''' </value>
  Public ReadOnly Property RP2 As String
    Get
      Return CreatedBy
    End Get
  End Property

  ''' <summary>
  ''' Gets the third report parameter.
  ''' </summary>
  ''' <value>
  ''' The third report parameter.
  ''' </value>
  Public ReadOnly Property RP3 As String
    Get
      Return CreatedOn
    End Get
  End Property

  ''' <summary>
  ''' Gets the fourth report parameter.
  ''' </summary>
  ''' <value>
  ''' The fourth report parameter.
  ''' </value>
  Public ReadOnly Property RP4 As String
    Get
      Return If(ContactNumber > 0, ContactNumber.ToString, String.Empty)
    End Get
  End Property

  ''' <summary>
  ''' Gets the sixth report parameter.
  ''' </summary>
  ''' <value>
  ''' The sixth report parameter.
  ''' </value>
  Public ReadOnly Property RP6 As String
    Get
      Return mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMembershipSalesGroup)
    End Get
  End Property

  ''' <summary>
  ''' Gets the seventh report parameter.
  ''' </summary>
  ''' <value>
  ''' The seventh report parameter.
  ''' </value>
  Public ReadOnly Property RP7 As String
    Get
      Return If(BatchNumber > 0, BatchNumber.ToString, String.Empty)
    End Get
  End Property

  ''' <summary>
  ''' Gets the eight report parameter.
  ''' </summary>
  ''' <value>
  ''' The eigth report parameter.
  ''' </value>
  Public ReadOnly Property RP8 As String
    Get
      Return If(TransactionNumber > 0, TransactionNumber.ToString, String.Empty)
    End Get
  End Property

  ''' <summary>
  ''' Gets the ninth report parameter.
  ''' </summary>
  ''' <value>
  ''' The ninth report parameter.
  ''' </value>
  Public ReadOnly Property RP9 As String
    Get
      Return TodaysDate()
    End Get
  End Property

  ''' <summary>
  ''' Gets the tenth report parameter.
  ''' </summary>
  ''' <value>
  ''' The tenth report parameter.
  ''' </value>
  Public ReadOnly Property RP10 As String
    Get
      Return mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink)
    End Get
  End Property

  ''' <summary>
  ''' Gets the eleventh report parameter.
  ''' </summary>
  ''' <value>
  ''' The eleventh report parameter.
  ''' </value>
  Public ReadOnly Property RP11 As String
    Get
      Return mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink)
    End Get
  End Property

  ''' <summary>
  ''' Gets the twelfth report parameter.
  ''' </summary>
  ''' <value>
  ''' The twelfth report parameter.
  ''' </value>
  Public ReadOnly Property RP12 As String
    Get
      Return If(IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber)) > 0, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlHoldingContactNumber), String.Empty)
    End Get
  End Property

  ''' <summary>
  ''' Gets the name of the data file.
  ''' </summary>
  ''' <value>
  ''' The name of the data file.
  ''' </value>
  Public ReadOnly Property DataFileName As String
    Get
      Return FulfillmentHistory.DataFileName
    End Get
  End Property

  ''' <summary>
  ''' Gets the fulfillment number.
  ''' </summary>
  ''' <value>
  ''' The fulfillment number.
  ''' </value>
  Public ReadOnly Property FulfillmentNumber As Integer
    Get
      Return FulfillmentHistory.FulfillmentNumber
    End Get
  End Property

  Private mvDataRow As CDBDataRow = Nothing
  ''' <summary>
  ''' Gets the data and sets up the fulfilment history document list.
  ''' </summary>
  ''' <value>
  ''' The data row.
  ''' </value>
  Private Property DataRow As CDBDataRow
    Get
      Return mvDataRow
    End Get
    Set(pValue As CDBDataRow)
      mvDataRow = pValue
    End Set
  End Property

  Private Function GetCmdFromRow(pRow As CDBDataRow) As ContactMailingDocument
    Dim vResult As New ContactMailingDocument(mvEnv)
    vResult.InitFromCDBDataRow(pRow)
    Return vResult
  End Function

  Private mvContactMailingDocument As ContactMailingDocument = Nothing
  ''' <summary>
  ''' Gets the contact mailing document.
  ''' </summary>
  ''' <value>
  ''' The contact mailing document.
  ''' </value>
  Private ReadOnly Property ContactMailingDocument As ContactMailingDocument
    Get
      If mvContactMailingDocument Is Nothing Then
        mvContactMailingDocument = New ContactMailingDocument(mvEnv)
        mvContactMailingDocument.InitFromCDBDataRow(DataRow)
      End If
      Return mvContactMailingDocument
    End Get
  End Property

  Private mvExternalApplication As ExternalApplication = Nothing
  ''' <summary>
  ''' Gets the contact mailing document.
  ''' </summary>
  ''' <value>
  ''' The contact mailing document.
  ''' </value>
  Private ReadOnly Property ExternalApplication As ExternalApplication
    Get
      If mvExternalApplication Is Nothing Then
        mvExternalApplication = New ExternalApplication(mvEnv)
        mvExternalApplication.Init(Package)
      End If
      Return mvExternalApplication
    End Get
  End Property

  Private mvMailingTemplate As MailingTemplate = Nothing
  ''' <summary>
  ''' Gets the mailing template.
  ''' </summary>
  ''' <value>
  ''' The mailing template.
  ''' </value>
  Private ReadOnly Property MailingTemplate As MailingTemplate
    Get
      If mvMailingTemplate Is Nothing Then
        mvMailingTemplate = New MailingTemplate(mvEnv)
        mvMailingTemplate.Init(ContactMailingDocument.MailingTemplateCode)
      End If
      Return mvMailingTemplate
    End Get
  End Property

  Private mvFulfillmentHistory As FulfillmentHistory = Nothing
  ''' <summary>
  ''' Gets the fulfillment history.
  ''' </summary>
  ''' <value>
  ''' The fulfillment history.
  ''' </value>
  Private ReadOnly Property FulfillmentHistory As FulfillmentHistory
    Get
      If mvFulfillmentHistory Is Nothing Then
        mvFulfillmentHistory = New FulfillmentHistory(mvEnv)
        mvFulfillmentHistory.Init()
        mvFulfillmentHistory.SetControlNumber()
      End If
      Return mvFulfillmentHistory
    End Get
  End Property

End Class
