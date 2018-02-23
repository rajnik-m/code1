Namespace Access
  Public Class MailMergeInformation

    Private mvContactNumber As Integer
    Private mvAddressNumber As Integer
    Private mvOurReference As String
    Private mvTheirReference As String
    Private mvDated As String
    Private mvSenderFullName As String
    Private mvSenderPosition As String
    Private mvStandardDoc As StandardDocument
    Private mvMailMergeFileName As String
    Private mvRelatedContactNumber As Integer

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvContact As Contact

    Private Enum MergeFileTypes
      mftCSV = 1
      mftWordPerfect = 2
    End Enum

    Public Enum MailMergeTypes
      mmtWord
      mmtWordPerfect
    End Enum

    'Copied from SQLDATA.bas
    Enum MailMergeHeaderTypes
      'Mail merge header types
      mmhAdHocLetter = 1
      mmhGeneralMailing
      mmhInvoices
      mmhReceipts
      mmhPayPlan
      mmhCreditStatements
      mmhMailingProductionTransactions
      mmhProvisionalCashTransactions
      'Put all new mail merge headers above this line
      mmhCustom
    End Enum

    'Document sources
    Const DS_WORD_PROCESSOR As String = "W"

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Public Sub Init(ByRef pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pAddressNumber As Integer, ByRef pOurRef As String, ByRef pTheirRef As String, ByRef pDated As String, ByRef pSenderContactNumber As Integer, ByRef pStandardDocument As StandardDocument)
      Dim vOrg As Organisation
      Dim vFound As Boolean

      mvEnv = pEnv
      mvContactNumber = pContactNumber
      mvAddressNumber = pAddressNumber
      mvOurReference = pOurRef
      mvTheirReference = pTheirRef
      mvDated = pDated
      Dim vUser As New CDBUser(mvEnv)
      vUser.InitFromContactNumber(pSenderContactNumber)
      mvSenderFullName = vUser.FullName
      mvSenderPosition = vUser.Position
      'Select the record to get the required data
      mvContact = New Contact(mvEnv)
      mvContact.InitRecordSetType(mvEnv, Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtVAT Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtAddressCountry, mvContactNumber, mvAddressNumber)
      vFound = mvContact.Existing
      If Not mvContact.Existing Then
        'See if this is an Organisation at a non-default Address
        mvContact.Init(mvContactNumber)
        If mvContact.Existing = True And mvContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vOrg = New Organisation(mvEnv)
          vOrg.InitWithAddress(mvEnv, mvContactNumber, mvAddressNumber)
          If vOrg.Existing Then
            vFound = True
            mvContact.SetAddress(vOrg.Address.AddressNumber)
          End If
        End If
      End If
      If vFound = False Then RaiseError(DataAccessErrors.daeCannotFindContactAtAddress, CStr(pContactNumber), CStr(pAddressNumber))
      mvStandardDoc = pStandardDocument
    End Sub

    Public Function MailmergeReportNumber() As Integer
      Dim vSQL As String
      If Not mvStandardDoc Is Nothing Then
        If Len(mvStandardDoc.MailmergeHeader) > 0 And mvStandardDoc.MailmergeHeader <> "ADHOC" Then
          vSQL = "SELECT report_number FROM reports WHERE application_name = 'AD' and mailmerge_header = '" & mvStandardDoc.MailmergeHeader & "'"
          If mvEnv.ClientCode.Length > 0 Then
            vSQL = vSQL & " AND (client = '" & mvEnv.ClientCode & "' OR client IS NULL) "
            If mvEnv.Connection.NullsSortAtEnd Then
              vSQL = vSQL & "ORDER BY client"
            Else
              vSQL = vSQL & "ORDER BY client DESC"
            End If
          Else
            vSQL = vSQL & " AND client IS NULL"
          End If
          MailmergeReportNumber = IntegerValue(mvEnv.Connection.GetValue(vSQL))
        End If
      End If
    End Function

    Public Property MailMergeFileName() As String
      Get
        MailMergeFileName = mvMailMergeFileName
      End Get
      Set(ByVal Value As String)
        mvMailMergeFileName = Value
      End Set
    End Property

    Public Property RelatedContactNumber() As Integer
      Get
        RelatedContactNumber = mvRelatedContactNumber
      End Get
      Set(ByVal Value As Integer)
        mvRelatedContactNumber = Value
      End Set
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvContactNumber
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As Integer
      Get
        AddressNumber = mvAddressNumber
      End Get
    End Property
    Public ReadOnly Property OurReference() As String
      Get
        OurReference = mvOurReference
      End Get
    End Property
    Public ReadOnly Property TheirReference() As String
      Get
        TheirReference = mvTheirReference
      End Get
    End Property
    Public ReadOnly Property Dated() As String
      Get
        Dated = mvDated
      End Get
    End Property
    Public ReadOnly Property SenderFullName() As String
      Get
        SenderFullName = mvSenderFullName
      End Get
    End Property
    Public ReadOnly Property SenderPosition() As String
      Get
        SenderPosition = mvSenderPosition
      End Get
    End Property

    Public Function ExportMergeData(ByRef pExternalApplication As ExternalApplication, ByRef pMergeData As String, ByRef pWarning As Boolean, ByRef pWarningMsg As String, Optional ByRef pRelatedContactNumber As Integer = 0) As Integer
      'Export a data set for the current record out to a temporary file
      'The file will be formatted for the wordprocessor calling the routine.
      'The temporary directory is that dictated by the windows API call.
      'Note that at the moment the only two mailmerge options supported
      'are for WINWORD and WPWIN60 (word and wordperfect)
      'We only support one custom_merge_data record (although the header supports multiple)
      Dim vIndex As Integer
      Dim vHeader As String = ""
      Dim vExtRS As CDBRecordSet
      Dim vExtDataRS As CDBRecordSet
      Dim vExtSQL As String
      Dim vAttrs() As String
      Dim vExtCount As Integer
      Dim vDBName As String
      Dim vMergeFile As New DiskFile
      Dim vMMType As MailMergeTypes
      Dim vExtDone As Boolean
      Dim vValues() As String
      Dim vSequenceNo As Integer
      Dim vSuppress As Boolean
      Dim vRecordCount As Integer
      Dim vMultiLine As Boolean
      Dim vError As Integer

      If pExternalApplication.ClassName = "Microsoft Works" Then 'NoTranslate
        'because works doesn't, the user has to find the file so call it CDBMERGE.CVS
        'pMergeData = pExternalApplication.EXEPath & "CDBMERGE.CSV"
        'gvSystem.KillFile(pMergeData) 'just in case it exists from last time
        'System.Windows.Forms.Application.DoEvents()
      Else
        pMergeData = GetWindowsTempFileName((pExternalApplication.Extension), pMergeData) 'Set a temporary file for the merge data
      End If

      'Open pMergeData For Output As vFilenum
      With vMergeFile
        .OpenFile(pMergeData, DiskFile.FileOpenModes.fomOutput)
        Select Case pExternalApplication.MergeFileType
          Case MergeFileTypes.mftCSV
            'requires mergefile for word - chr$(34) = (") double quotes  - comma delimited
            vMMType = MailMergeTypes.mmtWord
            vHeader = GetMMHeader(MailMergeHeaderTypes.mmhAdHocLetter, "ADHOC", True)
            .PrintLine_Renamed(vHeader)
          Case MergeFileTypes.mftWordPerfect
            vMMType = MailMergeTypes.mmtWordPerfect
            'Create a Word Perfect 5 format file
            'First the header (FF) (WPC) (32 00 00 00 pointer to data 4 bytes) (01 product type) (0A file type) (00 00 Encryption) (00 00 Reserved)
            vHeader = Chr(255) & "WPC" & Chr(50) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(10) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            'Now the index header block (FB FF Index marker) (05 00 Number of index headers) (32 00 Size of index block) (00 00 00 00 Position of next index)
            vHeader = vHeader & Chr(251) & Chr(255) & Chr(5) & Chr(0) & Chr(50) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            'Four index blocks (FF FF Deleted Index marker) (00 00 00 00 Size of index block) (00 00 00 00 Position of index)
            vHeader = vHeader & Chr(255) & Chr(255) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            vHeader = vHeader & Chr(255) & Chr(255) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            vHeader = vHeader & Chr(255) & Chr(255) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            vHeader = vHeader & Chr(255) & Chr(255) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
            .PrintString(vHeader)
          Case Else
            'Unknown Merge File Type so use Comma Delimited
        End Select

        If mvContact.Existing And mvContact.Address.Existing Then
          .PrintMailMergeItem(mvOurReference, vMMType, True)
          .PrintMailMergeItem(mvTheirReference, vMMType)
          .PrintMailMergeItem(mvDated, vMMType)
          .PrintMailMergeItem(CStr(mvContact.ContactNumber), vMMType)
          If mvContact.ContactType <> Contact.ContactTypes.ctcOrganisation Then
            .PrintMailMergeItem((mvContact.LabelName), vMMType)
            .PrintMailMergeItem((mvContact.TitleName), vMMType)
            .PrintMailMergeItem((mvContact.Initials), vMMType)
            .PrintMailMergeItem((mvContact.Forenames), vMMType)
            .PrintMailMergeItem((mvContact.Surname), vMMType)
            .PrintMailMergeItem((mvContact.PrefixHonorifics), vMMType)
            .PrintMailMergeItem((mvContact.Honorifics), vMMType)
          Else
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
            .PrintMailMergeItem("", vMMType)
          End If
          .PrintMailMergeItem((mvContact.Salutation), vMMType)
          .PrintMailMergeItem((mvContact.InformalSalutation), vMMType)
          .PrintMailMergeItem((mvContact.Position), vMMType)
          .PrintMailMergeItem((mvContact.OrganisationName), vMMType)
          .PrintMailMergeItem((mvContact.PositionLocation), vMMType)
          .PrintMailMergeItem(mvContact.Address.MergeAddressLine1, vMMType)
          .PrintMailMergeItem(mvContact.Address.MergeAddressLine2, vMMType)
          .PrintMailMergeItem(mvContact.Address.MergeAddressLine3, vMMType)
          .PrintMailMergeItem(mvContact.Address.Town, vMMType)
          .PrintMailMergeItem(mvContact.Address.County, vMMType)
          .PrintMailMergeItem(mvContact.Address.Postcode, vMMType)
          .PrintMailMergeItem(mvContact.Address.BuildingNumber, vMMType)
          .PrintMailMergeItem(mvContact.Address.Sortcode, vMMType)
          .PrintMailMergeItem(mvContact.Address.NonDefaultCountryDescription, vMMType)
          If InStr(vHeader, "Country Code") > 0 Or InStr(vHeader, "Country_Code") > 0 Then
            .PrintMailMergeItem(mvContact.Address.Country, vMMType)
          End If
          .PrintMailMergeItem(mvSenderFullName, vMMType)
          .PrintMailMergeItem(mvSenderPosition, vMMType)

          'Find if any external data is to be added
          If Len(mvEnv.ClientCode) > 0 And mvEnv.GetConfigOption("option_custom_data", False) Then
            vExtRS = mvEnv.Connection.GetRecordSet("SELECT * FROM custom_merge_data WHERE client = '" & mvEnv.ClientCode & "' AND usage_code = 'D' ORDER BY sequence_number")
            vExtRS.Fetch()
            While vExtRS.Status()
              vDBName = vExtRS.Fields("db_name").Value
              vSequenceNo = vExtRS.Fields("sequence_number").IntegerValue
              vMultiLine = vExtRS.Fields.FieldExists("multi_line").Bool
              vAttrs = vExtRS.Fields("attribute_names").Value.Split(","c)
              vExtCount = vAttrs.Length
              ReDim vValues(vExtCount)
              vExtSQL = vExtRS.Fields("select_sql").Value
              vExtSQL = ReplaceString(vExtSQL, "?", CStr(mvContactNumber))
              vExtSQL = ReplaceString(vExtSQL, "#R", CStr(pRelatedContactNumber))
              vExtSQL = ReplaceString(vExtSQL, "#", CStr(mvContactNumber))
              vExtDataRS = mvEnv.GetConnection(vDBName).GetRecordSet(vExtSQL)
              vRecordCount = 0
              vExtDone = False
              While vExtDataRS.Fetch() = True And Not vExtDone
                For vIndex = 1 To vExtCount
                  If vRecordCount = 0 Then
                    vValues(vIndex) = vExtDataRS.Fields(vIndex).Value
                  Else
                    vValues(vIndex) = vValues(vIndex) & vbCrLf & vExtDataRS.Fields(vIndex).Value
                  End If
                Next
                vRecordCount = vRecordCount + 1
                If Not vMultiLine Then vExtDone = True
              End While
              vExtDataRS.CloseRecordSet()
              vSuppress = False
              If vExtRS.Fetch() = True Then
                'Is the next one the same sequence number
                If vExtRS.Fields("sequence_number").IntegerValue = vSequenceNo Then
                  If vRecordCount = 0 Then
                    vSuppress = True
                  Else
                    vExtRS.Fetch()
                  End If
                End If
              End If
              If Not vSuppress Then
                For vIndex = 1 To vExtCount
                  .PrintMailMergeItem(vValues(vIndex), vMMType)
                Next
              End If
              pWarningMsg = String.Format(ProjectText.String19214, vDBName) 'Custom Merge Data refers to an invalid Database (%s)\r\n\r\nThe data will be incomplete
              pWarning = True
            End While
            vExtRS.CloseRecordSet()
          End If
          .PrintNewLine()
        End If
        .CloseFile()
      End With
      ExportMergeData = vError
    End Function

    Public Function GetMMHeader(ByVal pType As MailMergeHeaderTypes, ByVal pMMHeaderCode As String, ByVal pAddQuotes As Boolean) As String
      Dim vHeader As String = ""
      Dim vExtHeader As String
      Dim vLastNo As Integer
      Dim vMMHeader As New MailmergeHeader(mvEnv)
      Dim vRecordSet As CDBRecordSet
      Dim vItems() As String
      Dim vIndex As Integer

      If pMMHeaderCode.Length > 0 Then
        vMMHeader.Init(pMMHeaderCode)
        If vMMHeader.Existing Then vHeader = vMMHeader.MailmergeFields
      End If

      If Len(mvEnv.ClientCode) > 0 And mvEnv.GetConfigOption("option_custom_data", False) And pType = MailMergeHeaderTypes.mmhAdHocLetter Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT sequence_number,attribute_captions FROM custom_merge_data WHERE client = '" & mvEnv.ClientCode & "' AND usage_code = 'D' ORDER BY sequence_number")
        vLastNo = -1
        While vRecordSet.Fetch() = True
          If vRecordSet.Fields(1).IntegerValue <> vLastNo Then
            vLastNo = vRecordSet.Fields(1).IntegerValue
            vExtHeader = vRecordSet.Fields(2).Value
            vItems = Split(vExtHeader, ",")
            For vIndex = 0 To UBound(vItems)
              vItems(vIndex) = Trim(vItems(vIndex))
            Next
            vExtHeader = Join(vItems, ",")
            If vExtHeader.Length > 0 Then vHeader = vHeader & "," & vExtHeader
          End If
        End While
        vRecordSet.CloseRecordSet()
      End If

      If pAddQuotes Then
        'Include quotes around all fields
        vHeader = Chr(34) & vHeader & Chr(34)
        vHeader = Replace(vHeader, ",", Chr(34) & "," & Chr(34))
      End If
      Return vHeader
    End Function
    Public Function GetMergeFile(ByRef pEnv As CDBEnvironment, ByRef pMailmergeHeader As String, ByRef pExtension As String) As String
      Dim vHeaderType As MailMergeHeaderTypes
      Dim vFileNum As Integer
      Dim vHeader As String
      Dim vMergeFileName As String
      Dim vIndex As Integer

      mvEnv = pEnv
      If pExtension.Length > 0 Then
        vMergeFileName = GetWindowsTempFileName(pExtension)
      Else
        vMergeFileName = GetWindowsTempFileName(".txt")
      End If
      vFileNum = FreeFile()
      FileOpen(vFileNum, vMergeFileName, OpenMode.Output)
      Select Case pMailmergeHeader
        Case "INV"
          vHeaderType = MailMergeHeaderTypes.mmhInvoices
        Case "RECPT"
          vHeaderType = MailMergeHeaderTypes.mmhReceipts
        Case "PPLAN"
          vHeaderType = MailMergeHeaderTypes.mmhPayPlan
        Case "CSTAT"
          vHeaderType = MailMergeHeaderTypes.mmhCreditStatements
        Case "MPTRAN", "MPTRPP", "MPTEAM"
          vHeaderType = MailMergeHeaderTypes.mmhMailingProductionTransactions
        Case "PRVCSH"
          vHeaderType = MailMergeHeaderTypes.mmhProvisionalCashTransactions
        Case "GMMM"
          vHeaderType = MailMergeHeaderTypes.mmhGeneralMailing
        Case "ADHOC"
          vHeaderType = MailMergeHeaderTypes.mmhAdHocLetter
        Case Else
          If Len(pMailmergeHeader) > 0 Then
            vHeaderType = MailMergeHeaderTypes.mmhCustom
          Else
            pMailmergeHeader = "ADHOC"
            vHeaderType = MailMergeHeaderTypes.mmhAdHocLetter
          End If
      End Select
      vHeader = GetMMHeader(vHeaderType, pMailmergeHeader, False)
      vHeader = vHeader & vbCrLf
      For vIndex = 1 To Len(vHeader)
        If Mid(vHeader, vIndex, 1) = "," Then vHeader = vHeader & ","
      Next
      vHeader = vHeader & vbCrLf 'Word wont accept the merge file unless a data line is included
      PrintLine(vFileNum, vHeader)
      FileClose(vFileNum)
      GetMergeFile = vMergeFileName
    End Function

    Public Sub New(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub
  End Class
End Namespace
