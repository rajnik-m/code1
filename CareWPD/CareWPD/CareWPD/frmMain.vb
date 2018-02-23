Imports System.IO
Imports CDBNETCL.CareWebAccess
Imports CDBNETCL.Utilities

Public Class frmMain

  Private mvDataType As XMLWebDataSelectionTypes
  Private mvWebItem As WebItem
  Private mvEditing As Boolean
  Private mvWebNumber As Integer
  Private mvWebName As String
  Private mvWebURL As String
  Private mvLoginPage As Integer
  Private mvDefaultPage As Integer
  Private mvBaseUrl As String
  Private mvLastPageTab As Integer
  Private mvContextItemNumber As Integer
  Private mvImageTable As DataTable
  Private mvDocumentTable As DataTable
  Private mvInitialSettings As String = ""
  Private mvDefaultSettings As String = ""
  Private mvDashBoardDataSource As DashboardDataSource

  Private Const INVALID_WEB_NUMBER As Integer = 9999

  Private WithEvents mvPopupMenu As New WebPopupMenu
  Private WithEvents mvBrowserMenu As New BrowserPopupMenu

  Public Sub New(ByVal pWebNumber As Integer)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pWebNumber)
    Me.TopMost = False
  End Sub

  Private Sub InitialiseControls(ByVal pWebNumber As Integer)
    AppHelper.CurrentMainForm = Me
    splMaint.Panel1Collapsed = True
    If pWebNumber > 0 Then
      mvWebNumber = pWebNumber
    Else
      mvWebNumber = INVALID_WEB_NUMBER
    End If
    Dim vList As New ParameterList(True)
    vList.IntegerValue("WebNumber") = mvWebNumber
    sel.Init(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtWebSelectionPages, vList), XMLWebDataSelectionTypes.wstControl, mvWebNumber)
    SetCaption(sel.Caption)
    If mvWebNumber > 0 AndAlso mvWebNumber <> INVALID_WEB_NUMBER Then sel.ContextMenuStrip = mvPopupMenu
    splTop.Panel1Collapsed = True         'No Header
    mvParentForm = Nothing
    mvSelectedRow = -1
    cmdDelete.Visible = False
    cmdNew.Visible = False
    cmdClose.Visible = False
    cmdDefault.Visible = False
    HideGrid()
    cmdSave.Enabled = False       'For the moment
    Me.KeyPreview = True
  End Sub

  Private Sub SetCaption(ByVal pWebName As String)
    Me.Text = String.Format("Web Page Designer for: {0}", pWebName)
  End Sub

  Private Sub RefreshCard()
    Dim vList As New ParameterList(True)
    Select Case mvDataType
      Case XMLWebDataSelectionTypes.wstControl
        vList.IntegerValue("WebNumber") = mvWebItem.ID
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebInformation
      Case XMLWebDataSelectionTypes.wstMenu
        vList.IntegerValue("WebMenuNumber") = mvWebItem.ID
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebMenu
      Case XMLWebDataSelectionTypes.wstMenuItem
        vList.IntegerValue("WebMenuItemNumber") = mvWebItem.ID
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebMenuItem
      Case XMLWebDataSelectionTypes.wstPage
        vList.IntegerValue("WebPageNumber") = mvWebItem.ID
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebPage
      Case XMLWebDataSelectionTypes.wstPageItem
        vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebPageItem
      Case XMLWebDataSelectionTypes.wstImageItem
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebImage
      Case XMLWebDataSelectionTypes.wstDocumentItem
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebImage
      Case Else
        mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone
        cmdSave.Visible = False
        cmdDelete.Visible = False
    End Select
    cmdLink1.Visible = False
    cmdLink2.Visible = False
    cmdLink3.Visible = False
    cmdLink4.Visible = False
    cmdOther.Visible = False
    mvPopupMenu.SetContext(mvDataType, mvWebItem.ID)
    If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctNone Then
      epl.Visible = False
    Else
      cmdSave.Visible = True
      epl.Visible = False
      epl.Init(New EditPanelInfo(mvMaintenanceType, Nothing))

      If mvDataType = XMLWebDataSelectionTypes.wstPage AndAlso mvLastPageTab > 0 Then epl.TabSelectedIndex = mvLastPageTab
      epl.Visible = True
      If mvEditing Then
        Dim vDataRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetWebData(mvDataType, vList))
        epl.Populate(vDataRow)
        If mvDataType = XMLWebDataSelectionTypes.wstControl Then
          Debug.Print("Setting Web Information")
          mvWebName = epl.GetValue("WebName")
          SetCaption(mvWebName)
          mvWebURL = vDataRow("WebUrl").ToString
          mvLoginPage = IntegerValue(vDataRow("LoginPageNumber").ToString)
          mvDefaultPage = IntegerValue(vDataRow("WebPageNumber").ToString)
          mvBaseUrl = mvWebURL & "/Default.aspx"
          mvWebItem.HeaderItemNumber = IntegerValue(vDataRow("HeaderItemNumber").ToString)
          mvWebItem.FooterItemNumber = IntegerValue(vDataRow("FooterItemNumber").ToString)
          mvWebItem.HeaderHtml = vDataRow("HeaderHtml").ToString
          mvWebItem.FooterHtml = vDataRow("FooterHtml").ToString
          mvWebItem.LeftPanelHtml = vDataRow("LeftPanelHtml").ToString
          mvWebItem.RightPanelHtml = vDataRow("RightPanelHtml").ToString
          cmdDelete.Visible = False
          cmdLink1.Text = "Edit Header"
          cmdLink1.Visible = True
          cmdLink2.Text = "Edit Footer"
          cmdLink2.Tag = ""
          cmdLink2.Visible = True
          If mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebInformation Then
            cmdOther.Visible = True
            cmdOther.Text = "Help"
          End If
          'If AppValues.RunAsThames Then
          cmdLink3.Text = "Edit Left Panel"
          cmdLink3.Visible = True
          cmdLink4.Text = "Edit Right Panel"
          cmdLink4.Visible = True
          'End If
          SetImages()
          SetDocuments()
        Else
          cmdDelete.Visible = True
        End If
        If mvDataType = XMLWebDataSelectionTypes.wstPage Then
          Dim vContactInfo As ContactInfo = DataHelper.UserContactInfo()
          Dim vContactParms As String = String.Format("&cwpd={0}&ucn={1}&uan={2}", "Y", vContactInfo.ContactNumber, vContactInfo.AddressNumber)
          epl.SetValue("WebBrowser", String.Format("{0}?pn={1}{2}", mvBaseUrl, mvWebItem.ID, vContactParms))
          epl.SetContextMenu("WebBrowser", mvBrowserMenu)
          If epl.FindPanelControl("AccessViewName", False) IsNot Nothing AndAlso epl.FindPanelControl("LoginTypeRequired", False) IsNot Nothing Then
            epl.EnableControl("AccessViewName", epl.GetValue("LoginTypeRequired") = "R")
          End If
        End If
        If mvDataType = XMLWebDataSelectionTypes.wstPageItem Then
          Dim vControl As String = epl.GetValue("WebPageUserControl")

          If IsContentControl(vControl) Then
            cmdLink1.Text = "Edit Data"
            cmdLink1.Visible = True
          Else
            Select Case vControl
              Case "DISPLAYCONTACTDATA", "PRODUCTSELECTION", "BOOKINGOPTIONSELECTION", "EVENTSELECTION", "EVENTDELEGATESELECTION", _
                "PICKSESSIONS", "MEMBERSHIPTYPESELECTION", "UPDATEPOSITION", "SHOWEVENTBOOKINGS", "SURVEYSELECTION", "CONTACTCPDCYCLE", "DOWNLOADSELECTION", "VIEWTRANSACTION", _
                "DISPLAYRELATEDTEDCONTACTS", "SEARCHCONTACT", "DISPLAYRELATEDTEDORGANISATIONS", "UPDATECPDPOINTS", "UPDATECPDOBJECTIVES", "SEARCHDIRECTORY", _
                "DISPLAYTRANSACTIONS", "UPDATEADDRESS", "UPDATEPHONENUMBER", "UPDATEEMAILADDRESS", "PAYMULTIPLEPAYMENTPLANS", "EXAMSELECTION", "PAYERSELECTION", _
                "REGISTERCORPORATEMEMBER", "REGISTERCOMBINED", "DEDUPLICATEORGANISATIONS", "SHOWEXAMBOOKINGS", "SHOWEXAMHISTORY", "SETUSERORGANISATION", "INVOICEPAYMENT", "SELECTPAYPLANFORDD"
                mvInitialSettings = vDataRow("InitialParameters").ToString
                mvDefaultSettings = vDataRow("DefaultParameters").ToString
                cmdLink1.Text = "Parameters"
                cmdLink1.Visible = True
                cmdLink2.Text = "Customise"
                cmdLink2.Visible = True
                cmdLink2.Tag = "Customisation"
              Case "DOWNLOADDOCUMENT"
                mvInitialSettings = vDataRow("InitialParameters").ToString
                mvDefaultSettings = vDataRow("DefaultParameters").ToString
                cmdLink1.Visible = False
              Case Else
                mvInitialSettings = vDataRow("InitialParameters").ToString
                mvDefaultSettings = vDataRow("DefaultParameters").ToString
                cmdLink1.Text = "Parameters"
                cmdLink1.Visible = True
            End Select
            PopulateParentGroup(vDataRow.Item("ParentGroupName").ToString)
          End If
          SetWebPageControlItems(vControl)
        End If
      Else
        If mvDataType = XMLWebDataSelectionTypes.wstImageItem Then
          epl.SetValue("WebBrowser", mvWebURL & "/images/" & mvWebItem.FileName)
          cmdSave.Visible = False
        ElseIf mvDataType = XMLWebDataSelectionTypes.wstDocumentItem Then
          epl.SetValue("WebBrowser", mvWebURL & "/documents/" & mvWebItem.FileName)
          cmdSave.Visible = False
        End If
        SetDefaults()
        cmdDelete.Visible = False
      End If
    End If



    cmdSave.Enabled = True
    If epl.TabSelectedIndex <= 0 Then epl.Invalidate()
    bpl.RepositionButtons()
  End Sub

  Private Sub WebTabSelected(ByVal pSender As Object, ByVal pType As XMLWebDataSelectionTypes, ByVal pItem As WebItem) Handles sel.WebTabSelected
    Dim vCursor As New BusyCursor
    Try
      If mvDataType = XMLWebDataSelectionTypes.wstPage Then mvLastPageTab = epl.TabSelectedIndex
      mvDataType = pType
      mvWebItem = pItem
      mvEditing = pItem.ID > 0
      RefreshCard()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub RefreshTabSelector()
    Dim vList As New ParameterList(True)
    vList.IntegerValue("WebNumber") = mvWebNumber
    sel.Init(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtWebSelectionPages, vList), mvDataType, mvWebItem.ID)
    SetCaption(sel.Caption)
    SetImages()
    SetDocuments()
    If mvWebNumber > 0 AndAlso mvWebNumber <> INVALID_WEB_NUMBER Then sel.ContextMenuStrip = mvPopupMenu
  End Sub

  Protected Overrides Sub cmdLink_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim vOriginalDST As String = ""
    Try
      If sender Is cmdLink1 OrElse (sender Is cmdLink2 AndAlso (cmdLink2.Tag Is Nothing OrElse cmdLink2.Tag.ToString.Length = 0)) OrElse sender Is cmdLink3 OrElse sender Is cmdLink4 Then
        vOriginalDST = GetParameterValueFromList(mvInitialSettings, "DataSelectionType")
        If ProcessSave(False, sender) Then
          Dim vList As New ParameterList(True)
          Dim vHTML As String
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctWebInformation
              Dim vForm As New frmHTMLEditor
              Dim vHtmlItem As String
              vList.IntegerValue("WebNumber") = mvWebItem.ID
              If sender Is cmdLink3 Then
                vHTML = mvWebItem.LeftPanelHtml
                vHtmlItem = "LeftPanelHtml"
              ElseIf sender Is cmdLink4 Then
                vHTML = mvWebItem.RightPanelHtml
                vHtmlItem = "RightPanelHtml"
              ElseIf sender Is cmdLink2 Then
                vHTML = mvWebItem.FooterHtml
                vHtmlItem = "FooterHtml"
              Else
                vHTML = mvWebItem.HeaderHtml
                vHtmlItem = "HeaderHtml"
              End If
              If EditHtml(vHTML) = System.Windows.Forms.DialogResult.OK Then
                vList(vHtmlItem) = vHTML
                DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctWebInformation, vList)
                If sender Is cmdLink3 Then
                  mvWebItem.LeftPanelHtml = vHTML
                ElseIf sender Is cmdLink4 Then
                  mvWebItem.RightPanelHtml = vHTML
                ElseIf sender Is cmdLink1 Then
                  mvWebItem.HeaderHtml = vHTML
                ElseIf sender Is cmdLink2 Then
                  mvWebItem.FooterHtml = vHTML
                End If
              End If
            Case CareServices.XMLMaintenanceControlTypes.xmctWebPageItem
              Dim vUserControl As String = epl.GetValue("WebPageUserControl")
              If IsContentControl(vUserControl) Then
                EditWebPageItemData(mvWebItem.ID)
              Else

                If GetParametersForUserControl(vUserControl, mvInitialSettings, mvDefaultSettings) Then
                  vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
                  vList("InitialParameters") = mvInitialSettings
                  vList("DefaultParameters") = mvDefaultSettings
                  mvDashBoardDataSource = New DashboardDataSource
                  If vUserControl = "DISPLAYCONTACTDATA" AndAlso vOriginalDST <> GetParameterValueFromList(mvInitialSettings, "DataSelectionType") Then
                    'display list type has changed, so delete any record in display_list_items
                    vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
                    DataHelper.DeleteCustomisedDisplayList(vList)
                  End If
                  DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctWebPageItem, vList)
                End If
              End If
          End Select
        End If
      ElseIf sender Is cmdOther Then
        Select Case mvMaintenanceType
          Case CareServices.XMLMaintenanceControlTypes.xmctWebInformation
            Dim vForm As New frmBrowser(DataHelper.HelpURL("index.htm", False), False, True)
            vForm.Show()
        End Select
      Else
        Dim vList As New ParameterList
        Dim vFinderType As Boolean
        Dim vTransactionType As Boolean
        Dim vEventSelectionType As Boolean
        Dim vControl As String = epl.GetValue("WebPageUserControl")
        Select Case vControl
          Case "DISPLAYCONTACTDATA"
            vList.FillFromValueList(mvInitialSettings)
          Case "UPDATEADDRESS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses).ToString
          Case "UPDATEPHONENUMBER"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers).ToString
          Case "UPDATEEMAILADDRESS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers).ToString
          Case "MEMBERSHIPTYPESELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebMembershipTypes).ToString
          Case "PRODUCTSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebProducts).ToString
          Case "BOOKINGOPTIONSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebBookingOptions).ToString
          Case "EVENTSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebEvents).ToString
          Case "EXAMSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebExams).ToString
          Case "VIEWTRANSACTION"
            vTransactionType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis).ToString
          Case "EVENTDELEGATESELECTION", "EVENTDELEGATEENTRY"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventBookingDelegates).ToString
          Case "PICKSESSIONS"
            vEventSelectionType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions).ToString
          Case "SHOWEVENTBOOKINGS"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebEventBookings).ToString
          Case "SURVEYSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebSurveys).ToString
          Case "UPDATEPOSITION"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions).ToString
          Case "CONTACTCPDCYCLE"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDCycles).ToString
          Case "UPDATECPDPOINTS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPoints).ToString
          Case "UPDATECPDOBJECTIVES"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDObjectives).ToString
          Case "DOWNLOADSELECTION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebDocuments).ToString
          Case "DISPLAYRELATEDTEDCONTACTS"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebRelatedContacts).ToString
          Case "SEARCHCONTACT"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebContacts).ToString
          Case "DISPLAYRELATEDTEDORGANISATIONS"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftRelatedOrganisations).ToString
          Case "SEARCHDIRECTORY"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebDirectoryEntries).ToString
          Case "DISPLAYTRANSACTIONS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactProcessedTransactions).ToString
          Case "PAYMULTIPLEPAYMENTPLANS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlanPayments).ToString
          Case "REGISTERCORPORATEMEMBER", "REGISTERCOMBINED"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebMemberOrganisations).ToString
          Case "DEDUPLICATEORGANISATIONS"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtDuplicateOrganisationsForRegistration).ToString
          Case "SHOWEXAMBOOKINGS"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebExamBookings).ToString
          Case "SHOWEXAMHISTORY"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLDataFinderTypes.xdftWebExamHistory).ToString
          Case "SETUSERORGANISATION"
            vFinderType = True
            vList("DataSelectionType") = CInt(CareNetServices.XMLTableDataSelectionTypes.xtdstViewData).ToString
          Case "INVOICEPAYMENT"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactOutstandingInvoices).ToString
          Case "SELECTPAYPLANFORDD"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlans).ToString
          Case "PAYERSELECTION"
            vList("DataSelectionType") = CInt(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesAndPositions).ToString

        End Select
        If vFinderType Then
          Dim vIndex As CareServices.XMLDataFinderTypes
          vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
          mvDashBoardDataSource = New DashboardDataSource
          vIndex = CType(IntegerValue(vList("DataSelectionType").ToString), CareServices.XMLDataFinderTypes)
          mvDashBoardDataSource.SetDataType(vIndex)
        ElseIf vTransactionType Then
          Dim vIndex As CareServices.XMLTransactionDataSelectionTypes
          vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
          mvDashBoardDataSource = New DashboardDataSource
          vIndex = CType(IntegerValue(vList("DataSelectionType").ToString), CareServices.XMLTransactionDataSelectionTypes)
          mvDashBoardDataSource.SetDataType(vIndex)
        ElseIf vEventSelectionType Then
          Dim vIndex As CareServices.XMLEventDataSelectionTypes
          vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
          mvDashBoardDataSource = New DashboardDataSource
          vIndex = CType(IntegerValue(vList("DataSelectionType").ToString), CareServices.XMLEventDataSelectionTypes)
          mvDashBoardDataSource.SetDataType(vIndex)
        Else
          Dim vIndex As CareServices.XMLContactDataSelectionTypes
          vList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
          mvDashBoardDataSource = New DashboardDataSource
          vIndex = CType(IntegerValue(vList("DataSelectionType").ToString), CareServices.XMLContactDataSelectionTypes)
          mvDashBoardDataSource.SetDataType(vIndex)
        End If
        Dim vFrmDL As New frmDisplayList(frmDisplayList.ListUsages.DisplayListMaintenance, mvDashBoardDataSource, vList)
        vFrmDL.ShowDialog(Me)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overrides Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      If ProcessSave(False, sender) Then

      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Protected Overrides Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ProcessDelete()
  End Sub

  Private Function EditSiteData(ByVal pParameterName As String) As Boolean
    Dim vList As New ParameterList(True)
    vList.IntegerValue("WebNumber") = mvWebNumber
    Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetWebData(XMLWebDataSelectionTypes.wstControl, vList))
    Dim vHTML As String = ""
    Dim vHtmlItem As String
    If pParameterName = "SiteHeader" Then
      vHtmlItem = "HeaderHtml"
    Else
      vHtmlItem = "FooterHtml"
    End If
    vHTML = vRow(vHtmlItem).ToString
    If EditHtml(vHTML) = System.Windows.Forms.DialogResult.OK Then
      vList(vHtmlItem) = vHTML
      DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctWebInformation, vList)
      Return True
    End If
  End Function

  Private Function EditWebPageItemData(ByVal pWebPageItemNumber As Integer) As Boolean
    Dim vList As New ParameterList(True)
    vList.IntegerValue("WebPageItemNumber") = pWebPageItemNumber
    Dim vHTML As String = ""
    Try
      Dim vRow As DataRow = DataHelper.GetRowFromDataSet(DataHelper.GetWebData(XMLWebDataSelectionTypes.wstPageData, vList))
      vHTML = vRow("PageHtml").ToString
    Catch vException As CareException
      'No data for this item
    End Try
    If EditHtml(vHTML) = System.Windows.Forms.DialogResult.OK Then
      vList("WebPageHtml") = vHTML
      DataHelper.UpdateWebPageData(vList)
      Return True
    End If
  End Function

  Private Function EditHtml(ByRef pHtml As String) As DialogResult
    Dim vForm As New frmHTMLEditor
    vForm.ImageTable = mvImageTable
    vForm.DocumentTable = GetDocumentsTable()
    vForm.PageTable = GetPagesTable()
    vForm.BaseImageUrl = mvWebURL & "/Images/"

    'The following code uses Replace$ as it can be made case insensitive whereas String.Replace is not
    vForm.HTMLText = Replace$(pHtml, " src=""images/", " src=""" & mvWebURL & "/Images/", 1, -1, CompareMethod.Text)
    If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
      pHtml = Replace$(vForm.HTMLText, " src=""about:", " src=""", 1, -1, CompareMethod.Text)
      pHtml = Replace$(pHtml, String.Format(" src=""{0}/Images/", mvWebURL), " src=""Images/", 1, -1, CompareMethod.Text)
      pHtml = Replace$(pHtml, String.Format("{0}/", mvWebURL), "", 1, -1, CompareMethod.Text)
      If pHtml Is Nothing Then pHtml = String.Empty
      Return System.Windows.Forms.DialogResult.OK
    Else
      Return System.Windows.Forms.DialogResult.Cancel
    End If
  End Function

  Private Function GetDocumentsTable() As DataTable
    Dim vTable As DataTable = New DataTable
    Dim vColumn As New DataColumn("DocumentName")
    vTable.Columns.Add(vColumn)
    vColumn = New DataColumn("DocumentURL")
    vTable.Columns.Add(vColumn)
    Dim vNewRow As DataRow
    If mvDocumentTable IsNot Nothing Then
      For Each vRow As DataRow In mvDocumentTable.Rows
        vNewRow = vTable.NewRow()
        vNewRow("DocumentName") = vRow(0).ToString
        vNewRow("DocumentURL") = String.Format("{0}/documents/{1}", mvWebURL, vRow(0).ToString)
        vTable.Rows.Add(vNewRow)
      Next
    End If
    Return vTable
  End Function

  Private Function GetPagesTable() As DataTable
    Dim vList As New ParameterList(True)
    vList.IntegerValue("WebNumber") = mvWebNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetWebData(XMLWebDataSelectionTypes.wstPages, vList))
    Dim vColumn As New DataColumn("PageURL")
    vTable.Columns.Add(vColumn)
    For Each vRow As DataRow In vTable.Rows
      vRow("PageURL") = String.Format("{0}/Default.aspx?pn={1}", mvWebURL, vRow("WebPageNumber").ToString)
    Next
    Return vTable
  End Function

  Protected Sub GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As ParameterList) Handles epl.GetCodeRestrictions
    Select Case pParameterName
      Case "WebNumber"
        pList.IntegerValue(pParameterName) = mvWebNumber
      Case "WebPageNumber"
        pList.IntegerValue(pParameterName) = mvWebItem.PageNumber
    End Select
  End Sub

  Private Sub BeforeSelect(ByVal pSender As Object, ByRef pCancel As Boolean) Handles sel.BeforeSelect
    Dim vChangeNode As Boolean = True
    If epl.DataChanged Then
      If ConfirmCancel() = False Then
        pCancel = True
        vChangeNode = False
      Else
        'We have some data changed and we are going to cancel it - Check if it was a new appeal or segment
        If mvEditing = False Then
          epl.DataChanged = False
          'sel.RemoveSelectedNode()
        End If
      End If
    End If
    If vChangeNode AndAlso _
      (mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebPage _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebPageItem _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebMenu _
       Or mvMaintenanceType = CareServices.XMLMaintenanceControlTypes.xmctWebMenuItem) Then
      If Not mvEditing Then sel.RemoveSelectedNode()
    End If
  End Sub

  Private Sub GetAdditionalKeyValues(ByVal pList As ParameterList)
    Select Case mvDataType
      Case XMLWebDataSelectionTypes.wstPage, XMLWebDataSelectionTypes.wstMenu
        pList.IntegerValue("WebNumber") = mvWebNumber
      Case XMLWebDataSelectionTypes.wstPageItem
        pList.IntegerValue("WebPageNumber") = mvWebItem.PageNumber
        If mvWebItem.SequenceNumber > 0 Then pList.IntegerValue("SequenceNumber") = mvWebItem.SequenceNumber
      Case XMLWebDataSelectionTypes.wstMenuItem
        pList.IntegerValue("WebMenuNumber") = mvWebItem.MenuNumber
        pList.IntegerValue("ParentItemNumber") = mvWebItem.ParentID
        If mvWebItem.SequenceNumber > 0 Then pList.IntegerValue("SequenceNumber") = mvWebItem.SequenceNumber
    End Select
  End Sub

  Private Sub GetPrimaryKeyValues(ByVal pList As ParameterList, ByVal pRow As Integer, ByVal pForUpdate As Boolean)
    Select Case mvDataType
      Case XMLWebDataSelectionTypes.wstControl
        pList.IntegerValue("WebNumber") = mvWebNumber
      Case XMLWebDataSelectionTypes.wstMenu
        pList.IntegerValue("WebMenuNumber") = mvWebItem.ID
      Case XMLWebDataSelectionTypes.wstMenuItem
        pList.IntegerValue("WebMenuItemNumber") = mvWebItem.ID
      Case XMLWebDataSelectionTypes.wstPage
        pList.IntegerValue("WebPageNumber") = mvWebItem.ID
      Case XMLWebDataSelectionTypes.wstPageItem
        pList.IntegerValue("WebPageItemNumber") = mvWebItem.ID
    End Select
  End Sub

  Protected Overrides Function ProcessSave(ByVal pDefault As Boolean, ByVal sender As System.Object) As Boolean 'Return true if saved
    Try
      Dim vList As New ParameterList(True)
      If mvEditing Then
        'If editing an existing record then get the primary key values
        GetPrimaryKeyValues(vList, mvSelectedRow, True)
      Else
        'For new records add in any additional key values
        GetAdditionalKeyValues(vList)
      End If
      If TypeOf sender Is WebPopupMenu.WebMenuItems Then
        Select Case DirectCast(sender, WebPopupMenu.WebMenuItems)
          Case WebPopupMenu.WebMenuItems.wmiMoveItemUp
            If mvWebItem.SequenceNumber > 0 Then vList.IntegerValue("SequenceNumber") = mvWebItem.SequenceNumber - 1
          Case WebPopupMenu.WebMenuItems.wmiMoveItemDown
            If mvWebItem.SequenceNumber > 0 Then vList.IntegerValue("SequenceNumber") = mvWebItem.SequenceNumber + 1
        End Select
      End If
      If epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll) Then
        'Update or Insert record
        If vList.ContainsKey("NumberOfRows") AndAlso vList("NumberOfRows") = "" Then vList.Remove("NumberOfRows")
        If mvMaintenanceType = CareNetServices.XMLMaintenanceControlTypes.xmctWebPage Then
          Dim vDataSet As DataSet
          Dim vSelectData As New ParameterList(True)
          If epl.Controls.Find("FriendlyUrl", True).Length > 0 AndAlso epl.GetValue("FriendlyUrl").Length > 0 Then
            If Not epl.GetValue("FriendlyUrl").EndsWith(".aspx") Then
              epl.SetErrorField("FriendlyUrl", "Friendly URL should end with .aspx")
              Return False
            End If
            vSelectData("WebNumber") = mvWebNumber.ToString
            vSelectData("FriendlyUrl") = vList("FriendlyUrl")
            vDataSet = DataHelper.GetWebData(XMLWebDataSelectionTypes.wstPages, vSelectData)
            If vDataSet.Tables("DataRow") IsNot Nothing AndAlso vDataSet.Tables("DataRow").Rows.Count > 0 Then
              If epl.GetValue("WebPageNumber").Length = 0 OrElse epl.GetValue("WebPageNumber") <> vDataSet.Tables("DataRow").Rows(0).Item("WebPageNumber").ToString Then
                ShowInformationMessage(String.Format("The Record specified by the Friendly URL already exists for {0} page", vDataSet.Tables("DataRow").Rows(0).Item("WebPageName").ToString))
                Return False
              End If
            End If
          End If
        End If
        If mvEditing Then
          mvReturnList = DataHelper.UpdateItem(mvMaintenanceType, vList)
        Else
          Select Case mvMaintenanceType
            Case CareServices.XMLMaintenanceControlTypes.xmctWebInformation
              mvWebName = vList("WebName")
              mvWebURL = vList("WebUrl")
              If GetImages() = False Then Return False
            Case CareServices.XMLMaintenanceControlTypes.xmctWebPageItem
              If Not IsContentControl(vList("WebPageUserControl")) Then
                Dim vInitialSettings As String = ""
                Dim vDefaultSettings As String = ""
                If GetParametersForUserControl(vList("WebPageUserControl"), vInitialSettings, vDefaultSettings) = False Then Return False
                vList("InitialParameters") = vInitialSettings
                vList("DefaultParameters") = vDefaultSettings
              End If
          End Select
          mvReturnList = DataHelper.AddItem(mvMaintenanceType, vList)
        End If
        mvRefreshParent = True
        epl.DataChanged = False     'Data saved now
        If mvDataType = XMLWebDataSelectionTypes.wstControl Then
          If Not mvEditing Then
            mvWebItem.ID = mvReturnList.IntegerValue("WebNumber")
            mvWebNumber = mvWebItem.ID
          End If
          UpdateWebPortalInfo(mvWebNumber, vList("WebName").ToString)
          RefreshTabSelector()
        Else
          If mvEditing Then
            Select Case mvDataType
              Case XMLWebDataSelectionTypes.wstMenu
                sel.UpdateSelectedNodeText(vList("WebMenuName"))
              Case XMLWebDataSelectionTypes.wstMenuItem
                sel.UpdateSelectedNodeText(vList("MenuTitle"))
              Case XMLWebDataSelectionTypes.wstPage
                sel.UpdateSelectedNodeText(vList("WebPageName"))
              Case XMLWebDataSelectionTypes.wstPageItem
                sel.UpdateSelectedNodeText(vList("WebPageItemName"))
            End Select
          Else
            Select Case mvDataType
              Case XMLWebDataSelectionTypes.wstMenu
                mvWebItem.ID = mvReturnList.IntegerValue("WebMenuNumber")
              Case XMLWebDataSelectionTypes.wstMenuItem
                mvWebItem.ID = mvReturnList.IntegerValue("WebMenuItemNumber")
              Case XMLWebDataSelectionTypes.wstPage
                mvWebItem.ID = mvReturnList.IntegerValue("WebPageNumber")
              Case XMLWebDataSelectionTypes.wstPageItem
                mvWebItem.ID = mvReturnList.IntegerValue("WebPageItemNumber")
            End Select
            RefreshTabSelector()
          End If
        End If
        Return True
      End If
    Catch vEx As CareException
      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
        Case CareException.ErrorNumbers.enSpecifiedDataNotFound
          ShowInformationMessage(vEx.Message)
        Case Else
          Throw vEx
      End Select
    End Try
  End Function

  Private Sub mvPopupMenu_MenuSelected(ByVal pItem As WebPopupMenu.WebMenuItems) Handles mvPopupMenu.MenuSelected
    Try
      Select Case pItem
        Case WebPopupMenu.WebMenuItems.wmiImport
          Dim vOFD As New OpenFileDialog
          With vOFD
            .InitialDirectory = AppValues.ConfigurationValue(AppValues.ConfigurationValues.default_mailing_directory)
            .Title = "Import File Name"
            .Filter = "XML Files (*.xml)|*.xml|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .CheckFileExists = True
            .DefaultExt = "xml"
            .FileName = "WebExport.xml"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
              Dim vCursor As New BusyCursor
              Dim vDataSet As DataSet = New DataSet
              vDataSet.ReadXml(.FileName)
              Dim vID As Integer
              For Each vTable As DataTable In vDataSet.Tables
                For Each vCol As DataColumn In vTable.Columns
                  Select Case vCol.ColumnName
                    Case "WebNumber"
                      vTable.Rows(0)(vCol) = mvWebNumber
                    Case "WebName"
                      vTable.Rows(0)(vCol) = mvWebName
                    Case "WebUrl"
                      vTable.Rows(0)(vCol) = mvWebURL
                    Case "WebMenuNumber", "WebMenuItemNumber", "WebPageNumber", "WebPageItemNumber", "WebPageItemLinkNumber", _
                       "LoginPageNumber", "SubmitItemNumber", "ParentItemNumber", "HeaderItemNumber", "FooterItemNumber", "FpApplication", "LeftPanelItemNumber", "RightPanelItemNumber"
                      For Each vRow As DataRow In vTable.Rows
                        vID = IntegerValue(vRow(vCol).ToString)
                        If vID > 0 Then
                          vID = vID Mod 100000
                          vID += (mvWebNumber * 100000)
                          vRow(vCol) = vID
                        End If
                      Next
                  End Select
                Next
              Next
              Dim vList As New ParameterList(True)
              vList.IntegerValue("WebNumber") = mvWebNumber
              DataHelper.ImportWebData(vList, vDataSet.GetXml)
              mvImageTable = Nothing
              mvDocumentTable = Nothing
              RefreshTabSelector()
              vCursor.Dispose()
              ShowInformationMessage(InformationMessages.ImImportWebComplete)
            End If
          End With
        Case WebPopupMenu.WebMenuItems.wmiExport
          Dim vOFD As New SaveFileDialog
          With vOFD
            .InitialDirectory = AppValues.ConfigurationValue(AppValues.ConfigurationValues.default_mailing_directory)
            .Title = "Export File Name"
            .Filter = "XML Files (*.xml)|*.xml|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .CheckPathExists = True
            .OverwritePrompt = True
            .DefaultExt = "xml"
            .FileName = "WebExport.xml"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
              Dim vCursor As New BusyCursor
              Dim vList As New ParameterList(True)
              vList.IntegerValue("WebNumber") = mvWebNumber
              Dim vDataSet As DataSet = DataHelper.GetWebData(XMLWebDataSelectionTypes.wstExport, vList)
              vDataSet.WriteXml(.FileName)
              vCursor.Dispose()
              ShowInformationMessage(InformationMessages.ImExportWebComplete)
            End If
          End With
        Case WebPopupMenu.WebMenuItems.wmiRefresh
          mvImageTable = Nothing
          mvDocumentTable = Nothing
          RefreshTabSelector()
        Case WebPopupMenu.WebMenuItems.wmiNewImage
          AddNewImage()
        Case WebPopupMenu.WebMenuItems.wmiNewDocument
          AddNewDocument()
        Case WebPopupMenu.WebMenuItems.wmiNewPage
          sel.AddNode(XMLWebDataSelectionTypes.wstPage, mvWebItem)
        Case WebPopupMenu.WebMenuItems.wmiNewPageItem
          sel.AddNode(XMLWebDataSelectionTypes.wstPageItem, mvWebItem)
        Case WebPopupMenu.WebMenuItems.wmiCopyPage
          Dim vList As New ParameterList(True)
          vList.IntegerValue("WebNumber") = mvWebNumber
          vList.IntegerValue("WebPageNumber") = mvWebItem.ID
          vList("WebPageName") = "Copy"
          vList("AsCopy") = "Y"
          DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctWebPage, vList)
          RefreshTabSelector()
        Case WebPopupMenu.WebMenuItems.wmiNewMenu
          sel.AddNode(XMLWebDataSelectionTypes.wstMenu, mvWebItem)
        Case WebPopupMenu.WebMenuItems.wmiNewMenuItem
          sel.AddNode(XMLWebDataSelectionTypes.wstMenuItem, mvWebItem)
        Case WebPopupMenu.WebMenuItems.wmiDeletePage, WebPopupMenu.WebMenuItems.wmiDeletePageItem, _
             WebPopupMenu.WebMenuItems.wmiDeleteMenu, WebPopupMenu.WebMenuItems.wmiDeleteMenuItem
          ProcessDelete()
        Case WebPopupMenu.WebMenuItems.wmiAddPage
          AddPageToMenuItem()
        Case WebPopupMenu.WebMenuItems.wmiMoveItemUp
          ProcessSave(False, pItem)
          RefreshTabSelector()
        Case WebPopupMenu.WebMenuItems.wmiMoveItemDown
          ProcessSave(False, pItem)
          RefreshTabSelector()
      End Select
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub mvBrowserMenu_MenuSelected(ByVal pItem As BrowserPopupMenu.BrowserMenuItems) Handles mvBrowserMenu.MenuSelected
    Dim vList As New ParameterList(True)
    Dim vItemType As String = ""
    Select Case pItem
      Case BrowserPopupMenu.BrowserMenuItems.bmiAddLeftContent
        vItemType = "L"
      Case BrowserPopupMenu.BrowserMenuItems.bmiAddCenterContent
        vItemType = "C"
      Case BrowserPopupMenu.BrowserMenuItems.bmiAddRightContent
        vItemType = "R"
      Case BrowserPopupMenu.BrowserMenuItems.bmiAddFullWidthContent
        vItemType = "W"
      Case BrowserPopupMenu.BrowserMenuItems.bmiDeletePageItem
        If mvContextItemNumber > 0 Then
          If ConfirmDelete() = False Then Exit Sub
          vList.IntegerValue("WebPageItemNumber") = mvContextItemNumber
          DataHelper.DeleteItem(CareServices.XMLMaintenanceControlTypes.xmctWebPageItem, vList)
        End If
      Case BrowserPopupMenu.BrowserMenuItems.bmiRefresh
    End Select
    If vItemType.Length > 0 Then AddContentToPage(vItemType, "DATADISPLAY")
    mvContextItemNumber = 0
    RefreshTabSelector()
  End Sub

  Private Sub mvBrowserMenu_ModuleMenuSelected(ByVal pItem As BrowserPopupMenu.BrowserMenuItems, ByVal pControl As String) Handles mvBrowserMenu.ModuleMenuSelected
    AddContentToPage("C", pControl)
    mvContextItemNumber = 0
    RefreshTabSelector()
  End Sub

  Private Sub AddPageToMenuItem()
    If IntegerValue(epl.GetValue("WebPageNumber")) > 0 Then
      ShowInformationMessage("This Menu Item already has a Web Page defined against it")
    Else
      Dim vList As New ParameterList(True)
      vList.IntegerValue("WebNumber") = mvWebNumber
      vList("WebPageName") = epl.GetValue("MenuTitle")
      vList("WebPageTitle") = epl.GetValue("MenuDesc")
      Dim vReturnList As ParameterList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctWebPage, vList)
      'Repopulate web page combobox
      Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetWebData(XMLWebDataSelectionTypes.wstPages, vList))
      If vTable IsNot Nothing Then epl.SetComboDataSource("WebPageNumber", "WebPageNumber", "WebPageName", vTable, True)
      epl.SetValue("WebPageNumber", vReturnList("WebPageNumber"))
      ProcessSave(False, Me)
      RefreshTabSelector()
    End If
  End Sub

  Private Sub AddContentToPage(ByVal pItemType As String, ByVal pUserControl As String)
    If pItemType.Length > 0 Then
      Dim vContentControl As Boolean = IsContentControl(pUserControl)
      Dim vInitialSettings As String = ""
      Dim vDefaultSettings As String = ""
      If vContentControl OrElse GetParametersForUserControl(pUserControl, vInitialSettings, vDefaultSettings) Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("WebPageNumber") = mvWebItem.ID
        vList("WebPageItemType") = pItemType
        vList("WebPageUserControl") = pUserControl
        vList("InitialParameters") = vInitialSettings
        vList("DefaultParameters") = vDefaultSettings
        If mvDefaultPage > 0 AndAlso vContentControl = False Then vList.IntegerValue("SubmitItemNumber") = mvDefaultPage
        Dim vReturnList As ParameterList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctWebPageItem, vList)
        If vContentControl Then
          vList("WebPageItemNumber") = vReturnList("WebPageItemNumber")
          If pUserControl = "PRINTRECEIPT" Then
            vList("WebPageHtml") = "<Strong>Template Content Goes Here</Strong>"
          Else
            vList("WebPageHtml") = "<Strong>New Content Goes Here</Strong>"
          End If

          DataHelper.UpdateWebPageData(vList)
        End If
      End If
    End If
  End Sub

  Private Function GetParametersForUserControl(ByVal pUserControl As String, ByRef pInitialSettings As String, ByRef pDefaultSettings As String) As Boolean
    Dim vReturnValue As Boolean = True
    Dim vList As New ParameterList(True)
    vList("WebPageUserControl") = pUserControl
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtWebPageUserControls)
    If vTable IsNot Nothing Then
      vTable.DefaultView.RowFilter = String.Format("WebPageUserControl = '{0}'", pUserControl)
      If vTable.DefaultView.Count = 1 Then
        Dim vInitialParameters As String = vTable.DefaultView.Item(0)("InitialParameters").ToString
        Dim vDefaultParameters As String = vTable.DefaultView.Item(0)("DefaultParameters").ToString
        If vInitialParameters.Length + vDefaultParameters.Length > 0 Then
          vReturnValue = DataHelper.DisplayControlParameters(Me, mvWebNumber, pUserControl, vInitialParameters, vDefaultParameters, pInitialSettings, pDefaultSettings)
        End If
      End If
    End If
    Return vReturnValue
  End Function

  Private Function IsContentControl(ByVal pControlName As String) As Boolean
    Return pControlName = "DATADISPLAY" OrElse pControlName = "CONFIRMREGISTRATION" OrElse pControlName = "PRINTRECEIPT"
  End Function

  Private Function CanCreateRows(ByVal pControlName As String) As Boolean
    Return pControlName = "ADDCATEGORYOPTIONS" OrElse pControlName = "ADDCATEGORYCHECKBOXES"
  End Function

  Private Sub ProcessDelete()
    Try
      If ConfirmDelete() Then
        Dim vList As New ParameterList(True)
        GetPrimaryKeyValues(vList, 0, False)
        DataHelper.DeleteItem(mvMaintenanceType, vList)
        sel.RemoveSelectedNode()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ControlDoubleClick(ByVal sender As Object, ByVal pParameterName As String) Handles epl.ControlDoubleClick
    Try
      'This will be coming from the WebBrowser control and means we need to edit the HTML or Control associated with the item
      Dim vWebPageItem As Integer = GetWebPageItemNumber(pParameterName)
      Dim vRefresh As Boolean
      If vWebPageItem > 0 AndAlso pParameterName.EndsWith("tblContent") Then
        If EditWebPageItemData(vWebPageItem) Then vRefresh = True
      ElseIf pParameterName = "SiteHeader" OrElse pParameterName = "SiteFooter" Then
        If EditSiteData(pParameterName) Then vRefresh = True
      Else
        'double click is coming from a care module
        Dim vForm As New frmModuleContent(vWebPageItem)
        If vForm.HasControls Then
          If vForm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then vRefresh = True
        End If
      End If
      If vRefresh Then DirectCast(epl.FindPanelControl("WebBrowser"), WebBrowser).Refresh()
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub ControlContextMenu(ByVal sender As Object, ByVal pParameterName As String) Handles epl.ControlContextMenu
    mvContextItemNumber = GetWebPageItemNumber(pParameterName)
  End Sub

  Private Function GetWebPageItemNumber(ByVal pParameterName As String) As Integer
    If pParameterName.StartsWith("ItemNumber") Then
      Dim vPos As Integer = pParameterName.IndexOf("_")
      If vPos >= 0 Then
        Return IntegerValue(pParameterName.Substring(10, vPos - 10))
      End If
    End If
  End Function

  Private Function UpdateWebPortalInfo(ByVal pWebNumber As Integer, ByVal pWebName As String) As Boolean
    Try
      Dim vCP As CarePortal.PortalAdmin = GetPortalWS()
      Dim vList As New ParameterList(vCP.SetWebInfo(pWebNumber, pWebName))
      Return True
    Catch ex As Exception
      ShowErrorMessage("Failed to Update the Web Information on the Server")
    End Try
  End Function

  Private Sub SetImages()
    If mvImageTable IsNot Nothing Then
      sel.AddItems(XMLWebDataSelectionTypes.wstImages, mvImageTable)
    Else
      GetImages()
    End If
  End Sub

  Private Function GetImages() As Boolean
    'Return true if a succesfull call was made to get the images
    'We can use this to determine if the URL location is valid
    Debug.Print("Get Images")
    Try
      Dim vCP As CarePortal.PortalAdmin = GetPortalWS()
      Dim vDataSet As DataSet = DataHelper.GetDataSetFromResult(vCP.GetImages)
      If vDataSet IsNot Nothing And vDataSet.Tables.Count > 0 Then
        mvImageTable = vDataSet.Tables(0)
        sel.AddItems(XMLWebDataSelectionTypes.wstImages, mvImageTable)
      Else
        mvImageTable = Nothing
      End If
      Return True
    Catch vEX As System.Net.WebException
      ShowErrorMessage("A Care Portal cannot be found at the given Web URL. Please ensure that it has been installed and the URL is correct")
    Catch vEx As System.UriFormatException
      ShowErrorMessage("The Web URL must be a valid URL")
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Function

  Private Sub SetDocuments()
    If mvDocumentTable IsNot Nothing Then
      sel.AddItems(XMLWebDataSelectionTypes.wstDocuments, mvDocumentTable)
    Else
      GetDocuments()
    End If
  End Sub

  Private Function GetDocuments() As Boolean
    'Return true if a succesfull call was made to get the documents
    'We can use this to determine if the URL location is valid
    Debug.Print("Get Documents")
    Try
      Dim vCP As CarePortal.PortalAdmin = GetPortalWS()
      Dim vDataSet As DataSet = DataHelper.GetDataSetFromResult(vCP.GetDocuments)
      If vDataSet IsNot Nothing And vDataSet.Tables.Count > 0 Then
        mvDocumentTable = vDataSet.Tables(0)
        sel.AddItems(XMLWebDataSelectionTypes.wstDocuments, mvDocumentTable)
      Else
        mvDocumentTable = Nothing
      End If
      Return True
    Catch vEX As System.Net.WebException
      ShowErrorMessage("A Care Portal cannot be found at the given Web URL. Please ensure that it has been installed and the URL is correct")
    Catch vEx As System.UriFormatException
      ShowErrorMessage("The Web URL must be a valid URL")
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Function

  Private Sub UploadDocument(ByVal pFileName As String)
    Try
      Dim vCP As CarePortal.PortalAdmin = GetPortalWS()
      Dim vFileInfo As New FileInfo(pFileName)
      If vFileInfo.Exists Then
        Dim vBuffer As Byte() = Nothing
        Dim vFS As FileStream = Nothing
        vFS = New FileStream(pFileName, FileMode.Open)
        ReDim vBuffer(CInt(vFS.Length - 1))
        vFS.Read(vBuffer, 0, CInt(vFS.Length))
        vFS.Close()
        Dim vResult As String = vCP.UploadDocument(vFileInfo.Name, vBuffer)
        Dim vList As New ParameterList(vResult)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub UploadImage(ByVal pFileName As String)
    Try
      Dim vCP As CarePortal.PortalAdmin = GetPortalWS()
      Dim vFileInfo As New FileInfo(pFileName)
      If vFileInfo.Exists Then
        Dim vBuffer As Byte() = Nothing
        Dim vFS As FileStream = Nothing
        vFS = New FileStream(pFileName, FileMode.Open)
        ReDim vBuffer(CInt(vFS.Length - 1))
        vFS.Read(vBuffer, 0, CInt(vFS.Length))
        vFS.Close()
        Dim vResult As String = vCP.UploadImage(vFileInfo.Name, vBuffer)
        Dim vList As New ParameterList(vResult)
      End If
    Catch vEx As Exception
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Function GetPortalWS() As CarePortal.PortalAdmin
    Dim vCP As New CarePortal.PortalAdmin
    Dim vBuilder As New UriBuilder(New Uri(mvWebURL & "/Services/PortalAdmin.asmx"))
    vCP.Url = vBuilder.Uri.AbsoluteUri
    vCP.Credentials = System.Net.CredentialCache.DefaultCredentials
    Return vCP
  End Function

  Private Sub AddNewImage()
    Dim vOFD As New OpenFileDialog
    With vOFD
      .Title = "Select Image"
      .Filter = "Image Files(*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG"
      If .ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
        UploadImage(.FileName)
        GetImages()
      End If
    End With
  End Sub

  Private Sub AddNewDocument()
    Dim vOFD As New OpenFileDialog
    With vOFD
      .Title = "Select document"
      .Filter = "Document Files(*.DOC;*.PDF;*.TXT)|*.DOC;*.PDF;*.TXT"
      If .ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
        UploadDocument(.FileName)
        GetDocuments()
      End If
    End With
  End Sub

  Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    DataHelper.Logout("WPD")
  End Sub

  Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    If e.KeyCode = Keys.F1 Then
      Dim vForm As New frmBrowser(DataHelper.HelpURL("index.htm", False), False, True)
      vForm.Show()
    End If
  End Sub

  Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    DataHelper.ShowProgress(frmProgress.ProgressStatuses.psNone)
  End Sub

  Private Sub epl_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    Select Case pParameterName
      Case "WebPageControl"
        SetWebPageControlItems(pValue)
      Case "WebPageUserControl"
        If cmdLink1.Visible Then
          cmdLink2.Visible = pValue = "DISPLAYCONTACTDATA"
          If cmdLink2.Visible Then
            cmdLink2.Tag = "Customisation"
            cmdLink2.Text = "Customise"
          End If
          MyBase.bpl.RepositionButtons()
        End If
      Case "LoginTypeRequired"
        If epl.FindPanelControl("AccessViewName", False) IsNot Nothing Then
          Dim vEnable As Boolean = (pValue = "R")
          epl.EnableControl("AccessViewName", vEnable)
          If Not vEnable Then epl.SetValue("AccessViewName", String.Empty)
        End If
      Case "WebName", "WebUrl", "WebPageNumber", "LoginPageNumber", "UpdateDetailsPageNumber"
        cmdSave.Enabled = True
    End Select
  End Sub

  Private Sub SetWebPageControlItems(ByVal pValue As String)
    If IsContentControl(pValue) Then
      epl.SetControlVisible("ItemGroupName", False)
      epl.SetControlVisible("ParentGroupName", False)
      epl.SetControlVisible("NumberOfRows", False)
      epl.SetControlVisible("SubmitItemNumber", False)
      epl.SetControlVisible("SubmitUrl", False)
    Else
      epl.SetControlVisible("ItemGroupName", True)
      epl.SetControlVisible("ParentGroupName", True)
      epl.SetControlVisible("NumberOfRows", CanCreateRows(pValue))
      epl.SetControlVisible("SubmitItemNumber", True)
      epl.SetControlVisible("SubmitUrl", True)
    End If
  End Sub

  Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    If epl.Controls.Count > 0 Then
      Dim vTC As CDBNETCL.TabControl = TryCast(epl.Controls(0), CDBNETCL.TabControl)
      If vTC IsNot Nothing Then
        vTC.Dock = DockStyle.None
        vTC.Size = epl.Size
        vTC.Refresh()
      End If
    End If
  End Sub

  Private Sub PopulateParentGroup(ByVal pParentGroupName As String)
    If epl.Controls("ParentGroupName") IsNot Nothing AndAlso TypeOf epl.Controls("ParentGroupName") Is ComboBox Then
      Dim vParentGroup As ComboBox = epl.FindComboBox("ParentGroupName")
      If mvDataType = XMLWebDataSelectionTypes.wstPageItem Then
        Dim vControl As String = epl.GetValue("WebPageUserControl")
        Select Case vControl
          Case "ADDCATEGORY", "ADDCATEGORYCHECKBOXES", "ADDCATEGORYNOTES", "ADDCATEGORYOPTIONS", "ADDCATEGORYVALUE", "ADDSUPPRESSION"
            Dim vDataTable As New DataTable
            If vParentGroup.DataSource IsNot Nothing Then
              vDataTable = CType(vParentGroup.DataSource, DataTable)
              vDataTable.Rows.Add("RegisteredUser")
              vDataTable.Rows.Add("SelectedContact")
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            Else
              vDataTable.Columns.Add("ItemGroupName")
              vDataTable.Rows.Add("")
              vDataTable.Rows.Add("RegisteredUser")
              vDataTable.Rows.Add("SelectedContact")
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DataSource = vDataTable
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            End If
          Case "MAINTAINNUMBERS", "UPDATEADDRESS", "UPDATECONTACT", "UPDATEPHONENUMBER", "UPDATEEMAILADDRESS", "UPDATEPOSITION", "MEMBERSHIPTYPESELECTION", "ADDMEMBERCC", "ADDMEMBERCI", "ADDMEMBERDD", "INVOICEPAYMENT", "SELECTPAYPLANFORDD", "PROCESSPAYMENT", "ADDMEMBERCS"
            Dim vDataTable As New DataTable
            If vParentGroup.DataSource IsNot Nothing Then
              vDataTable = CType(vParentGroup.DataSource, DataTable)
              vDataTable.Rows.Add("SelectedContact")
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            Else
              vDataTable.Columns.Add("ItemGroupName")
              vDataTable.Rows.Add("")
              vDataTable.Rows.Add("SelectedContact")
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DataSource = vDataTable
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            End If
          Case "UPDATEORGANISATION"
            Dim vDataTable As New DataTable
            If vParentGroup.DataSource IsNot Nothing Then
              vDataTable = CType(vParentGroup.DataSource, DataTable)
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            Else
              vDataTable.Columns.Add("ItemGroupName")
              vDataTable.Rows.Add("")
              vDataTable.Rows.Add("SelectedOrganisation")
              vParentGroup.DataSource = vDataTable
              vParentGroup.DisplayMember = "ItemGroupName"
              vParentGroup.ValueMember = "ItemGroupName"
            End If
        End Select
      End If
      vParentGroup.SelectedValue = pParentGroupName
    End If
  End Sub
End Class
