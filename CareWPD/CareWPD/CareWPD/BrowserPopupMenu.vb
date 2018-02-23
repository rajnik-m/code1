Public Class BrowserPopupMenu
  Inherits ContextMenuStrip

  Public Enum BrowserMenuItems
    bmiAddLeftContent
    bmiAddCenterContent
    bmiAddRightContent
    bmiAddFullWidthContent
    bmiAddCategoryModule
    bmiAddCPDModule
    bmiAddDirectoryModule
    bmiAddDownloadModule
    bmiAddEventsModule
    bmiAddExamsModule
    bmiAddLoginOrRegistrationModule
    bmiAddMembershipOrPaymentPlanModule
    bmiAddSalesModule
    bmiAddSelfServiceModule
    bmiAddSponsorMeModule
    bmiAddSupportModule
    bmiAddSurveyModule
    bmiAddCAREUserModule
    bmiDeletePageItem
    bmiRefresh
  End Enum

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)
  Private mvClientCode As String = ""

  Public Event MenuSelected(ByVal pItem As BrowserMenuItems)
  Public Event ModuleMenuSelected(ByVal pItem As BrowserMenuItems, ByVal pControlName As String)

  Public Sub New()
    MyBase.New()
    mvClientCode = DataHelper.GetClientCode
    With mvMenuItems
      .Add(BrowserMenuItems.bmiAddLeftContent.ToString, New MenuToolbarCommand("AddLeft", "Add Content to Left", BrowserMenuItems.bmiAddLeftContent))
      .Add(BrowserMenuItems.bmiAddCenterContent.ToString, New MenuToolbarCommand("AddCenter", "Add Content to Center", BrowserMenuItems.bmiAddCenterContent))
      .Add(BrowserMenuItems.bmiAddRightContent.ToString, New MenuToolbarCommand("AddRight", "Add Content to Right", BrowserMenuItems.bmiAddRightContent))
      .Add(BrowserMenuItems.bmiAddFullWidthContent.ToString, New MenuToolbarCommand("AddFullWidth", "Add Full width Content", BrowserMenuItems.bmiAddFullWidthContent))
      .Add(BrowserMenuItems.bmiAddCategoryModule.ToString, New MenuToolbarCommand("AddCategoryModule", "Add Category Module", BrowserMenuItems.bmiAddCategoryModule))
      .Add(BrowserMenuItems.bmiAddCPDModule.ToString, New MenuToolbarCommand("AddCPDModule", "Add CPD Module", BrowserMenuItems.bmiAddCPDModule))
      .Add(BrowserMenuItems.bmiAddDirectoryModule.ToString, New MenuToolbarCommand("AddDirectoryModule", "Add Directory Module", BrowserMenuItems.bmiAddDirectoryModule))
      .Add(BrowserMenuItems.bmiAddDownloadModule.ToString, New MenuToolbarCommand("AddDownloadModule", "Add Downloads Module", BrowserMenuItems.bmiAddDownloadModule))
      .Add(BrowserMenuItems.bmiAddEventsModule.ToString, New MenuToolbarCommand("AddEventsModule", "Add Events Module", BrowserMenuItems.bmiAddEventsModule))
      .Add(BrowserMenuItems.bmiAddExamsModule.ToString, New MenuToolbarCommand("AddExamsModule", "Add Exams Module", BrowserMenuItems.bmiAddExamsModule))
      .Add(BrowserMenuItems.bmiAddLoginOrRegistrationModule.ToString, New MenuToolbarCommand("AddLoginOrRegistrationModule", "Add Login / Registration Module", BrowserMenuItems.bmiAddLoginOrRegistrationModule))
      .Add(BrowserMenuItems.bmiAddMembershipOrPaymentPlanModule.ToString, New MenuToolbarCommand("AddMembershipOrPaymentPlanModule", "Add Membership / Payment Plan Module", BrowserMenuItems.bmiAddMembershipOrPaymentPlanModule))
      .Add(BrowserMenuItems.bmiAddSalesModule.ToString, New MenuToolbarCommand("AddSalesModule", "Add Sales Module", BrowserMenuItems.bmiAddSalesModule))
      .Add(BrowserMenuItems.bmiAddSelfServiceModule.ToString, New MenuToolbarCommand("AddSelfServiceModule", "Add Self Service Module", BrowserMenuItems.bmiAddSelfServiceModule))
      .Add(BrowserMenuItems.bmiAddSponsorMeModule.ToString, New MenuToolbarCommand("AddSponsorMeModule", "Add Sponsor Me Module", BrowserMenuItems.bmiAddSponsorMeModule))
      .Add(BrowserMenuItems.bmiAddSupportModule.ToString, New MenuToolbarCommand("AddSupportModule", "Add Support Module", BrowserMenuItems.bmiAddSupportModule))
      .Add(BrowserMenuItems.bmiAddSurveyModule.ToString, New MenuToolbarCommand("AddSurveyModule", "Add Survey Module", BrowserMenuItems.bmiAddSurveyModule))
      .Add(BrowserMenuItems.bmiAddCAREUserModule.ToString, New MenuToolbarCommand("AddCAREUserModule", "Add User Module", BrowserMenuItems.bmiAddCAREUserModule))
      .Add(BrowserMenuItems.bmiDeletePageItem.ToString, New MenuToolbarCommand("DeletePageItem", "Delete Page Item", BrowserMenuItems.bmiDeletePageItem))
      .Add(BrowserMenuItems.bmiRefresh.ToString, New MenuToolbarCommand("Refresh", "Refresh", BrowserMenuItems.bmiRefresh))
    End With

    For Each vItem As MenuToolbarCommand In mvMenuItems
      vItem.OnClick = AddressOf MenuHandler
      Me.Items.Add(vItem.MenuStripItem)
    Next
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtWebPageUserControls)
    If vTable IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        Dim vHideItem As Boolean = False
        Select Case vRow.Item("WebPageUserControl").ToString
          Case "FINDFUNDRAISER", "FUNDRAISERPAGEINFO", "ADDFUNDRAISINGPAGE", "FUNDRAISERDONATIONS"
            mvMenuItems(BrowserMenuItems.bmiAddSponsorMeModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "ADDCONTACT", "ADDRELATEDCONTACT", "CONTACTSELECTION", "DISPLAYCONTACTDATA", "ADDCOMMUNICATIONNOTE", "ADDACTION", "RAPIDPRODUCTPURCHASE", "CONTACTSELECTIONEXTERNALREF", "ADDBANKACCOUNT", "DATADISPLAY", "DISPLAYRELATEDTEDCONTACTS", "DISPLAYRELATEDTEDORGANISATIONS", "EMAILSELECTEDCONTACTS", "ADDPOSITION"
            mvMenuItems(BrowserMenuItems.bmiAddCAREUserModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "ADDCATEGORY", "ADDCATEGORYCHECKBOXES", "ADDCATEGORYNOTES", "ADDCATEGORYOPTIONS", "ADDCATEGORYVALUE"
            mvMenuItems(BrowserMenuItems.bmiAddCategoryModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "BOOKEVENT", "BOOKEVENTCC", "BOOKINGOPTIONSELECTION", "EVENTSELECTION", "EVENTSELECTIONCALENDAR", "PICKSESSIONS", "SHOWEVENTBOOKINGS", "EVENTDELEGATEACTIVITIES", "EVENTDELEGATEENTRY", "EVENTDELEGATESELECTION"
            mvMenuItems(BrowserMenuItems.bmiAddEventsModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "BOOKEXAM", "EXAMSELECTION", "SHOWEXAMBOOKINGS", "SHOWEXAMHISTORY"
            mvMenuItems(BrowserMenuItems.bmiAddExamsModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "LOGIN", "LOGOUT", "REGISTER", "CONFIRMREGISTRATION", "EMAILPASSWORD", "REGISTERMEMBER", "REGISTERCORPORATEMEMBER", "REGISTERCOMBINED", "DEDUPLICATEORGANISATIONS"
            mvMenuItems(BrowserMenuItems.bmiAddLoginOrRegistrationModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "MODIFYDD", "ADDMEMBERCC", "ADDMEMBERCI", "ADDMEMBERDD", "ADDPAYMENTPLANDD", "ADDSUBSCRIPTIONDD", "MEMBERSHIPTYPESELECTION", "PAYMEMBERSHIPCC", "PAYSUBSCRIPTIONCC", "PAYMULTIPLEPAYMENTPLANS", "ADDMEMBERCS", _
               "CONVERTPAYPLANTODD", "SELECTPAYPLANFORDD"
            mvMenuItems(BrowserMenuItems.bmiAddMembershipOrPaymentPlanModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "INVOICEPAYMENT", "PRODUCTSELECTION", "MAKEDONATIONCC", "PRODUCTPURCHASE", "PRODUCTPURCHASECC", "VIEWTRANSACTION", "PROCESSPAYMENT", "PAYERSELECTION"
            mvMenuItems(BrowserMenuItems.bmiAddSalesModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "ADDEXTERNALREFERENCE", "ADDLINK", "ADDDEFAULTADDRESS", "ADDNETWORKCONTACT", "ADDSUPPRESSION", "FINDLOCALGROUP", "MAINTAINACTIVITYGROUP", "MAINTAINNUMBERS", "RECORDACTIVITY", "RECORDSTANDARDEMAIL", "SHOWDEFAULTADDRESS", "SHOWNETWORK", "SUBMITENQUIRY", "UPDATEADDRESS", "UPDATECONTACT", "UPDATEEMAILADDRESS", "UPDATEPASSWORD", "UPDATEPHONENUMBER", "UPDATEPOSITION", "ADDORGANISATION", "SEARCHORGANISATION", "UPDATEORGANISATION", "SEARCHCONTACT", "DISPLAYTRANSACTIONS", "PRINTRECEIPT", "SETUSERORGANISATION"
            mvMenuItems(BrowserMenuItems.bmiAddSelfServiceModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "NAVIGATEBUTTON", "SUBMITALL", "PRINTBUTTON"
            mvMenuItems(BrowserMenuItems.bmiAddSupportModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "SURVEYSELECTION", "SURVEYRESPONSES"
            mvMenuItems(BrowserMenuItems.bmiAddSurveyModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "VIEWDIRECTORYDETAILS", "SEARCHDIRECTORY", "DIRECTORYPREFERENCES"
            mvMenuItems(BrowserMenuItems.bmiAddDirectoryModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "CONTACTCPDCYCLE", "UPDATECPDPOINTS", "UPDATECPDOBJECTIVES"
            mvMenuItems(BrowserMenuItems.bmiAddCPDModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
          Case "DOWNLOADSELECTION", "DOWNLOADDOCUMENT"
            mvMenuItems(BrowserMenuItems.bmiAddDownloadModule).FindMenuStripItem(Me).DropDownItems.Add(vRow.Item("ControlTitle").ToString, Nothing, AddressOf ModuleMenuHandler).Tag = vRow.Item("WebPageUserControl")
        End Select
      Next
    End If

    MenuToolbarCommand.SetAccessControl(mvMenuItems)

    Dim vList As New ParameterList(True)
    Dim vReturn As ParameterList
    vList("ModuleCode") = "WPDC"
    vReturn = DataHelper.CheckLicenseData(CareNetServices.LicenseCheckTypes.CheckModule, vList)
    If vReturn.ContainsKey("LicensedUsers") AndAlso vReturn.IntegerValue("LicensedUsers") > 0 Then
      'Licensed for Care Modules
    Else
      mvMenuItems(BrowserMenuItems.bmiAddCAREUserModule).FindMenuStripItem(Me).Enabled = False
    End If
    vList("ModuleCode") = "WPDS"
    vReturn = DataHelper.CheckLicenseData(CareNetServices.LicenseCheckTypes.CheckModule, vList)
    If vReturn.ContainsKey("LicensedUsers") AndAlso vReturn.IntegerValue("LicensedUsers") > 0 Then
      'Licensed for Sponsor me
    Else
      mvMenuItems(BrowserMenuItems.bmiAddSponsorMeModule).FindMenuStripItem(Me).Enabled = False
    End If
  End Sub

  Private Sub ModuleMenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
      Select Case vMenuItem.Tag.ToString
        Case "FINDFUNDRAISER", "FUNDRAISERPAGEINFO", "ADDFUNDRAISINGPAGE", "FUNDRAISERDONATIONS"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddSponsorMeModule, vMenuItem.Tag.ToString)
        Case "ADDCONTACT", "ADDRELATEDCONTACT", "CONTACTSELECTION", "DISPLAYCONTACTDATA", "ADDCOMMUNICATIONNOTE", "ADDACTION", "RAPIDPRODUCTPURCHASE", "CONTACTSELECTIONEXTERNALREF",
          "ADDBANKACCOUNT", "DATADISPLAY", "DISPLAYRELATEDTEDCONTACTS", "DISPLAYRELATEDTEDORGANISATIONS", "EMAILSELECTEDCONTACTS", "ADDPOSITION"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddCAREUserModule, vMenuItem.Tag.ToString)
        Case "ADDCATEGORY", "ADDCATEGORYCHECKBOXES", "ADDCATEGORYNOTES", "ADDCATEGORYOPTIONS", "ADDCATEGORYVALUE"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddCategoryModule, vMenuItem.Tag.ToString)
        Case "BOOKEVENT", "BOOKEVENTCC", "BOOKINGOPTIONSELECTION", "EVENTSELECTION", "EVENTSELECTIONCALENDAR", "PICKSESSIONS", "SHOWEVENTBOOKINGS", "EVENTDELEGATEACTIVITIES", "EVENTDELEGATEENTRY", "EVENTDELEGATESELECTION"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddEventsModule, vMenuItem.Tag.ToString)
        Case "BOOKEXAM", "EXAMSELECTION", "SHOWEXAMBOOKINGS", "SHOWEXAMHISTORY"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddExamsModule, vMenuItem.Tag.ToString)
        Case "LOGIN", "LOGOUT", "REGISTER", "CONFIRMREGISTRATION", "EMAILPASSWORD", "REGISTERMEMBER", "REGISTERCORPORATEMEMBER", "REGISTERCOMBINED", "DEDUPLICATEORGANISATIONS"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddLoginOrRegistrationModule, vMenuItem.Tag.ToString)
        Case "MODIFYDD", "ADDMEMBERCC", "ADDMEMBERCI", "ADDMEMBERDD", "ADDPAYMENTPLANDD", "ADDSUBSCRIPTIONDD", "MEMBERSHIPTYPESELECTION", "PAYMEMBERSHIPCC", "PAYSUBSCRIPTIONCC", "PAYMULTIPLEPAYMENTPLANS", "ADDMEMBERCS", _
             "CONVERTPAYPLANTODD", "SELECTPAYPLANFORDD"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddMembershipOrPaymentPlanModule, vMenuItem.Tag.ToString)
        Case "PRODUCTSELECTION", "MAKEDONATIONCC", "PRODUCTPURCHASE", "PRODUCTPURCHASECC", "VIEWTRANSACTION", "PROCESSPAYMENT", "PAYERSELECTION", "INVOICEPAYMENT"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddSalesModule, vMenuItem.Tag.ToString)
        Case "ADDEXTERNALREFERENCE", "ADDLINK", "ADDDEFAULTADDRESS", "ADDNETWORKCONTACT", "ADDSUPPRESSION", "FINDLOCALGROUP", "MAINTAINACTIVITYGROUP", "MAINTAINNUMBERS",
          "RECORDACTIVITY", "RECORDSTANDARDEMAIL", "SHOWDEFAULTADDRESS", "SHOWNETWORK", "SUBMITENQUIRY", "UPDATEADDRESS", "UPDATECONTACT", "UPDATEEMAILADDRESS",
          "UPDATEPASSWORD", "UPDATEPHONENUMBER", "UPDATEPOSITION", "ADDORGANISATION", "SEARCHORGANISATION", "UPDATEORGANISATION", "SEARCHCONTACT", "DISPLAYTRANSACTIONS", "PRINTRECEIPT", "SETUSERORGANISATION"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddSelfServiceModule, vMenuItem.Tag.ToString)
        Case "NAVIGATEBUTTON", "SUBMITALL", "PRINTBUTTON"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddSupportModule, vMenuItem.Tag.ToString)
        Case "SURVEYSELECTION", "SURVEYRESPONSES"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddSurveyModule, vMenuItem.Tag.ToString)
        Case "VIEWDIRECTORYDETAILS", "SEARCHDIRECTORY", "DIRECTORYPREFERENCES"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddDirectoryModule, vMenuItem.Tag.ToString)
        Case "CONTACTCPDCYCLE", "UPDATECPDPOINTS", "UPDATECPDOBJECTIVES"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddCPDModule, vMenuItem.Tag.ToString)
        Case "DOWNLOADSELECTION", "DOWNLOADDOCUMENT"
          RaiseEvent ModuleMenuSelected(BrowserMenuItems.bmiAddDownloadModule, vMenuItem.Tag.ToString)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub MenuHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim vCursor As New BusyCursor
    Try
      Dim vMenuItem As BrowserMenuItems = CType(DirectCast(DirectCast(sender, ToolStripMenuItem).Tag, MenuToolbarCommand).CommandID, BrowserMenuItems)
      Select Case vMenuItem
        Case BrowserMenuItems.bmiAddLeftContent, BrowserMenuItems.bmiAddCenterContent, BrowserMenuItems.bmiAddRightContent, _
         BrowserMenuItems.bmiDeletePageItem, BrowserMenuItems.bmiRefresh, BrowserMenuItems.bmiAddFullWidthContent
          RaiseEvent MenuSelected(vMenuItem)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub

  Private Sub WebPopupMenu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
    For Each vItem As ToolStripItem In Me.Items
      vItem.Visible = True
    Next
  End Sub


End Class
