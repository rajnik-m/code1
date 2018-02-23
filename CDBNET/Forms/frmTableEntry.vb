Public Class frmTableEntry

  Private mvTableName As String
  Private mvParams As ParameterList
  Private mvCriteria As ParameterList
  Private mvEditMode As CareNetServices.XMLTableMaintenanceMode
  Private mvRequiresRefresh As Boolean
  Private mvEditPanelInfo As EditPanelInfo
  Private mvStockFlagBefore As Integer    'Flag that checkes if Stock Item Flag was Y or N
  Private mvStockFlagAfter As Integer    'Flag that checkes if Stock Item Flag is changed
  Private mvInternalResourceMesgShown As Boolean
  Private mvCancel As Boolean
  Private mvAddMore As Boolean
  Private mvReturnParams As Boolean
  Private mvTableMaintenance As Boolean

  Public Sub New(ByVal pEditMode As CareNetServices.XMLTableMaintenanceMode, ByVal pTable As String, ByVal pParams As ParameterList, ByVal pCriteria As ParameterList, Optional ByVal pAddMore As Boolean = False, Optional ByVal pReturnParams As Boolean = False)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvEditMode = pEditMode
    mvTableName = pTable
    mvParams = pParams
    mvCriteria = pCriteria
    mvAddMore = pAddMore
    mvReturnParams = pReturnParams
    InitialiseControls()
  End Sub

  Private Sub frmTableEntry_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    'set the splitter distance to required width of the edit panel
    'this causes the notes textbox to resize along with the form
    'spc.SplitterDistance = epl.RequiredWidth
  End Sub

  Private Sub epl_QueryWildcardAndValidation(ByVal sender As Object, ByVal pParameterName As String, ByRef pWildcardsSupported As Boolean, ByRef pValidationRequired As Boolean) Handles epl.QueryWildcardAndValidation
    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmSelect Then
      pWildcardsSupported = True
      pValidationRequired = False
    End If
  End Sub

  Public Property TableMaintenance As Boolean
    Get
      Return mvTableMaintenance
    End Get
    Set(pValue As Boolean)
      mvTableMaintenance = pValue
    End Set
  End Property

  Private Sub epl_ShowStatusMessage(ByVal sender As System.Object, ByVal pMessage As System.String) Handles epl.ShowStatusMessage
    txtNotes.Text = pMessage.Replace(Chr(10).ToString, Environment.NewLine)
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      mvInternalResourceMesgShown = False
      mvCancel = False
      If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmSelect Then
        Dim vValid As Boolean
        Dim vParamCount As Integer = mvParams.Count
        vValid = epl.AddValuesToList(mvParams, True, EditPanel.AddNullValueTypes.anvtCheckBoxesOnly, False)
        If vValid AndAlso epl.DataChanged Then
          'Add IgnoreUnknownParameters as the ValidateParameters may not contain all the Columns
          'and will throw an error while selecting the data
          If Not mvParams.Contains("IgnoreUnknownParameters") Then mvParams("IgnoreUnknownParameters") = "Y"
        End If
        If vValid Then
          Me.Close()
          Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End If
      Else
        'Add/Edit mode. 
        If epl.DataChanged OrElse mvTableName = "report_version_history" Then
          If ProcessSave() Then
            mvCancel = True
            CloseForm()
          End If
        Else
          CloseForm()
        End If
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      ElseIf vException.ErrorNumber = CareException.ErrorNumbers.enParameterInvalidValue Then
        ShowInformationMessage(vException.Message)
      Else
        DataHelper.HandleException(vException)
      End If
      mvCancel = True 'if an error occurs,allow the form to close
    End Try
  End Sub

  Private Sub cmdAddMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddMore.Click
    Try
      If ProcessSave() Then
        epl.Clear()
        mvParams = New ParameterList(True)
        If mvTableName = "config_names" Then
          mvParams("CreatedBy") = If(My.User.Name.Contains("\"), My.User.Name.Substring(My.User.Name.LastIndexOf("\") + 1), My.User.Name)
          mvParams("CreatedOn") = AppValues.TodaysDate
          mvParams("CreatedVersion") = String.Format("{0:#0}.{1:0}.{2:0000}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)
          mvParams("AmendedBy") = If(My.User.Name.Contains("\"), My.User.Name.Substring(My.User.Name.LastIndexOf("\") + 1), My.User.Name)
          mvParams("AmendedOn") = AppValues.TodaysDate
          mvParams("AmendedVersion") = String.Format("{0:#0}.{1:0}.{2:0000}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)
        Else
          mvParams("CreatedBy") = AppValues.Logname
          mvParams("CreatedOn") = AppValues.TodaysDate
          mvParams("AmendedBy") = AppValues.Logname
          mvParams("AmendedOn") = AppValues.TodaysDate
        End If
        SetDefaults()
        InitForm()
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vException)
      End If
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    CloseForm()
  End Sub

  Private Sub cmdHTMLEditor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdHTMLEditor.Click
    Try
      Dim vForm As New frmHTMLEditor
      vForm.InnerHtml = epl.GetValue("HtmlText")
      If vForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then epl.SetValue("HtmlText", vForm.InnerHtml)
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

#Region "Private Methods"

  ''' <summary>
  ''' Set the default values and disable the controls that have been passed in as criteria from the parent form
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SetDefaults()
    Dim vList As New ParameterList(True)
    Dim vParam As String

    For Each vParam In mvParams.Keys
      vList(vParam) = mvParams(vParam)
    Next

    If Not mvCriteria Is Nothing AndAlso mvCriteria.Count > 0 Then
      For Each vParam In mvCriteria.Keys
        If Not FindControl(epl, vParam, False) Is Nothing Then
          epl.SetReadOnly(vParam, True)
          vList(vParam) = mvCriteria(vParam)
        End If
      Next
    End If
    epl.Populate(vList)

    If mvTableName = "membership_controls" Then
      If epl.GetValue("CmtCalcProportionalBalance") = "N" AndAlso epl.GetValue("AdvancedCmt") = "N" Then epl.EnableControl("AdvancedCmt", False)
    End If

    If mvTableName.Equals("payment_frequencies", StringComparison.InvariantCultureIgnoreCase) = True AndAlso FindControl(epl, "OffsetMonths", False) IsNot Nothing Then
      Dim vPeriod As String = epl.GetValue("Period")
      If vPeriod.Length > 0 AndAlso vPeriod.Equals("M", StringComparison.InvariantCultureIgnoreCase) = False Then epl.EnableControl("OffsetMonths", False)
    End If

    If mvTableName.Equals("prize_draws", StringComparison.InvariantCultureIgnoreCase) Then
      If vList.ContainsKey("CloseDate") _
      AndAlso (mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew OrElse String.IsNullOrWhiteSpace(vList("CloseDate")) = False) Then
        'Make control mandatory
        Dim vDTP As DateTimePicker = epl.FindPanelControl(Of DateTimePicker)("CloseDate", False)
        If vDTP IsNot Nothing Then
          vDTP.ShowCheckBox = False
        End If
      End If
    End If

    If mvTableName.Equals("contact_alerts", StringComparison.InvariantCultureIgnoreCase) Then
      If vList.ContainsKey("ContactAlertType") AndAlso vList.ContainsKey("FpApplicationNumber") Then
        epl.EnableControl("ContactAlertType", False)
        vList.Remove("FpApplicationNumber")
      End If
      If epl.GetValue("ContactAlertType").ToUpper.Equals("C") Then
        epl.EnableControl("ContactAlertMessageType", False)
      End If
    End If

  End Sub

  ''' <summary>
  ''' Validates the controls and displays saves the values in the DB
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function ProcessSave() As Boolean 'Return true if saved
    Try
      Dim vParams As New ParameterList(True)
      If mvTableName = "config" Then epl.SetErrorField("ConfigValue", "", False)
      If mvTableName = "packages" Then
        epl.SetErrorField("DefaultDocument", "", False)
      End If
      Dim vValid As Boolean = epl.AddValuesToList(vParams, True, EditPanel.AddNullValueTypes.anvtAll, False)
      'check if the value entered is a unc path
      If mvTableName = "config" AndAlso vParams("ConfigName") = "me_membership_card_image" AndAlso Not vParams("ConfigValue").StartsWith("\\") Then
        vValid = False
        epl.SetErrorField("ConfigValue", InformationMessages.ImUNCPathOnly, True)
      End If

      Select Case mvTableName
        Case "gift_aid_controls"
          If vParams("ClaimFileFormat") = "O" AndAlso vParams("SubmitterContact").Length = 0 Then
            epl.SetErrorField("SubmitterContact", InformationMessages.ImFieldMandatory)
            vValid = False
          End If

        Case "packages"
          If vParams.ContainsKey("DefaultDocument") AndAlso vParams.Item("DefaultDocument").Length > 0 Then
            If Not vParams("DefaultDocument").StartsWith("\\") Then
              vValid = False
              epl.SetErrorField("DefaultDocument", InformationMessages.ImUNCPathOnly, True)
            ElseIf (vParams.ContainsKey("DocfileExtension") AndAlso vParams.Item("DocfileExtension").Length > 0 AndAlso Not vParams("DefaultDocument").ToLower.EndsWith(vParams.Item("DocfileExtension").ToLower)) _
              OrElse vParams.Item("DefaultDocument").Contains(".") = False OrElse vParams.Item("DefaultDocument").EndsWith(".") Then
              vValid = False
              epl.SetErrorField("DefaultDocument", InformationMessages.ImInvalidFileExtension, True)
            ElseIf vParams.Item("DefaultDocument").Contains(".") AndAlso Not vParams.Item("DefaultDocument").EndsWith(".") Then
              'Default Document set and Docfile Extension not set- set Docfile Extension from Default Document
              vParams("DocfileExtension") = vParams.Item("DefaultDocument").Substring(vParams.Item("DefaultDocument").LastIndexOf("."))
            End If
          End If
          If vParams.ContainsKey("StorageType") AndAlso vParams.ContainsKey("StoragePath") Then
            If Not String.IsNullOrWhiteSpace(vParams("StorageType")) AndAlso vParams("StorageType").Equals("E", StringComparison.InvariantCultureIgnoreCase) Then
              'External storage
              Dim vExternalPath As String = vParams("StoragePath").Trim
              If vExternalPath.Length = 0 Then
                vValid = epl.SetErrorField("StoragePath", InformationMessages.ImStoragePathMustBeSetForExternal)
              ElseIf vExternalPath.StartsWith("\\") = False Then
                vValid = epl.SetErrorField("StoragePath", InformationMessages.ImUNCPathOnly)
              End If
            End If
          End If

        Case "payment_frequencies"
          If vParams.ContainsKey("OffsetMonths") AndAlso IntegerValue(vParams("OffsetMonths")) > 0 Then
            Dim vPayments As Integer = (IntegerValue(vParams("Frequency")) * IntegerValue(vParams("Interval")))
            Dim vMaxOffset As Integer = (IntegerValue(vParams("Interval")) - 1)
            If vPayments < 12 Then

            End If
          End If
      End Select



      If vValid Then
        If mvTableName = "membership_controls" OrElse mvTableName = "membership_types" Then vValid = CheckMembershipGroups()
        If vValid Then
          If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
            vValid = ConfirmUpdate()
          ElseIf mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
            vValid = ConfirmInsert()
          End If

          If vValid Then
            Select Case mvTableName
              Case "config_names"
                If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
                  If vParams.ContainsKey("CreatedVersion") = True AndAlso vParams.ContainsKey("AmendedVersion") = False Then vParams("AmendedVersion") = vParams("CreatedVersion")
                  If vParams.ContainsKey("CreatedBy") = True AndAlso vParams.ContainsKey("AmendedBy") = False Then
                    vParams("AmendedBy") = vParams("CreatedBy")
                    vParams("AmendedOn") = vParams("CreatedOn")
                  End If
                End If

              Case "financial_controls"
                Dim vDateString As String = vParams("LoanCapitalisationDate")
                Dim vDate As Date = DateSerial(Year(Now), IntegerValue(vDateString.Substring(2)), IntegerValue(vDateString.Substring(0, 2)))
                vParams("LoanCapitalisationDate") = vDate.ToShortDateString

              Case "gift_aid_controls"
                'Need to manipulate the AccountingPeriodStart and TaxYearStart dates from dd/mm to dd/mm/yyyy
                Dim vDateString As String = ""
                Dim vAccountDate As Date = Nothing
                If vParams.ContainsKey("AccountingPeriodStart") Then
                  vDateString = vParams("AccountingPeriodStart")
                  vAccountDate = DateSerial(Year(Now), IntegerValue(vDateString.Substring(2, 2)), IntegerValue(vDateString.Substring(0, 2)))
                  vParams("AccountingPeriodStart") = vAccountDate.ToShortDateString
                End If
                If vParams.ContainsKey("TaxYearStart") Then
                  vDateString = vParams("TaxYearStart")
                  vAccountDate = DateSerial(Year(Now), IntegerValue(vDateString.Substring(2)), IntegerValue(vDateString.Substring(0, 2)))
                  vParams("TaxYearStart") = vAccountDate.ToShortDateString
                End If

              Case "sub_topics"
                vParams("CallBackMinutes") = ""
                If vParams("CallBackD").ToString.Length > 0 OrElse vParams("CallBackH").ToString.Length > 0 OrElse vParams("CallBackM").ToString.Length > 0 Then
                  vParams("CallBackMinutes") = (TimeSpan.FromDays(DoubleValue(vParams("CallBackD").ToString)).TotalMinutes + TimeSpan.FromHours(DoubleValue(vParams("CallBackH").ToString)).TotalMinutes + TimeSpan.FromMinutes(DoubleValue(vParams("CallBackM").ToString)).TotalMinutes).ToString
                End If
                vParams.Remove("CallBackD")
                vParams.Remove("CallBackH")
                vParams.Remove("CallBackM")
            End Select

            vParams("MaintenanceTableName") = mvTableName
            If mvTableMaintenance Then vParams("TableMaintenance") = "Y"

            'To save the values entered or modified in the table entry form. An Insert or Update is performed as required
            If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
              DataHelper.AddTableMaintenanceData(vParams)
            Else
              'Amendment of a record
              If mvStockFlagAfter = 1 Then
                vParams("StockFlagAfter") = "Y"
              End If

              If mvTableName = "membership_controls" OrElse mvTableName = "membership_types" Then
                Dim vQuestion As String = String.Empty
                Dim vOrgGroup As String = String.Empty
                Dim vMemberTypeCode As String = String.Empty
                Dim vHistoric As Boolean
                vParams("SetMembershipGroups") = IIf(SetMembershipGroups(vQuestion, vOrgGroup, vMemberTypeCode, vHistoric), "Y", "N").ToString
                vParams("OrgGroup") = vOrgGroup
                vParams("MemberTypeCode") = vMemberTypeCode
                vParams("Historic") = IIf(vHistoric, "Y", "N").ToString
              End If

              'If we are updating values in tables that dont have a primary key set then pass in the old
              'values as well so that we can initialise the record and update it
              Dim vList As New ParameterList(True)
              vList("TableName") = mvTableName
              vList("TableMaintenance") = "Y"   'use this flag to determine if we need to send back original values
              If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPrimaryKeys, vList) = 0 Then
                GetOriginalValues(vParams, mvParams)
                If mvCriteria IsNot Nothing Then GetOriginalValues(vParams, mvCriteria)
              End If
              DataHelper.UpdateTableMaintenanceData(vParams)
              If mvTableName = "membership_types" Then
                CheckFutureMembershipChange()
              End If
            End If
          End If

          'Update the grid on the main form since the data has changed
          mvRequiresRefresh = True
          'On Add/Edit mode returns ParameterList of Flag is set to true.
          If mvReturnParams Then mvParams.FillFromValueList(vParams.ValueList)
        End If
      End If
      Return vValid
    Catch vEx As CareException

      Select Case vEx.ErrorNumber
        Case CareException.ErrorNumbers.enDuplicateRecord
          If mvTableName = "rate_modifiers" Then
            epl.SetErrorField("SequenceNumber", InformationMessages.ImDuplicateValue)
            epl.TabSelectedIndex = 1
          Else
            ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
          End If
        Case CareException.ErrorNumbers.enPackageStorageTypeCannotChange, CareException.ErrorNumbers.enValueAlreadyUsed, _
             CareException.ErrorNumbers.enCPDCycleTypesStartEndMonthsCannotChange, CareException.ErrorNumbers.enViewInContactCardCannotBeChanged
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enRecordCannotBeMadeHistoric
          ShowInformationMessage(vEx.Message)

        Case CareException.ErrorNumbers.enSurveyAnswerRangeNotAppropriate,
          CareException.ErrorNumbers.enSurveyAnswerListNotAppropriate,
          CareException.ErrorNumbers.enSurveyAnswerMinimumNotAppropriate,
          CareException.ErrorNumbers.enSurveyAnswerMaximumNotAppropriate,
          CareException.ErrorNumbers.enSurveyAnswerListEmpty,
          CareException.ErrorNumbers.enSurveyQuestionNumberInvalid,
          CareException.ErrorNumbers.enMaximumValueGreaterThanMinimum
          ShowInformationMessage(vEx.Message)
        Case CareException.ErrorNumbers.enMailingMarketingFlagInvalid
          ShowInformationMessage(vEx.Message)
        Case Else
          DataHelper.HandleException(vEx)
      End Select
    End Try



  End Function

  Private Sub GetOriginalValues(ByVal pList As ParameterList, ByVal pParams As ParameterList)
    For Each vKey As String In pParams.Keys
      Select Case vKey
        Case "AmendedOn", "AmendedBy", "AmendedVersion", "CreatedBy", "CreatedOn", "CreatedVersion"
          'Ignore
        Case Else
          If epl.PanelInfo.PanelItems.Exists(vKey) Then
            pList("Old" & vKey) = pParams(vKey)
          ElseIf vKey = "OrderNumber" AndAlso epl.PanelInfo.PanelItems.Exists("PaymentPlanNumber") Then
            pList("Old" & vKey) = pParams(vKey)
          End If
      End Select
    Next
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()

    cmdAddMore.Visible = mvAddMore AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew

    'for config and lookup_group_details the control info will vary based on the actual value selected
    'so pass in the criteria as well to fetch the controls
    If mvTableName = "config" OrElse mvTableName = "lookup_group_details" Then
      mvEditPanelInfo = New EditPanelInfo(mvEditMode, mvTableName, mvCriteria)
    Else
      If mvTableName = "html_scripts" Then cmdHTMLEditor.Visible = True
      mvEditPanelInfo = New EditPanelInfo(mvEditMode, mvTableName)
    End If

    If mvParams Is Nothing Then mvParams = New ParameterList(True)
    If mvTableName = "config_names" Then
      mvParams("CreatedBy") = If(My.User.Name.Contains("\"), My.User.Name.Substring(My.User.Name.LastIndexOf("\") + 1), My.User.Name)
      mvParams("CreatedOn") = AppValues.TodaysDate
      mvParams("CreatedVersion") = String.Format("{0:#0}.{1:0}.{2:0000}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)
      mvParams("AmendedBy") = If(My.User.Name.Contains("\"), My.User.Name.Substring(My.User.Name.LastIndexOf("\") + 1), My.User.Name)
      mvParams("AmendedOn") = AppValues.TodaysDate
      mvParams("AmendedVersion") = String.Format("{0:#0}.{1:0}.{2:0000}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)
    Else
      mvParams("CreatedBy") = AppValues.Logname
      mvParams("CreatedOn") = AppValues.TodaysDate
      mvParams("AmendedBy") = AppValues.Logname
      mvParams("AmendedOn") = AppValues.TodaysDate
    End If

    If mvTableName = "sub_topics" Then
      Dim vItem As PanelItem = mvEditPanelInfo.PanelItems("CallBackMinutes")
      vItem.ControlWidth = IntegerValue(vItem.ControlWidth / 5)
      vItem.ControlCaption = "Call Back:  Days"
      vItem.ParameterName = "CallBackD"
      vItem.EntryLength = 2
      Dim vRect As New Rectangle(vItem.ControlLeft + 50, vItem.ControlTop, vItem.ControlWidth, vItem.ControlHeight)
      Dim vNewItem As New PanelItem("CallBackH", PanelItem.ControlTypes.ctTextBox, vRect, "Hours", 40, PanelItem.FieldTypes.cftInteger)
      vNewItem.EntryLength = 2
      vNewItem.MaximumValue = "23"
      mvEditPanelInfo.PanelItems.AddAfter(vNewItem, vItem)
      vItem = vNewItem
      vRect.X = vRect.X + 90
      vNewItem = New PanelItem("CallBackM", PanelItem.ControlTypes.ctTextBox, vRect, "Minutes", 50, PanelItem.FieldTypes.cftInteger)
      vNewItem.EntryLength = 2
      vNewItem.MaximumValue = "59"
      mvEditPanelInfo.PanelItems.AddAfter(vNewItem, vItem)
      If mvParams.ContainsKey("CallBackMinutes") AndAlso mvParams("CallBackMinutes").ToString.Length > 0 Then
        Dim vTimeSpan As TimeSpan = TimeSpan.FromMinutes(DoubleValue(mvParams("CallBackMinutes")))
        mvParams("CallBackD") = vTimeSpan.Days.ToString.PadLeft(2, "0"c)
        mvParams("CallBackH") = vTimeSpan.Hours.ToString.PadLeft(2, "0"c)
        mvParams("CallBackM") = vTimeSpan.Minutes.ToString.PadLeft(2, "0"c)
      End If
    End If

    RemoveLookupRestrictions()
    epl.Init(mvEditPanelInfo)
    SetDefaults()
    'BR18139
    If mvTableName = "ownership_groups" Then
      Dim vPrincDeptLogname As TextLookupBox = epl.FindTextLookupBox("PrincipalDepartmentLogname", False)
      Dim vPrincDept As TextLookupBox = epl.FindTextLookupBox("PrincipalDepartment", False)
      Dim vDept As String = vPrincDept.Text
      vDept = vPrincDept.Text
      Dim vList As New ParameterList(True)
      If vDept.Length > 0 Then
        'restrict principal users combo box
        vPrincDeptLogname.FillComboWithRestriction(vDept, "", False, vList)
      Else
        'all users should be in principal dept logname combo box
        vPrincDeptLogname.Text = ""
        vPrincDeptLogname.FillComboBox(vList)
      End If
      If mvParams.ContainsKey("PrincipalDepartmentLogname") Then
        vPrincDeptLogname.Text = mvParams("PrincipalDepartmentLogname")
      End If
    End If

    'disable primary key fields in edit mode
    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      For Each vPanelItem As PanelItem In mvEditPanelInfo.PanelItems
        If vPanelItem.ReadOnlyItem Then
          epl.EnableControl(vPanelItem.ParameterName, False)
        End If
      Next
    End If

    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      Select Case mvTableName.ToLower
        Case "contact_groups", "event_groups"
          epl.EnableControl("Client", False)    'Disable Client code
        Case "organisation_groups"
          epl.EnableControl("Client", False)    'Disable Client code
          If epl.GetValue("OrganisationGroup").Equals("ORG", StringComparison.InvariantCultureIgnoreCase) Then epl.EnableControl("ViewInContactCard", False)
        Case "payment_frequencies"
          SetPaymentFrequencyOffset(False)
      End Select
    End If

    epl.SetFocus()

    'Set Form height
    Dim vDiff As Integer = epl.RequiredHeight - epl.Height
    Me.Height += vDiff

    'Set a minimum width for the notes so that it has enough space to be displayed correctly
    If txtNotes.Width < 150 Then vDiff = 150 - txtNotes.Width 'Minimum width for notes currently set to 150 
    vDiff += epl.RequiredWidth - spc.Panel1.Width
    'Set Form width
    Me.Width += vDiff
    'Adjust the first panel to where the edit panel for the control ends.
    'The remaining space will be used to display the notes
    spc.SplitterDistance = epl.RequiredWidth

    If mvTableName.Equals("packages", StringComparison.InvariantCultureIgnoreCase) Then
      SetPackageTableFieldAvailabilty()
      AddHandler epl.ValueChanged, AddressOf HandlePackageTableChange
    End If

    epl.DataChanged = False
  End Sub

  Private Sub CloseForm()
    If mvRequiresRefresh Then Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

#End Region

  Private Function CheckMembershipGroups() As Boolean
    'Ask question whether MembershipGroups data should be updated
    Dim vOrgGroup As String = String.Empty
    Dim vHistoric As Boolean
    Dim vMemberTypeCode As String = String.Empty
    Dim vQuestion As String = String.Empty
    Dim vUseMembershipGroups As Boolean
    Dim vSave As Boolean = True

    vUseMembershipGroups = False
    If mvTableName = "membership_types" Then
      vUseMembershipGroups = AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.organisation_group).Length > 0
    Else
      vUseMembershipGroups = True
    End If

    If vUseMembershipGroups Then
      If SetMembershipGroups(vQuestion, vOrgGroup, vMemberTypeCode, vHistoric) Then
        vSave = ShowQuestion(vQuestion, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes
      End If
    End If

    Return vSave
  End Function

  Private Function SetMembershipGroups(ByRef pQuestion As String, ByRef pOrgGroup As String, ByRef pMemberType As String, ByRef pHistoric As Boolean) As Boolean
    'Set up variables etc. for asking whether to update and to actually perform the update
    Dim vAskQuestion As Boolean

    If mvTableName = "membership_controls" Then
      Dim vOrgGroup As String = epl.GetOptionalValue("OrganisationGroup")
      Dim vOriginalOrgGroup As String = mvParams.ValueIfSet("OrganisationGroup")

      If vOrgGroup <> vOriginalOrgGroup Then
        'Organisation Group has changed
        If vOrgGroup.Length > 0 Then
          'If the original value was unset then add MembershipGroups
          If vOriginalOrgGroup.Length = 0 Then
            pQuestion = QuestionMessages.QmOrganisationGroupSet    'The Organisation Group has been set so Membership Groups data will now be created. & vbCrLf & Do you wish to continue?
            pOrgGroup = epl.GetValue("OrganisationGroup")
            vAskQuestion = True
          End If
        Else
          'The OrganisationGroup has been unset so make MembershipGroups historic
          pQuestion = QuestionMessages.QmOrganisationGroupUnset    'The Organisation Group has been unset so Membership Groups data will now be made historic. & vbCrLf & Do you wish to continue?
          pHistoric = True
          vAskQuestion = True
        End If
      End If
    ElseIf mvTableName = "membership_types" Then
      Dim vBranchMem As String = epl.GetOptionalValue("BranchMembership")
      Dim vOriginalBranchMem As String = mvParams.ValueIfSet("BranchMembership")

      If vBranchMem <> vOriginalBranchMem Then
        'BranchMembership flag has changed
        If BooleanValue(epl.GetValue("BranchMembership")) Then
          'BranchMembership flag is set so add MembershipGroups
          pQuestion = QuestionMessages.QmBranchMembershipSet    'The Branch Membership flag has been set so Membership Groups data will now be created. & vbCrLf & Do you wish to continue?
          vAskQuestion = True
        Else
          'The BranchMembership flag has been unset so make MembershipGroups historic
          pQuestion = QuestionMessages.QmBranchMembershipUnset    'The Branch Membership flag has been unset so Membership Groups data will now be made historic. & vbCrLf & Do you wish to continue?
          pHistoric = True
          pMemberType = epl.GetValue("MembershipType")
          vAskQuestion = True
        End If
      End If
    End If

    Return vAskQuestion
  End Function


  Private Sub cmdAddMore_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddMore.Enter
    txtNotes.Text = String.Empty
  End Sub

  Private Sub cmdOK_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Enter
    txtNotes.Text = String.Empty
  End Sub

  Private Sub cmdCancel_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Enter
    txtNotes.Text = String.Empty
  End Sub

  Private Sub frmTableEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    mvStockFlagBefore = 1     'Initial Flag for Stock Item. 0 = Stock Item flag is 'N' , 1 = Stock Item flag is 'Y'
    mvStockFlagAfter = 0      'Initial Flag for Stock Item. 0 = Stock Item flag is not changed , 1 = Stock Item flag is changed
    InitForm()
  End Sub

  ''' <summary>
  ''' Get the formatted date/time value
  ''' </summary>
  ''' <param name="pParamName"></param>
  ''' <param name="pFormat"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetFormattedValue(ByVal pParamName As String, ByVal pFormat As String) As String
    Dim vValue As String = String.Empty
    If mvParams.ContainsKey(pParamName) Then
      vValue = DateValue(mvParams(pParamName)).ToString(pFormat)
    End If
    Return vValue
  End Function

  Private Sub RemoveLookupRestrictions()
    With mvEditPanelInfo
      Select Case mvTableName
        Case "contact_controls"
          If .PanelItems.Exists("ContactParentRelationship") Then .PanelItems("ContactParentRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("PrimaryRelationship") Then .PanelItems("PrimaryRelationship").RemoveLookupRestriction = True
        Case "financial_controls"
          If .PanelItems.Exists("InmemoriamRelationship") Then .PanelItems("InmemoriamRelationship").RemoveLookupRestriction = True
        Case "gaye_controls"
          If .PanelItems.Exists("EmployerAgencyRelationship") Then .PanelItems("EmployerAgencyRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("EmployerPayrollRelationship") Then .PanelItems("EmployerPayrollRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("PostTaxEmprPayrollRelation") Then .PanelItems("PostTaxEmprPayrollRelation").RemoveLookupRestriction = True
        Case "legacy_controls"
          If .PanelItems.Exists("JointLegatorRelationship") Then .PanelItems("JointLegatorRelationship").RemoveLookupRestriction = True
        Case "marketing_controls"
          If .PanelItems.Exists("DerivedToJointRelationship") Then .PanelItems("DerivedToJointRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("DerivedToDerivedRelationshi") Then .PanelItems("DerivedToDerivedRelationshi").RemoveLookupRestriction = True
        Case "membership_controls"
          If .PanelItems.Exists("RealToJointRelationship") Then .PanelItems("RealToJointRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("RealToRealRelationship") Then .PanelItems("RealToRealRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("BranchParentRelationship") Then .PanelItems("BranchParentRelationship").RemoveLookupRestriction = True
        Case "membership_types"
          If .PanelItems.Exists("Relationship") Then .PanelItems("Relationship").RemoveLookupRestriction = True
        Case "organisation_groups"
          If .PanelItems.Exists("PrimaryRelationship") Then .PanelItems("PrimaryRelationship").RemoveLookupRestriction = True
        Case "relationship_group_details"
          If .PanelItems.Exists("Relationship") Then .PanelItems("Relationship").RemoveLookupRestriction = True
        Case "relationships"
          If .PanelItems.Exists("ParentRelationship") Then .PanelItems("ParentRelationship").RemoveLookupRestriction = True
          If .PanelItems.Exists("ComplimentaryRelationship") Then .PanelItems("ComplimentaryRelationship").RemoveLookupRestriction = True
        Case "service_controls"
          If .PanelItems.Exists("ModifierRelationship") Then .PanelItems("ModifierRelationship").RemoveLookupRestriction = True
        Case "vat_rate_history"
          If .PanelItems.Exists("VatRate") Then .PanelItems("VatRate").RemoveLookupRestriction = True
      End Select
    End With
  End Sub

  Private Sub InitForm()
    If mvTableName = "performances" Then
      Dim vTextLookupBox As TextLookupBox = epl.FindTextLookupBox("ExpenditureGroup")
      If Not vTextLookupBox Is Nothing Then vTextLookupBox.MultipleValuesSupported = True
    ElseIf mvTableName = "contact_controls" Then
      Dim vTextLookupBox As TextLookupBox = epl.FindTextLookupBox("TylSuppressionExclusionList")
      If Not vTextLookupBox Is Nothing Then vTextLookupBox.MultipleValuesSupported = True
    ElseIf mvTableName = "ownership_group_users" Then
      epl.FindDateTimePicker("ValidFrom").MinDate = Today
    End If

    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso mvTableName <> "currency_codes" Then
      epl.SetValue("CurrencyCode", AppValues.ControlValue(AppValues.ControlValues.currency_code), , , False)
    End If

    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso mvTableName = "membership_types" Then
      epl.SetValue("AllowAsFirstType", "Y")
    End If

    'Read the original values from the param list as the date values that are set through the edit panel
    'wont be set up correctly
    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      If mvTableName = "gift_aid_controls" Then
        epl.SetValue("AccountingPeriodStart", GetFormattedValue("AccountingPeriodStart", "dd/MM"))
        epl.SetValue("TaxYearStart", GetFormattedValue("TaxYearStart", "dd/MM"))
      ElseIf mvTableName = "financial_controls" Then
        epl.SetValue("LoanCapitalisationDate", GetFormattedValue("LoanCapitalisationDate", "dd/MM"))
      ElseIf mvTableName = "workstream_groups" Then
        epl.EnableControl("WorkstreamGroup", False)
      ElseIf mvTableName = "workstream_group_outcomes" Then
        epl.EnableControl("WorkstreamGroupOutcome", False)
        epl.EnableControl("WorkstreamGroup", False)
      ElseIf mvTableName = "workstream_group_actions" Then
        epl.EnableControl("MasterAction", False)
        epl.EnableControl("WorkstreamGroup", False)
      End If
    End If

    'Disable the ReadyForPayment checkbox while amending authorisation statuses
    If mvTableName = "authorisation_statuses" Then
      If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
        epl.EnableControl("ReadyForPayment", False)
      End If
    End If

    'Disable the UniqueId textbox at all times as this is a readonly field
    If mvTableName = "contact_groups" Or mvTableName = "organisation_groups" Or mvTableName = "event_groups " Then epl.EnableControl("LastUsedId", False)

    If mvTableName = "products" And mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      'Enable PackProduct checkbox only if the prod is a stock item
      epl.EnableControl("PackProduct", False) 'Always start with Pack Product disabled
      If epl.GetOptionalValue("StockItem") = "Y" Then
        epl.EnableControl("LastStockCount", False)
        epl.EnableControl("PackProduct", True) 'Enable for a Stock Product
      Else
        mvStockFlagBefore = 0
      End If
    End If

    'Disable eligible for gift aid checkbox while opening the form if the values is unchecked
    If mvTableName = "products" AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
      'Always start with Pack Product unchecked & disabled
      epl.SetValue("PackProduct", "N")
      epl.EnableControl("PackProduct", False)
    End If

    If mvTableName = "rates" AndAlso mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmSelect Then
      Dim vproduct As String = epl.GetOptionalValue("Product")
      If vproduct.Length > 0 Then
        Dim vlist As New ParameterList(True)
        vlist("Product") = vproduct
        Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vlist)
        If vDataRow("Course").ToString <> "Y" Then
          epl.SetValue("FixedPrice", "N")
          epl.EnableControl("FixedPrice", False)
        Else
          epl.EnableControl("FixedPrice", True)
        End If
      End If
    End If

    'warehouse and last stock count fields cannot be amended
    If mvTableName = "product_warehouses" AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      epl.EnableControl("Warehouse", False)
      epl.EnableControl("LastStockCount", False)
    End If

    'if the mailing template exists in mailing_template_documents then disable ExplicitSelection and set value to checked
    If mvTableName = "mailing_templates" AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
      Dim vMailingTemplate As String = epl.GetOptionalValue("MailingTemplate")
      Dim vList As New ParameterList(True)
      vList("MailingTemplate") = vMailingTemplate
      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMailingTemplateDocuments, vList) > 0 Then
        epl.SetValue("ExplicitSelection", "Y")
        epl.EnableControl("ExplicitSelection", False)
      End If
    End If

    'in add mode disable all the controls. 
    'in edit mode leave notes enabled and disable the remaining controls
    If mvTableName = "mailing_history" Then
      For Each vItem As PanelItem In mvEditPanelInfo.PanelItems
        If (vItem.AttributeName <> "notes" AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend) _
         OrElse mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
          epl.EnableControl(vItem.ParameterName, False)
        End If
      Next
      cmdAddMore.Enabled = mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew
      cmdOK.Enabled = mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew
    End If

    'Set Default duration to the value set up in the config
    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso mvTableName = "cpd_cycle_types" Then
      epl.SetValue("DefaultDuration", AppValues.ConfigurationValue(AppValues.ConfigurationValues.cpd_cycle_default_duration))
    End If

    If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmSelect Then
      Select Case mvTableName
        Case "branches"
          If epl.GetOptionalValue("Historical") = "N" Then epl.EnableControl("Historical", False)
        Case "devices"
          If epl.GetOptionalValue("Email") = "N" Then
            epl.SetValue("AutoEmail", "N")
            epl.EnableControl("AutoEmail", False)
          End If
        Case "event_child_discount_levels"
          epl.EnableControl("EventChildDiscountNumber", False)
        Case "event_extra_fee_multipliers"
          epl.EnableControl("EventExtraFeeNumber", False)
        Case "event_fee_band_discounts"
          epl.EnableControl("EventFeeBandDiscountNumber", False)
        Case "event_fees"
          epl.EnableControl("EventFeeNumber", False)
        Case "internal_resources"
          epl.EnableControl("ResourceNumber", False)
        Case "ownership_access_levels"
          epl.EnableControl("OwnershipAccessLevel", False)
        Case "service_control_restrictions"
          epl.EnableControl("ServiceRestrictionNumber", False)
        Case "service_control_start_days", "service_start_days"
          epl.EnableControl("StartDayNumber", False)
        Case "sources"
          epl.EnableControl("SourceNumber", False)
        Case "venues"
          If epl.GetValue("OrganisationNumber").Length = 0 Then epl.EnableControl("ContactNumber", False)
        Case "survey_answers"
          epl.EnableControl("SurveyAnswerNumber", False)
        Case "survey_questions"
          epl.EnableControl("SurveyQuestionNumber", False)
        Case "survey_versions"
          epl.EnableControl("SurveyVersionNumber", False)
        Case "surveys"
          epl.EnableControl("SurveyNumber", False)
        Case "rate_modifiers"
          epl.EnableControl("RateModifierNumber", False)
        Case "web_documents"
          epl.EnableControl("WebDocumentNumber", False)
          epl.EnableControl("DownloadCount", False)
          epl.EnableControl("LastDownloadedOn", False)
        Case "exam_accreditation_statuses"
          epl.EnableControl("AccreditationStatusId", False)
      End Select
    End If
    If mvTableName <> "rate_modifiers" Then
      SetPriceLimitFields("CurrentPrice")
      SetPriceLimitFields("FuturePrice")
    End If

    If mvTableName = "mailings" Then
      epl.EnableControl("BulkMailerMailing", False)
      epl.EnableControl("BulkMailerStatisticsDate", False)
    End If
  End Sub

  Private Sub SetPriceLimitFields(ByVal pParamName As String)
    Dim vString As String

    If pParamName = "CurrentPrice" Then
      vString = "Current"
    Else
      vString = "Future"
    End If

    If Val(epl.GetOptionalValue(pParamName)) = 0 Then
      epl.EnableControl(vString & "PriceLowerLimit", True)
      epl.EnableControl(vString & "PriceUpperLimit", True)
    Else
      epl.EnableControl(vString & "PriceLowerLimit", False)
      epl.SetValue(vString & "PriceLowerLimit", String.Empty)

      epl.EnableControl(vString & "PriceUpperLimit", False)
      epl.SetValue(vString & "PriceUpperLimit", String.Empty)
    End If
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pValue As System.String) Handles epl.ValueChanged
    Dim vControl As Control = epl.FindPanelControl(pParameterName)
    If TypeOf vControl Is CheckBox Then
      Static vInhibit As Boolean
      Dim vUsedElseWhere As Boolean
      Dim vChecked As Boolean
      Dim vPanelItem As PanelItem
      Dim vCheckBox As CheckBox = DirectCast(vControl, CheckBox)

      If vInhibit Then
        vInhibit = False
      Else
        'Ignore product rate can only be checked when Incentive Type is I
        vCheckBox = epl.FindCheckBox(pParameterName)
        vPanelItem = DirectCast(vCheckBox.Tag, PanelItem)
        If vPanelItem.AttributeName = "ignore_product_and_rate" Then
          If epl.GetOptionalValue("IncentiveType") <> "I" AndAlso vCheckBox.Checked Then
            vCheckBox.Checked = False
            Beep()
          End If
        End If

        If vControl Is epl.ActiveControl Then
          vChecked = vCheckBox.Checked
          If vPanelItem.TableName = "devices" Then
            'Auto email will be enabled/disabled based on the value of email
            If vPanelItem.AttributeName = "email" Then
              epl.EnableControl("AutoEmail", vChecked)
              If Not vChecked Then epl.SetValue("AutoEmail", "N")
              epl.SetValue("WwwAddress", "N")
            ElseIf vPanelItem.AttributeName = "www_address" Then
              vCheckBox = epl.FindCheckBox("Email")
              If Not vCheckBox Is Nothing Then
                vCheckBox.Checked = False
                epl.EnableControl("AutoEmail", False)
                epl.SetValue("AutoEmail", "N")
              End If
            End If
          ElseIf vPanelItem.TableName = "products" Then
            ' check only if amending an existing product,not when creating a new one. 
            If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso (vPanelItem.AttributeName = "donation" _
            OrElse vPanelItem.AttributeName = "subscription" OrElse vPanelItem.AttributeName = "stock_item" _
            OrElse vPanelItem.AttributeName = "course" OrElse vPanelItem.AttributeName = "accommodation" _
            OrElse vPanelItem.AttributeName = "postage_packing" OrElse vPanelItem.AttributeName = "uses_product_numbers" _
            OrElse vPanelItem.AttributeName = "sponsorship_event" OrElse vPanelItem.AttributeName = "eligible_for_gift_aid" _
            OrElse vPanelItem.AttributeName = "accrues_interest" OrElse vPanelItem.AttributeName = "exam") Then
              Dim vList As New ParameterList(True)
              vList("Product") = epl.GetValue("Product")
              vList("MaintenanceTableName") = "products"
              Try
                DataHelper.CheckUsedElsewhere(vList)
              Catch vEx As CareException
                If vEx.ErrorNumber = CareException.ErrorNumbers.enRecordCannotBeChanged Then
                  ShowWarningMessage(vEx.Message)
                  vInhibit = True
                  vCheckBox.Checked = Not vCheckBox.Checked 'Invert the value of the checkbox
                End If
              End Try
            ElseIf mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso vPanelItem.AttributeName = "pack_product" Then
              'Only need to check PackedProducts & StockMovements table
              Dim vProduct As TextBox = epl.FindTextBox("Product")
              If vProduct IsNot Nothing Then
                Dim vList As New ParameterList(True)
                vList("Product") = vProduct.Text
                If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPackedProductsProduct, vList) > 0 Then
                  ShowWarningMessage(InformationMessages.ImRecordCannotBeChanged, "Packed Products", "Product")  '%s refer to this %s & vbCrLf & vbCrLf & Record cannot be changed
                  vUsedElseWhere = True
                ElseIf DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPackedProductsLinkProduct, vList) > 0 Then
                  ShowWarningMessage(InformationMessages.ImProductPartOfPackProduct)  'This Product is part of a Pack Product & vbCrLf & vbCrLf & Record cannot be changed
                  vUsedElseWhere = True
                ElseIf DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctStockMovementsProduct, vList) > 0 Then
                  ShowWarningMessage(InformationMessages.ImRecordCannotBeChanged, "Stock Movements", "Product")  '%s refer to this %s & vbCrLf & vbCrLf & Record cannot be changed
                  vUsedElseWhere = True
                End If
                If vUsedElseWhere Then
                  vInhibit = True
                  vCheckBox.Checked = Not vCheckBox.Checked 'Invert the value of the checkbox
                End If
              End If
            End If

            If vPanelItem.AttributeName = "stock_item" Then
              Dim vPackProd As CheckBox = epl.FindCheckBox("PackProduct")
              If Not vPackProd Is Nothing Then
                If vCheckBox.Checked Then
                  'Stock Product - enable Pack Product
                  vPackProd.Enabled = True
                  vPackProd.Checked = False
                  If mvStockFlagBefore = 0 Then
                    mvStockFlagAfter = 1
                  Else
                    mvStockFlagAfter = 0
                  End If
                Else
                  'Not a Stock Product - disable Pack Product
                  vPackProd.Checked = False
                  vPackProd.Enabled = False
                End If
              End If
            End If
          ElseIf vPanelItem.TableName = "rates" Then
            If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso vPanelItem.AttributeName = "concessionary" Then
              Dim vProduct As TextLookupBox = epl.FindTextLookupBox("Product")
              Dim vRate As TextBox = epl.FindTextBox("Rate")
              If Not vProduct Is Nothing AndAlso Not vRate Is Nothing Then
                Dim vList As New ParameterList(True)
                vList("Product") = vProduct.Text
                vList("Rate") = vRate.Text
                If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMembershipPrices, vList) > 0 Then
                  'Changing the concessionary flag could invalidate the MembershipPrices data so prevent it's change
                  ShowWarningMessage(InformationMessages.ImRecordCannotBeChanged, "Membership Prices", "Rate") '%s refer to this %s & vbCrLf & vbCrLf & Record cannot be changed
                  vInhibit = True
                  vCheckBox.Checked = Not vCheckBox.Checked 'Invert the value of the checkbox
                End If
              End If
            End If
          ElseIf vPanelItem.TableName = "purchase_order_types" Then
            If vChecked Then
              If pParameterName = "RegularPayments" Then
                epl.SetValue("PaymentSchedule", "N", pReportError:=False)
                epl.SetValue("AdHocPayments", "N", pReportError:=False)
              ElseIf pParameterName = "AdHocPayments" Then
                epl.SetValue("PaymentSchedule", "N", pReportError:=False)
                epl.SetValue("RegularPayments", "N", pReportError:=False)
              ElseIf pParameterName = "PaymentSchedule" Then
                epl.SetValue("AdHocPayments", "N", pReportError:=False)
                epl.SetValue("RegularPayments", "N", pReportError:=False)
              End If
            End If
          ElseIf mvTableName = "membership_controls" AndAlso pParameterName = "AdvancedCmt" Then
            epl.SetErrorField(pParameterName, "")   'Clear any error
          End If
        ElseIf vPanelItem.AttributeName = "stock_item" Then
          mvStockFlagBefore = 0
          mvStockFlagAfter = 0
        End If
      End If
    ElseIf TypeOf vControl Is TextBox Then
      Dim vTextBox As TextBox = DirectCast(vControl, TextBox)

      ' if current price is zero then allow the entry of lower and upper limits for current price
      If mvTableName <> "rate_modifiers" AndAlso (pParameterName = "CurrentPrice" OrElse pParameterName = "FuturePrice") Then
        SetPriceLimitFields(pParameterName)
      End If

      If pParameterName = "MembershipCardDuration" Then
        If epl.GetOptionalValue("MembershipCard") = "N" Then
          vTextBox.Text = ControlText.TxtZero
        End If
      End If

      Select Case mvTableName
        Case "bank_account_claim_days"
          If pParameterName.Equals("ClaimDay") Then epl.SetErrorField("ClaimDay", String.Empty)

        Case "payment_frequencies"
          Select Case pParameterName
            Case "Frequency", "Interval"
              SetPaymentFrequencyOffset(True)
            Case "OffsetMonths"
              epl.SetErrorField(pParameterName, "")
          End Select
      End Select

    ElseIf TypeOf vControl Is TextLookupBox Then
      Dim vTextLookupBox As TextLookupBox = DirectCast(vControl, TextLookupBox)
      If vTextLookupBox.IsValid AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso mvTableName = "rates" AndAlso pParameterName = "Product" Then
        If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.nominal_account_validation) Then
          Dim vNominal As String
          Dim vSubsNominal As String
          Dim vRestriction As String

          If pValue.Length > 0 Then
            ''Need to refresh the nominal account suffix combos due to change of product
            Dim vList As New ParameterList(True)
            vList("Product") = vTextLookupBox.Text
            Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
            vNominal = vDataRow("NominalAccount").ToString
            vSubsNominal = vDataRow("SubsequentNominalAccount").ToString

            Dim vTextLookup As TextLookupBox = epl.FindTextLookupBox("NominalAccountSuffix", False)
            If Not vTextLookup Is Nothing Then
              vList = New ParameterList(True)
              vList("ProductNominalAccount") = vNominal
              vList("Active") = "Y"
              vTextLookup.FillComboWithRestriction(vList)
            End If

            vTextLookup = epl.FindTextLookupBox("SubsequentNominalSuffix", False)
            If Not vTextLookup Is Nothing Then
              vList = New ParameterList(True)
              vList("ProductNominalAccount") = vSubsNominal
              vRestriction = "product_nominal_account = '" & vSubsNominal & "' AND history_only = 'N'"
              vList("Active") = "Y"
              vTextLookup.FillComboWithRestriction(vList)
            End If
          End If
        End If

        If pValue.Length > 0 Then
          Dim vList As New ParameterList(True)
          vList("Product") = vTextLookupBox.Text
          Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
          If vDataRow("Course").ToString <> "Y" Then
            epl.EnableControl("FixedPrice", False)
            epl.SetValue("FixedPrice", "N")
          Else
            epl.EnableControl("FixedPrice", True)
          End If
        End If
      End If

      If vTextLookupBox.IsValid AndAlso mvTableName = "ownership_groups" AndAlso pParameterName = "PrincipalDepartment" AndAlso epl.DataChanged Then
        Dim vPrincDeptLogname As TextLookupBox = epl.FindTextLookupBox("PrincipalDepartmentLogname", False)
        Dim vPrincDept As TextLookupBox = epl.FindTextLookupBox("PrincipalDepartment", False)
        Dim vDept As String = vControl.Text
        vDept = vPrincDept.Text
        Dim vList As New ParameterList(True)
        If vDept.Length > 0 Then
          'restrict principal users combo box
          vPrincDeptLogname.Text = ""
          vPrincDeptLogname.FillComboWithRestriction(vDept, "", False, vList)
        Else
          'all users should be in principal dept logname combo box
          vPrincDeptLogname.Text = ""
          vPrincDeptLogname.FillComboBox(vList)
        End If
      End If

      If vTextLookupBox.IsValid AndAlso mvTableName = "membership_prices" AndAlso pParameterName = "MembershipType" AndAlso pValue.Length > 0 AndAlso epl.DataChanged Then
        Dim vTextLookup As TextLookupBox = epl.FindTextLookupBox("Rate", False)
        If Not vTextLookup Is Nothing Then
          Dim vList As New ParameterList(True)
          vList("MembershipType") = vTextLookupBox.Text
          Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
          vList = New ParameterList(True)
          vList("Product") = vDataRow("FirstPeriodsProduct").ToString
          vTextLookup.FillComboWithRestriction(vList)
        End If
      End If

      Select Case mvTableName
        Case "contact_alerts"
          If vTextLookupBox.IsValid AndAlso pParameterName.Equals("ContactAlertType") Then
            If pValue.Equals("C", StringComparison.InvariantCultureIgnoreCase) Then
              epl.SetValue("ContactAlertMessageType", "W", True)
            Else
              epl.EnableControl("ContactAlertMessageType", True)
            End If
          End If
        Case "membership_controls"
          If pParameterName = "CmtCalcProportionalBalance" Then
            If pValue = "N" Then epl.SetValue("AdvancedCmt", "N")
            epl.EnableControl("AdvancedCmt", (pValue <> "N"))
          End If
        Case "membership_entitlement", "membership_types"
          If pParameterName = "CmtProrateOldCosts" AndAlso pValue = "F" AndAlso vTextLookupBox.OriginalText <> "F" Then
            epl.SetValue("CmtProrateNewCosts", "N")
          End If
        Case "venues"
          If vTextLookupBox.IsValid AndAlso epl.GetValue("OrganisationNumber").Length > 0 Then
            epl.EnableControl("ContactNumber", True)
          Else
            epl.SetValue("ContactNumber", String.Empty, True)
          End If
        Case "purchase_order_controls"
          If vTextLookupBox.IsValid AndAlso epl.PanelInfo.PanelItems.Exists("PoPaymentType") AndAlso
            epl.GetValue("PoPaymentType").Length > 0 AndAlso epl.PanelInfo.PanelItems.Exists("DistributionCode") Then

            Dim vDistributionCode As String = vTextLookupBox.GetDataRowItem("DistributionCode")
            If vDistributionCode.Length > 0 Then epl.SetValue("DistributionCode", vDistributionCode)

          End If
        Case "gift_aid_controls"
          If vTextLookupBox.IsValid AndAlso pValue = "O" AndAlso epl.PanelInfo.PanelItems.Exists("SubmitterContact") AndAlso
            epl.GetValue("SubmitterContact").Length = 0 Then
            epl.SetErrorField("SubmitterContact", InformationMessages.ImFieldMandatory)

          Else
            epl.SetErrorField("SubmitterContact", "")
          End If
        Case "activity_cpd_points"
          If pParameterName.Equals("CpdCycleType", StringComparison.InvariantCultureIgnoreCase) Then
            Dim vCPDCategoryTypeTLB As TextLookupBox = epl.FindTextLookupBox("CpdCategoryType", False)
            If vCPDCategoryTypeTLB IsNot Nothing Then
              Dim vFilter As String = If(pValue.Length > 0, "CpdCycleType = '" & pValue & "'", String.Empty)
              vCPDCategoryTypeTLB.SetFilter(vFilter, True, True)
            End If
          ElseIf pParameterName.Equals("CpdCategoryType", StringComparison.InvariantCultureIgnoreCase) Then
            epl.FindTextLookupBox("CpdCategory").FillComboWithRestriction(pValue)
          End If
        Case "payment_frequencies"
          If pParameterName.Equals("period", StringComparison.InvariantCultureIgnoreCase) AndAlso pValue.Length > 0 Then
            SetPaymentFrequencyOffset(True)
          End If
        Case "bank_account_claim_days"
          If pParameterName.Equals("ClaimType") Then epl.SetErrorField("ClaimDay", String.Empty)   'Clear any ClaimDay error
      End Select
    ElseIf TypeOf vControl Is MaskedTextBox Then
      Dim vTextBox As MaskedTextBox = DirectCast(vControl, MaskedTextBox)
      If pParameterName = "IbanNumber" AndAlso vTextBox.Text.Length > 0 Then
        Try
          epl.SetErrorField("IbanNumber", "")
          DataHelper.CheckIbanNumber(vTextBox.Text)
        Catch vException As Exception
          epl.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
        End Try
      End If

    End If
  End Sub

  Function ValidatePostcode(ByVal pPostcode As String, ByVal pParameterName As String) As Boolean
    Dim vValid As Boolean = True
    pPostcode = pPostcode.Trim
    If pPostcode.Contains("  ") Then
      Do
        pPostcode = pPostcode.Replace("  ", " ")
      Loop While pPostcode.Contains("  ")
    End If

    If pPostcode <> String.Empty Then
      If AppValues.DefaultCountryCode = "UK" Then 'In VB6 this used to check IsDefaultCountryUK. 
        vValid = epl.ValidatePostcodeFormat(pPostcode)
        If vValid Then
          Try
            Dim vList As New ParameterList(True)
            vList("Postcode") = pPostcode
            'This will raise an error if the postcode does not exist in mailsort database
            DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMailsortCode, vList)
          Catch vEx As CareException
            If vEx.ErrorNumber = CareException.ErrorNumbers.enPostcodeValidationError Then
              epl.SetErrorField(pParameterName, vEx.Message)             'Postcode not listed in Mailsort Database
              vValid = False
            End If
          End Try
        Else
          epl.SetErrorField(pParameterName, InformationMessages.ImInvalidPostCode)    'Invalid Postcode
        End If
      End If
    End If

    Return vValid
  End Function

  Private Sub ValidateAllItems(ByVal pSender As Object, ByVal pList As ParameterList, ByRef pValid As Boolean) Handles epl.ValidateAllItems
    If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmSelect Then
      Dim vAttr As String
      Dim vSQLCheck As Boolean
      Dim vValue As String = String.Empty
      Dim vLen1 As Integer
      Dim vLen2 As Integer
      Dim vChanged As Boolean
      Dim vIRProductSet As Boolean
      Dim vIRRateSet As Boolean
      Dim vIRContactSet As Boolean
      Dim vDate1 As String
      Dim vDate2 As String
      Dim vPanelItem As PanelItem
      Dim vList As ParameterList
      Dim vParamValue As String = String.Empty

      pValid = True

      Select Case mvTableName
        Case "custom_forms", "custom_form_controls", "journal_types", "journal_type_controls", "custom_data_set_details"
          vSQLCheck = True
      End Select

      For Each vPanelItem In epl.PanelInfo.PanelItems
        If vPanelItem.ControlType <> PanelItem.ControlTypes.ctReadOnly AndAlso vPanelItem.ControlType <> PanelItem.ControlTypes.ctTab Then
          vParamValue = epl.GetValue(vPanelItem.ParameterName)
          If vPanelItem.ValidationTable.Length > 0 Then
            'If the field is blank but is a lookup dependant on another field
            'and the other field is not blank then it is mandatory
            'This would normally be the case where for example: it is not valid to have a rate without a product
            If pValid AndAlso vPanelItem.RestrictionAttribute.Length > 0 Then
              Dim vSkipVal As Boolean = False
              Dim vControl As Control = epl.FindPanelControl(ProperName(vPanelItem.RestrictionAttribute), False)
              If vControl Is Nothing Then
                vSkipVal = True
              End If
              If Not vSkipVal Then
                vLen1 = vParamValue.Length
                vLen2 = epl.GetValue(ProperName(vPanelItem.RestrictionAttribute)).Length
                If ((vLen1 > 0) AndAlso (vLen2 = 0)) OrElse ((vLen2 > 0) AndAlso (vLen1 = 0)) Then
                  Select Case AttributeName(vPanelItem.ParameterName)
                    Case "from_attribute", "to_attribute", "subsidiary_attribute", "subsidiary_validation_attribute", "activity_value"
                      'OK for the restriction attribute to be blank
                    Case "branch_rate"
                      If epl.GetValue("BranchMembership") = "Y" Then pValid = False
                    Case Else
                      pValid = False
                  End Select
                  If Not pValid Then
                    If vLen1 = 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFieldMustNotBeBlank) 'Field must Not be Blank
                    Else
                      epl.SetErrorField(ProperName(vPanelItem.RestrictionAttribute), InformationMessages.ImFieldMustNotBeBlank) 'Field must Not be Blank
                    End If
                  End If
                End If
              End If
            End If
          End If

          If pValid Then
            If pValid AndAlso vSQLCheck Then
              If vPanelItem.AttributeName.Contains("sql") Then
                If vParamValue.Contains("?") Then
                  ShowWarningMessage(InformationMessages.ImQuestionMarksInSQLStatements) 'Question marks in SQL statements should be changed to hash symbols '#' before saving
                  pValid = False
                End If
              End If
            End If

            If pValid AndAlso vParamValue.Length > 0 AndAlso vPanelItem.AttributeName = "sundry_credit_trans_type" OrElse vPanelItem.AttributeName = "reversal_transaction_type" Then
              vList = New ParameterList(True)
              vList("TransactionType") = vParamValue
              vList("TransactionSign") = "D"
              vList("NegativesAllowed") = "Y"

              If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctTransactionTypes, vList) = 0 Then
                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMustBeADebitTypeTransaction) 'Must be a debit-type Transaction Type that allows negative values
                pValid = False
              End If
            End If

            If pValid Then
              Select Case mvTableName
                Case "bank_account_claim_days"
                  If vPanelItem.AttributeName.Equals("claim_day") Then
                    If epl.GetValue("ClaimType").Equals("DD", StringComparison.InvariantCultureIgnoreCase) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_dd_fixed_claim_date, False) Then
                      If IntegerValue(vParamValue) > 28 Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImInvalidDDClaimDay, True)
                      End If
                    End If
                  End If

                Case "banks"
                  If vPanelItem.AttributeName = "postcode" Then
                    pValid = ValidatePostcode(vParamValue, vPanelItem.ParameterName)
                  End If

                Case "contact_controls"
                  Select Case vPanelItem.AttributeName
                    Case "start_of_day", "end_of_day", "start_of_lunch", "end_of_lunch"
                      'Set lengths of start/end_of_lunch - they could be null values
                      vAttr = epl.GetValue("StartOfLunch")
                      vLen1 = vAttr.Trim.Length
                      If vAttr = ":" Then vLen1 = 0
                      vAttr = epl.GetValue("EndOfLunch")
                      vLen2 = vAttr.Trim.Length
                      If vAttr = ":" Then vLen2 = 0
                  End Select

                  If vPanelItem.AttributeName = "start_of_day" Then
                    If epl.GetValue("EndOfDay").Trim.Length > 1 Then
                      If TimeValue(epl.GetValue("StartOfDay")) >= TimeValue(epl.GetValue("EndOfDay")) Then
                        pValid = False
                        vValue = InformationMessages.ImStartOfDayTimeMustBeBeforeTheEndOfDayTime    'The Start Of Day time must be before the End Of Day time
                      End If
                    ElseIf vLen1 > 0 Then 'start_of_lunch
                      If TimeValue(epl.GetValue("StartOfDay")) >= TimeValue(epl.GetValue("StartOfLunch")) Then
                        pValid = False
                        vValue = InformationMessages.ImStartOfDayTimeMustBeBeforeTheStartOfLunchTime    'The Start Of Day time must be before the Start Of Lunch time
                      End If
                    ElseIf vLen2 > 0 Then 'end_of_lunch
                      If TimeValue(epl.GetValue("StartOfDay")) >= TimeValue(epl.GetValue("EndOfLunch")) Then
                        pValid = False
                        vValue = InformationMessages.ImStartOfDayTimeMustBeBeforeTheEndOfLunchTime    'The Start Of Day time must be before the End Of Lunch time
                      End If
                    End If
                  ElseIf vPanelItem.AttributeName = "end_of_day" Then
                    If epl.GetValue("StartOfDay").Trim.Length > 1 Then
                      If TimeValue(epl.GetValue("EndOfDay")) <= TimeValue(epl.GetValue("StartOfDay")) Then
                        pValid = False
                        vValue = InformationMessages.ImEndOfDayTimeMustBeAfterTheStartOfDayTime   'The End Of Day time must be after the Start Of Day time
                      End If
                    ElseIf vLen1 > 0 Then 'start_of_lunch
                      If TimeValue(epl.GetValue("EndOfDay")) <= TimeValue(epl.GetValue("StartOfLunch")) Then
                        pValid = False
                        vValue = InformationMessages.ImEndOfDayTimeMustBeAfterTheStartOfLunchTime    'The End Of Day time must be after the Start Of Lunch time
                      End If
                    ElseIf vLen2 > 0 Then 'end_of_lunch
                      If TimeValue(epl.GetValue("EndOfDay")) <= TimeValue(epl.GetValue("EndOfLunch")) Then
                        pValid = False
                        vValue = InformationMessages.ImEndOfDayTimeMustBeBeforeTheEndOfLunchTime    'The End Of Day time must be before the End Of Lunch time
                      End If
                    End If
                  End If
                  If vPanelItem.AttributeName = "start_of_lunch" And vLen1 > 0 Then
                    If vLen2 = 0 Then 'end_of_lunch
                      pValid = False
                      vValue = InformationMessages.ImBothStartOfLunchAndEndOfLunch    'Either both the Start Of Lunch time and End Of Lunch time must be specified or both must be null
                    ElseIf TimeValue(epl.GetValue("StartOfLunch")) <= TimeValue(epl.GetValue("StartOfDay")) Then
                      pValid = False
                      vValue = InformationMessages.ImStartOfLunchTimeMustBeAfterTheStartOfDayTime    'The Start Of Lunch time must be after the Start Of Day time
                    ElseIf TimeValue(epl.GetValue("StartOfLunch")) >= TimeValue(epl.GetValue("EndOfDay")) Then
                      pValid = False
                      vValue = InformationMessages.ImStartOfLunchTimeCannotBeAfterTheEndOfDayTime    'The Start Of Lunch time can not be after the End Of Day time
                    End If
                  End If
                  If vPanelItem.AttributeName = "end_of_lunch" And vLen2 > 0 Then
                    If vLen1 = 0 Then 'start_of_lunch
                      pValid = False
                      vValue = InformationMessages.ImBothStartOfLunchAndEndOfLunch    'Either both the Start Of Lunch time and End Of Lunch time must be specified or both must be null
                    ElseIf TimeValue(epl.GetValue("EndOfLunch")) <= TimeValue(epl.GetValue("StartOfLunch")) Then
                      pValid = False
                      vValue = InformationMessages.ImEndOfLunchTimeMustBeAfterTheStartOfLunchTime    'The End Of Lunch time must be after the Start Of Lunch time
                    ElseIf TimeValue(epl.GetValue("EndOfLunch")) <= TimeValue(epl.GetValue("StartOfDay")) Then
                      pValid = False
                      vValue = InformationMessages.ImEndOfLunchTimeMustBeAfterTheStartOfDayTime    'The End Of Lunch time must be after the Start Of Day time
                    ElseIf TimeValue(epl.GetValue("EndOfLunch")) >= TimeValue(epl.GetValue("EndOfDay")) Then
                      pValid = False
                      vValue = InformationMessages.ImEndOfLunchTimeMustBeBeforeTheEndOfDayTime    'The End Of Lunch time must be before the End Of Day time
                    End If
                  End If
                  If vValue.Length > 0 Then epl.SetErrorField(vPanelItem.ParameterName, vValue)

                Case "contact_groups", "organisation_groups", "event_groups"
                  Select Case vPanelItem.AttributeName
                    Case "tab_prefix"
                      If (mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend AndAlso vParamValue <> GetOriginalValue(vPanelItem.ParameterName)) OrElse mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew Then
                        vList = New ParameterList(True)
                        vList("TabPrefix") = vParamValue
                        If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctTabPrefixesContactGroups, vList) > 0 OrElse DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctTabPrefixesOrganisationGroups, vList) > 0 Then
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImTabPrefixInUse)    'This Tab Prefix is already in use
                          pValid = False
                        End If
                        If pValid Then
                          If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctTabPrefixesEventGroups, vList) > 0 Then
                            epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImTabPrefixInUse)    'This Tab Prefix is already in use
                            pValid = False
                          End If
                        End If
                      End If
                  End Select

                Case "cpd_cycle_types"
                  If vPanelItem.ParameterName.Equals("StartMonth", StringComparison.InvariantCultureIgnoreCase) _
                  OrElse vPanelItem.ParameterName.Equals("EndMonth", StringComparison.InvariantCultureIgnoreCase) Then
                    Dim vStartMonth As Nullable(Of Integer)
                    Dim vEndMonth As Nullable(Of Integer)
                    If vPanelItem.ParameterName.Equals("StartMonth", StringComparison.InvariantCultureIgnoreCase) Then
                      If vParamValue.Length > 0 Then vStartMonth = IntegerValue(vParamValue)
                      If epl.GetValue("EndMonth").Length > 0 Then vEndMonth = IntegerValue(epl.GetValue("EndMonth"))
                    Else
                      If vParamValue.Length > 0 Then vEndMonth = IntegerValue(vParamValue)
                      If epl.GetValue("StartMonth").Length > 0 Then vStartMonth = IntegerValue(epl.GetValue("StartMonth"))
                    End If
                    If Not ((vStartMonth.HasValue = True AndAlso vEndMonth.HasValue = True) _
                    OrElse (vStartMonth.HasValue = False AndAlso vEndMonth.HasValue = False)) Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCPDCycleTypesStartEndMonthsInvalid)   'The Start Month and End Month must either both be set or both be null
                      pValid = False
                    End If
                  End If

                Case "currency_rates"
                  If vPanelItem.AttributeName = "date_to" Then
                    Dim vCode As String = epl.GetValue("CurrencyCode")
                    vValue = epl.GetValue("DateFrom")
                    If vCode <> GetOriginalValue("CurrencyCode") AndAlso vValue <> GetOriginalValue("DateFrom") _
                    AndAlso vParamValue <> GetOriginalValue("DateTo") Then
                      vList = New ParameterList(True)
                      vList("CurrencyCode") = vCode
                      vList("DateFrom") = vValue
                      vList("DateTo") = vParamValue

                      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCurrencyRates, vList) > 0 Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImOverlappingCurrencyRate)    'This overlaps with an existing Currency Rate
                        pValid = False
                      End If
                    End If
                  End If

                Case "event_extra_fee_multipliers"
                  If vPanelItem.AttributeName = "from_time" OrElse vPanelItem.AttributeName = "to_time" Then
                    If TimeValue(epl.GetValue("ToTime")) <= TimeValue(epl.GetValue("FromTime")) Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImToTimeMustBeAfterFromTime)    'The To Time must be after the From Time
                    End If
                    If pValid AndAlso vPanelItem.AttributeName = "to_time" Then
                      'Check no over-laps with other records
                      vList = New ParameterList(True)
                      vList("ToTime") = vParamValue
                      vList("FromTime") = epl.GetValue("FromTime")
                      vList("EventPricingMatrix") = epl.GetValue("EventPricingMatrix")
                      vList("FeeMultiplierType") = epl.GetValue("FeeMultiplierType")
                      If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then vList("EventExtraFeeNumber") = epl.GetValue("EventExtraFeeNumber")

                      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctEventExtraFeeMultipliers, vList) > 0 Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImOverlappingEventExtraFieldMultiplier)    'This overlaps with an existing Event Extra Field Multiplier
                      End If
                    End If
                  End If

                Case "event_pricing_matrices"
                  If vPanelItem.AttributeName = "child_exempt_vat_rate" Then
                    vLen1 = epl.GetValue("ChildVatExemptTopic").Length
                    vLen2 = epl.GetValue("ChildExemptVatRate").Length
                    If vLen1 = 0 And vLen2 > 0 Then
                      pValid = False
                      vValue = InformationMessages.ImChildExemptVATRateAndChildVATExemptTopic    'The Child Exempt VAT Rate can only be set if the Child VAT Exempt Topic is set
                    ElseIf vLen1 > 0 And vLen2 = 0 Then
                      pValid = False
                      vValue = InformationMessages.ImChildExemptVATRateMustBeSetIfChildVATExemptTopicSet    'The Child Exempt VAT Rate must be set if the Child VAT Exempt Topic is set
                    End If
                    If pValid = False Then epl.SetErrorField(vPanelItem.ParameterName, vValue)
                  ElseIf vPanelItem.AttributeName = "extra_session_fee_product" Then
                    vLen1 = epl.GetValue("ExtraSessionFeeTopic").Length
                    vLen2 = epl.GetValue("ExtraSessionFeeProduct").Length
                    If vLen1 = 0 And vLen2 > 0 Then
                      pValid = False
                      vValue = InformationMessages.ImExtraSessionFeeProductAndExtraSessionFeeTopic    'The Extra Session Fee Product can only be set if the Extra Session Fee Topic is set
                    ElseIf vLen1 > 0 And vLen2 = 0 Then
                      pValid = False
                      vValue = InformationMessages.ImExtraSessionFeeProductMustBeSetIfExtraSessionFeeTopicSet    'The Extra Session Fee Product must be set if the Extra Session Fee Topic is set
                    End If
                    If pValid = False Then epl.SetErrorField(vPanelItem.ParameterName, vValue)
                  ElseIf vPanelItem.AttributeName = "event_fee_end_date" OrElse vPanelItem.AttributeName = "event_fee_start_date" Then
                    vDate1 = epl.GetValue("EventFeeStartDate")
                    vDate2 = epl.GetValue("EventFeeEndDate")
                    If vDate1.Length > 0 AndAlso vDate2.Length > 0 Then
                      If DateDiff("d", vDate1, vDate2) < 0 Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImEventPricingMatrixStartEndInvalid)
                      End If
                    End If
                  End If

                Case "financial_controls"
                  Select Case vPanelItem.AttributeName
                    Case "adjustment_transaction_type"
                      If vParamValue.Length > 0 Then
                        vList = New ParameterList(True)
                        vList("TransactionType") = vParamValue
                        vList("TransactionSign") = "C"
                        vList("NegativesAllowed") = "Y"

                        If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctTransactionTypes, vList) = 0 Then
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMustBeACreditTypeTransaction) 'Must be a credit-type Transaction Type that allows negative values
                          pValid = False
                        End If
                      End If
                    Case "first_claim_transaction_type"
                      If vParamValue.Length > 0 Then
                        If epl.GetValue("OneOffClaimTransactionType") = vParamValue Then
                          pValid = epl.SetErrorField("FirstClaimTransactionType", InformationMessages.ImFirstClaimAndOneOffTransTypeMustBeDifferent)
                        End If
                      End If
                    Case "fundraising_payment_type"
                      If vParamValue.Length > 0 AndAlso epl.GetValue("FundraisingStatus").Length = 0 Then
                        epl.SetErrorField("FundraisingStatus", InformationMessages.ImFieldMustNotBeBlank)    'Field must Not be Blank
                        pValid = False
                      End If
                    Case "one_off_claim_transaction_type"
                      If vParamValue.Length > 0 Then
                        If epl.GetValue("FirstClaimTransactionType") = vParamValue Then
                          pValid = epl.SetErrorField("OneOffClaimTransactionType", InformationMessages.ImFirstClaimAndOneOffTransTypeMustBeDifferent)
                        End If
                      End If
                  End Select

                Case "incentive_scheme_products"
                  If vPanelItem.AttributeName = "incentive_type" Then
                    If AppValues.ConfigurationValue(AppValues.ConfigurationValues.fixed_renewal_M).Length > 0 Then
                      If epl.GetValue("IncentiveType") = "I" Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotAssignInitialPeriodIncentiveType)    'Cannot assign Initial Period Incentive Type if A Fixed Renewal Cycle is Set
                        pValid = False
                      End If
                    End If
                  End If

                Case "internal_resources"
                  If Not mvInternalResourceMesgShown Then
                    If epl.GetValue("Product").Length > 0 Then vIRProductSet = True
                    If epl.GetValue("Rate").Length > 0 Then vIRRateSet = True
                    If epl.GetValue("ResourceContactNumber").Length > 0 Then vIRContactSet = True
                  End If

                Case "membership_controls"
                  Select Case vPanelItem.AttributeName
                    Case "advanced_cmt"
                      If BooleanValue(vParamValue) = True AndAlso epl.GetValue("CmtCalcProportionalBalance") = "N" Then
                        'Error - CMT does not handle pro-rating
                        pValid = epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotSetAdvancedCMT)
                      End If
                    Case "organisation_group"
                      If vParamValue.Length > 0 Then
                        If vParamValue = "ORG" Then
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImOrganisationGroupCannotBeSetToDefaultGroup)    'The Organisation Group cannot be set to the default Group of 'ORG'
                          pValid = False
                        ElseIf GetOriginalValue("OrganisationGroup").Length > 0 Then
                          If vParamValue <> GetOriginalValue("OrganisationGroup") Then
                            'The organisationGroup has changed
                            vList = New ParameterList(True)
                            vList("OrganisationGroup") = GetOriginalValue("OrganisationGroup")
                            If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMembershipGroups, vList) > 0 Then
                              epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImMembershipGroupsSetUpUsingOrganisationGroup, GetOriginalValue("OrganisationGroup")))    'There is already Membership Groups data set up using Organisation Group '%s'; the Organisation Group cannot be changed
                              pValid = False
                            End If
                          End If
                        End If
                      End If
                  End Select

                Case "membership_prices"
                  If vPanelItem.AttributeName = "rate" AndAlso vParamValue.Length > 0 AndAlso epl.GetValue("MembershipType").Length > 0 Then
                    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmNew AndAlso epl.GetValue("PaymentMethod").Length > 0 AndAlso epl.GetValue("PaymentFrequency").Length > 0 Then
                      vList = New ParameterList(True)
                      vList("MembershipType") = epl.GetValue("MembershipType")
                      vList("Rate") = vParamValue
                      vList("PaymentMethod") = epl.GetValue("PaymentMethod")
                      vList("PaymentFrequency") = epl.GetValue("PaymentFrequency")
                      If epl.FindPanelControl("Overseas", False) IsNot Nothing Then
                        vList("Overseas") = epl.GetValue("Overseas")
                        vList("Activity") = epl.GetValue("Activity")
                        vList("ActivityValue") = epl.GetValue("ActivityValue")
                      End If
                      Dim vResult As ParameterList = DataHelper.CheckMembershipTypeRate(vList)
                      Dim vMessage As String = String.Empty
                      If vResult.ContainsKey("Message") Then vMessage = vResult("Message")
                      If vMessage.Length > 0 Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, vMessage)
                      End If
                    End If
                  End If

                Case "membership_type_categories"
                  If vPanelItem.AttributeName = "membership_type" OrElse vPanelItem.AttributeName = "activity" Then
                    If vParamValue.Length = 0 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFieldMandatory)    'Field is mandatory and cannot be left empty
                    End If
                  End If
                  vList = New ParameterList(True)
                  vList("MembershipType") = epl.GetValue("MembershipType")
                  vList("Activity") = epl.GetValue("Activity")
                  vList("ActivityValue") = epl.GetValue("ActivityValue")
                  If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMembershipTypeCategories, vList) > 1 Then
                    epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImEntryAlreadyExists)
                  End If

                Case "membership_types"
                  If vPanelItem.AttributeName = "adult_gift_member_eligible_ga" Then
                    If vParamValue = "Y" Then
                      If epl.GetValue("MembershipLevel") = "J" Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotSetAdultGiftMemberEligibleGiftAid)    'Can not set the Adult Gift Member Eligible Gift Aid flag for a Junior Membership
                      End If
                    End If
                  ElseIf vPanelItem.AttributeName = "membership_card_duration" Then
                    If vParamValue.Length = 0 AndAlso epl.GetValue("MembershipCard") = "Y" Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMembershipCardDurationMustBeSpecified)    'Membership Card Duration must be specified for a Membership Card
                      pValid = False
                    End If
                  ElseIf vPanelItem.AttributeName = "cancelled_associate_product" Then
                    If vParamValue.Length = 0 And epl.GetValue("AssociateMembershipType").Length > 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCancelledAssociateProduct)    'Cancelled Associate Product must be specified for Memberships with an Associated Membership
                      pValid = False
                    End If
                  ElseIf mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend AndAlso (vPanelItem.AttributeName = "payer_required" OrElse vPanelItem.AttributeName = "associate_membership_type" OrElse vPanelItem.AttributeName = "members_per_order") Then
                    vList = New ParameterList(True)
                    vList("MembershipType") = epl.GetValue("MembershipType")
                    Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
                    If Not vDataRow Is Nothing Then
                      If vParamValue <> vDataRow(vPanelItem.ParameterName).ToString Then
                        vChanged = True
                      End If
                      If vChanged Then
                        'Fields changed, prevent change if this is for a subsequent membership type
                        vList("SubsequentMembershipType") = vList("MembershipType")
                        vList.Remove("MembershipType")
                        vDataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
                        If Not vDataRow Is Nothing Then
                          epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImMembershipTypeAlsoSubsequent, StrConv(vPanelItem.AttributeName.Replace("_", " "), VbStrConv.ProperCase)))    'This Membership Type is also a Subsequent Type, %s can not be changed
                          pValid = False
                        End If
                      End If
                    End If
                  End If

                  If vPanelItem.AttributeName = "associate_membership_type" AndAlso vParamValue.Length > 0 Then
                    vList = New ParameterList(True)
                    vList("MembershipType") = vParamValue
                    Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
                    If vDataRow Is Nothing OrElse epl.GetValue("FixedCycle") <> vDataRow("FixedCycle").ToString Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFixedCycleMustMatch)    'The Fixed Cycle of the Associate Membership Type must match the Fixed Cycle specified for this Membership Type
                      pValid = False
                    End If
                  End If

                  If epl.GetValue("SubsequentMembershipType").Length > 0 Then
                    'Future member Type validation
                    If vPanelItem.AttributeName = "subsequent_membership_type" Then
                      If epl.GetValue("MembershipType") = epl.GetValue("SubsequentMembershipType") Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSubsequentMTMustBeDifferent)    'The Subsequent Membership Type must be different to the Original Membership Type
                        pValid = False
                      End If
                      If pValid Then
                        vList = New ParameterList(True)
                        vList("MembershipType") = vParamValue
                        Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
                        If Not vDataRow Is Nothing Then
                          If vDataRow("AssociateMembershipType").ToString.Length > 0 Then
                            epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSubsequentMTMustNotHaveAT)    'The Subsequent Membership Type must not have an Associate Type
                            pValid = False
                          End If
                          If pValid AndAlso DoubleValue(vDataRow("MembersPerOrder").ToString) <> 1 Then
                            epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSubsequentMTAndMembersPerOrder)    'The Subsequent Membership Type must have Members Per Order set to '1'
                            pValid = False
                          End If
                          If pValid AndAlso Not vDataRow("FixedCycle").ToString = epl.GetValue("FixedCycle") Then
                            epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSubsequentMTFixedCycle)    'The Fixed Cycle of the Subsequent Membership Type must match the Fixed Cycle specified for this Membership Type
                            pValid = False
                          End If
                          If pValid AndAlso Not vDataRow("Annual").ToString = epl.GetValue("Annual") Then
                            epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImSubsequentMTAnnual, vDataRow("Annual").ToString, epl.GetValue("Annual")))    'The Annual of the Subsequent Membership Type (%s) must match the Annual specified for this Membership Type (%s)
                            pValid = False
                          End If
                          If pValid AndAlso Not vDataRow("MembershipTerm").ToString = epl.GetValue("MembershipTerm") Then
                            epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImSubsequentMTMembershipTerm, vDataRow("MembershipTerm").ToString, epl.GetValue("MembershipTerm")))    'The Membership Term of the Subsequent Membership Type (%s) must match the Membership Term specified for this Membership Type (%s)
                            pValid = False
                          End If
                          If pValid AndAlso Not vDataRow("PayerRequired").ToString = epl.GetValue("PayerRequired") Then
                            epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImSubsequentMTPayerRequired, vDataRow("PayerRequired").ToString, epl.GetValue("PayerRequired")))    'The Payer Required of the Subsequent Membership Type (%s) must match the Payer Required specified for this Membership Type (%s)
                            pValid = False
                          End If
                        End If
                      End If
                    End If

                    If vPanelItem.AttributeName = "members_per_order" AndAlso IntegerValue(epl.GetValue("MembersPerOrder")) <> 1 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMembersPerOrderAndSubsequentType)    'The Members Per Order must be set to '1' if a Subsequent Type is specified
                      pValid = False
                    End If

                    If vPanelItem.AttributeName = "associate_membership_type" AndAlso vParamValue.Length > 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImAssocMembershipTypeAndSubsequentType)    'The Associate Membership Type must be null if a Subsequent Type is specified
                      pValid = False
                    End If

                  End If

                  If vPanelItem.AttributeName = "subsequent_membership_type" OrElse vPanelItem.AttributeName = "subsequent_trigger" Then
                    If epl.GetValue("SubsequentMembershipType").Length > 0 AndAlso epl.GetValue("SubsequentTrigger").Length = 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMembershipTypeTriggerRequired)    'A Trigger Method must be specifed for Membership Types with a Subsequent Type
                      pValid = False
                    End If
                    If pValid AndAlso epl.GetValue("SubsequentTrigger").Length > 0 AndAlso epl.GetValue("SubsequentMembershipType").Length = 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSubsequentTypeRequired)    'A Subsequent Type must be specifed for Membership Types with a Trigger Method
                      pValid = False
                    End If
                    If pValid And epl.GetValue("SubsequentTrigger") = "A" AndAlso epl.GetValue("MaxJuniorAge").Length = 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImTriggerMethodInvalid)    'The Trigger Method can only be set to 'A' for Membership Types with a Maximum Junior Age
                      pValid = False
                    End If
                  End If

                  If vPanelItem.AttributeName = "fixed_cycle" Then
                    If GetOriginalValue(vPanelItem.ParameterName) <> vParamValue Then  'the value of the Fixed Cycle has been changed
                      'prevent this change if there are records in either the members table or the orders table that have this membership type
                      vList = New ParameterList(True)
                      vList("MembershipType") = epl.GetValue("MembershipType")

                      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctFixedCycles, vList) > 0 Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotChangeFixedCycle) 'The Fixed Cycle cannot be changed because members have been created using this Membership Type
                      End If
                    End If
                  End If
                  If vPanelItem.AttributeName = "membership_term" AndAlso IntegerValue(vParamValue) = 0 Then
                    Select Case epl.GetValue("Annual")
                      Case "M", "W"
                        'For Monthly and Weekly Memberships the Membership Term cannot be null or less than 1
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMembershipTypeMembershipTermCannotBeNullOrZero) 'The Membership Term cannot be null or zero with this Annual setting
                      Case Else
                        If Not vParamValue = String.Empty Then
                          'Invalid Membership Term of 0 specified
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImMembershipTypeMembershipTermCannotBeZero) 'The Membership Term, when set, must be greater than zero
                        End If
                    End Select
                  End If

                Case "prize_draws"
                  Select Case vPanelItem.AttributeName.ToLower
                    Case "appeal"
                      If vParamValue.Length > 0 Then
                        'Campaign & Appeal must not have already been used & must be mailing type SR
                        vList = New ParameterList(True)
                        vList("Campaign") = epl.GetValue("Campaign")
                        vList("Appeal") = epl.GetValue("Appeal")
                        If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew Then vList("PrizeDraw") = epl.GetValue("PrizeDraw")

                        If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPrizeDraws, vList) > 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCampaignAppealUsedInPrizeDraw)    'This Campaign and Appeal has already been used in a Prize Draw
                        Else
                          If mvEditMode <> CareNetServices.XMLTableMaintenanceMode.xtmmNew Then vList.Remove("PrizeDraw")
                          vList("MailingType") = "SR"
                          If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPrizeDrawAppeals, vList) = 0 Then
                            pValid = False
                            epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImInvalidPrizeDrawAppealType)    'The Appeal must be for a 'Sale or Return' mailing type for use in a Prize Draw
                          End If
                        End If
                      End If

                    Case "bank_account"
                      If epl.GetValue("Campaign").Length > 0 Then
                        'Campaign & Appeal chosen so Bank Account must be chosen
                        If vParamValue.Length = 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImBankAccountCampaignRequired)    'Bank Account must be specified if a Campaign has been selected
                        End If
                      End If

                    Case "close_date"
                      If vParamValue.Length > 0 Then
                        Dim vCloseDate As Date = DateValue(vParamValue)
                        Dim vDrawDate As Date = DateValue(epl.GetValue("DrawDate"))
                        If vDrawDate.CompareTo(vCloseDate) > 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImPrizeDrawCloseDateBeforeDrawDate)
                        End If
                      Else
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFieldMandatory)
                      End If

                    Case "product"
                      If epl.GetValue("Campaign").Length > 0 Then
                        'Campaign & Appeal chosen so Product must be chosen
                        'Product must be using product_numbers and have the sales_quantity set > 0
                        If vParamValue.Length > 0 Then
                          vList = New ParameterList(True)
                          vList("Product") = vParamValue
                          Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
                          If Not vDataRow Is Nothing Then
                            If BooleanValue(vDataRow("UsesProductNumbers").ToString) Then
                              If vDataRow("NextProductNumber").ToString.Length > 0 Then
                                If DoubleValue(vDataRow("SalesQuantity").ToString) < 1 Then pValid = False
                              Else
                                pValid = False
                              End If
                            Else
                              pValid = False
                            End If
                          Else
                            pValid = False
                          End If
                          If pValid = False Then epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImPrizeDrawProductNotSetupCorrectly) 'The Product has not been correctly set up for a Prize Draw Mailing
                        Else
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImProductCampaignRequired)    'Product must be specified if a Campaign has been selected
                        End If
                      End If

                    Case "transaction_type"
                      If epl.GetValue("Campaign").Length > 0 Then
                        'Campaign & Appeal chosen so Transaction Type must be chosen
                        If vParamValue.Length = 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImTransactionTypeCampaignRequired)    'Transaction Type must be chosen if a Campaign has been selected
                        End If
                      End If

                  End Select

                Case "product_warehouses"
                  If vPanelItem.AttributeName = "product" Then
                    vList = New ParameterList(True)
                    vList("Product") = vParamValue
                    Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
                    If Not vDataRow Is Nothing Then
                      pValid = BooleanValue(vDataRow("StockItem").ToString)
                    Else
                      pValid = False
                    End If
                    If pValid = False Then epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImProductMustBeStockItem) 'Product must be a stock item
                  End If

                Case "products"
                  If vPanelItem.AttributeName = "donation" Then
                    If CShort(epl.GetValue("Donation") = "Y") + CShort(epl.GetValue("Subscription") = "Y") + CShort(epl.GetValue("StockItem") = "Y") +
                      CShort(epl.GetValue("Course") = "Y") + CShort(epl.GetValue("Accommodation") = "Y") + CShort(epl.GetValue("PostagePacking") = "Y") +
                      CShort(epl.GetValue("UsesProductNumbers") = "Y") + CShort(epl.GetValue("SponsorshipEvent") = "Y") +
                      CShort(epl.GetOptionalValue("AccruesInterest") = "Y") + CShort(epl.GetOptionalValue("Exam") = "Y") < -1 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImTooManyProductFlagsSet)    'Too many product flags have been set. See notes.
                    End If
                  ElseIf vPanelItem.AttributeName = "next_product_number" Then
                    If vParamValue.Length = 0 Then
                      If epl.GetValue("UsesProductNumbers") = "Y" Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImNextProductNumberRequired)    'Next Product Number must be specified when Uses Product Number checked
                        pValid = False
                      End If
                    End If
                  ElseIf vPanelItem.AttributeName = "despatch_method" Then
                    If vParamValue.Length = 0 Then
                      If epl.GetValue("Subscription") = "Y" Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImDespatchMethodRequiredSubscriptionProduct)    'Despatch Method must be specified for a Subscription Product
                        pValid = False
                      ElseIf epl.GetValue("StockItem") = "Y" Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImDespatchMethodRequiredStockProduct)    'Despatch Method must be specified for a Stock Product
                        pValid = False
                      End If
                    End If
                  ElseIf vPanelItem.AttributeName.Contains("cost_of_sale") Then
                    Dim vSelected As Short = CShort(epl.GetValue("CostOfSale").Length > 0) + CShort(epl.GetValue("CostOfSaleAccount").Length > 0) +
                       CShort(epl.GetValue("CostOfSaleAccrual").Length > 0)
                    If Not (vSelected = -3 OrElse vSelected = 0) Then 'all or none entered
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCostSaleCostofSaleAccountAndAccural)
                    End If
                  Else
                    If epl.GetValue("StockItem") = "Y" Then
                      If vPanelItem.AttributeName = "warehouse" Then
                        If vParamValue.Length = 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImWarehouseRequiredForStockProducts)    'Warehouse must be specified for Stock Products
                        Else
                          If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
                            If mvStockFlagAfter = 1 Then
                              pValid = True
                            Else
                              vList = New ParameterList(True)
                              vList("Product") = epl.GetValue("Product")
                              vList("Warehouse") = epl.GetValue("Warehouse")

                              If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctProductWarehouses, vList) = 0 Then
                                pValid = False
                                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImWarehouseDoesntExistForThisProduct)    'Warehouse must exist in Product Warehouses for this Product
                              End If
                            End If
                          End If
                        End If
                      ElseIf vPanelItem.AttributeName = "bin_number" Then
                        If vParamValue.Length = 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImBinNumberRequiredForStockProducts)    'Bin Number must be specified for Stock Products
                        End If
                      ElseIf vPanelItem.AttributeName = "last_stock_count" Then
                        If vParamValue.Length = 0 Then
                          vParamValue = "0"
                        End If
                      End If
                    End If

                    If AppValues.ConfigurationOption(AppValues.ConfigurationOptions.nominal_account_validation) AndAlso mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
                      If vPanelItem.AttributeName = "nominal_account" OrElse vPanelItem.AttributeName = "subsequent_nominal_account" Then
                        If vParamValue <> GetOriginalValue(vPanelItem.ParameterName) Then
                          Try
                            'value has changed
                            vList = New ParameterList(True)
                            vList("CheckNominalAccount") = IIf(vPanelItem.AttributeName = "nominal_account", "Y", "N").ToString
                            vList("Product") = epl.GetValue("Product")
                            vList("NominalAccount") = vParamValue
                            'This will raise an exception if changing the value would invalidate the Nominal Account Suffixes in the rates tabel
                            DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtRateNominalAccounts, vList)
                          Catch vEx As CareException
                            If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidNominalAccountSuffixes Then
                              'Some suffixes are invalid for the new nominal account
                              pValid = False
                              epl.SetErrorField(vPanelItem.ParameterName, vEx.Message)    'The Nominal Account code can not be changed as it will invalidate the Nominal Account Suffixes on the Rates table
                            Else
                              DataHelper.HandleException(vEx)
                            End If
                          End Try
                        End If
                      End If
                    End If
                  End If

                Case "rates"
                  If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
                    If vPanelItem.AttributeName = "current_price" OrElse vPanelItem.AttributeName = "future_price" Then
                      If IntegerValue(GetOriginalValue(vPanelItem.ParameterName)) = 0 AndAlso IntegerValue(vParamValue) <> 0 Then
                        vList = New ParameterList(True)
                        vList("Product") = epl.GetValue("Product")
                        vList("Rate") = epl.GetValue("Rate")
                        If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPackedProductsLinkProduct, vList) > 0 Then
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCurrentFuturePriceCannotBeChanged)    'The current && future price can not be changed as it is part of a Pack
                        End If
                      End If
                    End If
                  End If
                  If vPanelItem.AttributeName = "days_prior_to" Then
                    If epl.GetValue("DaysPriorFrom") <> String.Empty Then
                      If IntegerValue(epl.GetValue("DaysPriorTo")) > IntegerValue(epl.GetValue("DaysPriorFrom")) Then
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImDayPriorFromGreaterDayPriorTo) 'Days prior from should be greater than days prior to
                      End If
                    End If
                  End If
                  If FindControl(epl, "LoanInterest", False) IsNot Nothing Then
                    Select Case vPanelItem.ParameterName
                      Case "CurrentPrice", "FuturePrice", "LoanInterest", "VatExclusive"
                        vList = New ParameterList(True, True)
                        vList("Product") = epl.GetValue("Product")
                        If vPanelItem.ParameterName = "LoanInterest" Then
                          vList("Donation") = "N"
                          vList("Subscription") = "N"
                          vList("StockItem") = "N"
                          vList("Course") = "N"
                          vList("Accommodation") = "N"
                          vList("PostagePacking") = "N"
                          vList("UsesProductNumbers") = "N"
                          vList("SponsorshipEvent") = "N"
                          vList("AccruesInterest") = "N"
                          vList("Exam") = "N"
                        Else
                          vList("AccruesInterest") = "Y"
                        End If
                        Dim vProductRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
                        Select Case vPanelItem.ParameterName
                          Case "CurrentPrice", "FuturePrice"
                            If BooleanValue(epl.GetOptionalValue("LoanInterest")) = True Then
                              If DoubleValue(vParamValue) <> 0 Then
                                pValid = False
                                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImLoanInterestPriceNotZero)    'Current and Future Price must be zero when Loan Interest is set
                              End If
                            ElseIf vProductRow IsNot Nothing Then
                              If DoubleValue(vParamValue) <> 0 Then
                                pValid = False
                                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImLoanAccruesInterestPriceNotZero)    'Current and Future Price must be zero when the Product has the Accrues Interest flag set
                              End If
                            End If
                          Case "LoanInterest"
                            If BooleanValue(vParamValue) = True AndAlso vProductRow Is Nothing Then
                              pValid = False
                              epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImLoanInterestInvalidProduct)    'Loan Interest can only be set for a Product with no flags set
                            End If
                          Case "VatExclusive"
                            If BooleanValue(vParamValue) = True Then
                              If BooleanValue(epl.GetOptionalValue("LoanInterest")) = True Then
                                pValid = False
                                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImLoanInterestNotVatExclusive)   'Rate cannot be VAT Exclusive when Loan Interest is set
                              ElseIf vProductRow IsNot Nothing Then
                                pValid = False
                                epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImLoanAccruesInterestNotVatExclusive)    'Rate cannot be VAT Exclusive when the Product has the Accrues Interest flag set
                              End If
                            End If
                        End Select
                    End Select
                  End If
                Case "room_types"
                  If vPanelItem.AttributeName = "enforce_allocation" Then
                    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend AndAlso vParamValue <> GetOriginalValue(vPanelItem.ParameterName) Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImEnforceAllocationFlagCannotBeAmended)    'The Enforce Allocation Flag cannot be amended
                      pValid = False
                    ElseIf vParamValue = "Y" AndAlso IntegerValue(epl.GetValue("Capacity")) < 2 Then
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImEnforceAllocationFlagCannotBeSet)    'The Enforce Allocation Flag cannot be set if Capacity is less than 2
                      pValid = False
                    End If
                  End If
                  If vPanelItem.AttributeName = "capacity" Then
                    If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
                      If epl.GetValue("EnforceAllocation") = "Y" AndAlso vParamValue <> GetOriginalValue(vPanelItem.ParameterName) Then
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotAmendedCapacity)    'Capacity cannot be amended when Enforce Allocation is set
                        pValid = False
                      End If
                    End If
                  End If

                Case "statuses"
                  If vPanelItem.AttributeName = "contact_group" AndAlso vParamValue.Length > 0 Then
                    vList = New ParameterList(True)
                    vList("Status") = epl.GetValue("Status")
                    If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCancellationReasons, vList) > 0 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImSelectedStatusInUsedByCancellationReason)    'The selected status is being used by a cancellation reason. You cannot apply a Contact Group to it.
                    End If
                  End If

                Case "sub_topics"
                  If vPanelItem.AttributeName = "activity_duration" Then
                    If vParamValue.Length > 0 AndAlso epl.GetValue("Activity").Length = 0 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImActivityDurationRequiresActivity)    'The Activity Duration can only be set if an Activity has been set
                    End If
                  End If

                Case "users"
                  If vPanelItem.AttributeName = "history_only" Then
                    If BooleanValue(vParamValue) Then
                      vList = New ParameterList(True)
                      vList("Logname") = epl.GetValue("Logname") 'NFPCARE-88: was previously set as PrincipalUser
                      Dim vCount As Integer = DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctPrincipalUsers, vList)
                      If vCount > 0 Then
                        epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImUserIsPrincipalUser, vCount))    'This user is a principal user for %s contact(s) and cannot be set as historic
                        pValid = False
                      End If
                    End If
                  End If
                  If vPanelItem.AttributeName = "logname" Then  'BR14116 - Restricting from creating a user with apostrophe
                    If epl.GetValue("Logname").IndexOf("'") >= 0 Then
                      epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImLognameContainsApostrophe))    'User cannot be created if it contains apostrophe
                      pValid = False
                    End If
                  End If

                Case "activity_groups"
                  If vPanelItem.AttributeName = "source" Then
                    If vParamValue.Length > 0 Then
                      If epl.GetValue("UsageCode") = "E" Then  '+ data structure info
                        'Contact Entry
                        If epl.GetValue("Campaign").Length > 0 OrElse epl.GetValue("Appeal").Length > 0 Then
                          'Source is set so Campaign & Appeal can not be set
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotSetCampaignAppeal)    'The Campaign and Appeal can not be set if the Source is set
                        End If
                      End If
                    End If
                  ElseIf vPanelItem.AttributeName = "campaign" OrElse vPanelItem.AttributeName = "appeal" Then
                    If vParamValue.Length > 0 Then
                      If epl.GetValue("UsageCode") = "E" Then
                        'Contact Entry
                        If epl.GetValue("Source").Length > 0 Then
                          'Campaign & Appeal are set so Source can not be set
                          pValid = False
                          epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCannotSetSource)    'The Source can not be set if the Campaign and Appeal are set
                        Else
                          'Ensure that both Campaign and Appeal are set
                          If epl.GetValue("Campaign").Length > 0 Then
                            If epl.GetValue("Appeal").Length = 0 Then pValid = False
                          Else
                            pValid = False
                          End If
                          If pValid = False Then epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImBothCampaignAndAppealMustBeSet) 'Both Campaign and Appeal must be set
                        End If
                        If pValid Then
                          'OK, now ensure Campaign & Appeal are unique
                          If vPanelItem.AttributeName = "appeal" AndAlso epl.GetValue("ActivityGroup").Length > 0 AndAlso epl.GetValue("Campaign").Length > 0 AndAlso epl.GetValue("Appeal").Length > 0 Then
                            vList = New ParameterList(True)
                            vList("Campaign") = epl.GetValue("Campaign")
                            vList("Appeal") = epl.GetValue("Appeal")
                            If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then vList("ActivityGroup") = epl.GetValue("ActivityGroup") 'NFPCARE-16:The parameter for the activity group was passed in as 'activity_group' instead of 'ActivityGroup'. 
                            If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctActivityGroups, vList) > 0 Then
                              pValid = False
                              epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImInvalidCampaignAppealInActivityGroup)    'This Campaign and Appeal has already been used in an Activity Group
                            End If
                          End If
                        End If
                      Else
                        'Other entry types - Campaign & Appeal should be null
                        pValid = False
                        epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCampaignAppealOnlyValidForContactEntry)    'The Campaign and Appeal can only be set for Contact Entry
                      End If
                    End If
                  End If

                  ' BR11756
                Case "service_control_restrictions"
                  If vPanelItem.AttributeName = "valid_from" Then
                    vList = New ParameterList(True)
                    vList("ContactNumber") = epl.GetValue("ContactNumber")
                    vList("ShortStayDuration") = epl.GetValue("ShortStayDuration")
                    vList("LateBookingDays") = epl.GetValue("LateBookingDays")
                    vList("DateFrom") = epl.GetValue("ValidFrom")
                    vList("DateTo") = epl.GetValue("ValidTo")
                    If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctServiceControlRestrictions, vList) > 0 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImRestrictionAlreadyExists)    'A restriction already exists for this contact/validity/duration/late booking days combination
                    End If
                  End If

                Case "geographical_regions"
                  If vPanelItem.AttributeName = "organisation_number" Then
                    vList = New ParameterList(True)
                    vList("GeographicalRegionType") = epl.GetValue("GeographicalRegionType")
                    vList("OrganisationRequired") = "Y"

                    If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctGeographicalRegionTypes, vList) > 0 AndAlso vParamValue.Length = 0 Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFieldMustNotBeBlank)    'Field must Not be Blank
                    End If
                  End If

                Case "geographical_region_types"
                  If mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmAmend Then
                    If vPanelItem.AttributeName = "organisation_required" Then
                      If GetOriginalValue(vPanelItem.ParameterName) <> vParamValue AndAlso vParamValue = "Y" Then
                        vList = New ParameterList(True)
                        vList("GeographicalRegionType") = GetOriginalValue("GeographicalRegionType")
                        If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctGeographicalRegions, vList) > 0 Then
                          pValid = False
                          ShowWarningMessage(InformationMessages.ImRecordCannotBeChanged, "Geographical Regions", "Geographical Region Type")    '%s refer to this %s & vbCrLf & vbCrLf & Record cannot be changed
                        End If
                      End If
                    End If
                  End If

                Case "distribution_affiliates"
                  If vPanelItem.AttributeName = "product" Then
                    vList = New ParameterList(True)
                    vList("Product") = vParamValue
                    Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
                    If vDataRow Is Nothing OrElse vDataRow("Donation").ToString = "N" Then
                      pValid = False
                      epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImProductMustHaveDonationFlagSet)    'Product must have the donation flag set
                    End If
                  End If

                Case "ownership_access_levels" 'NFPCARE-87
                  If vPanelItem.AttributeName = "ownership_access_level" Then
                    If epl.GetValue(vPanelItem.ParameterName).Length = 0 Then epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImFieldMandatory)
                    pValid = False
                  End If

                Case "payment_plan_surcharges"
                  If epl.GetValue("LiableProduct") = epl.GetValue("SurchargeProduct") Then
                    epl.SetErrorField("SurchargeProduct", InformationMessages.ImSurchargeProductMustBeDifferent)
                    pValid = False
                  End If

                Case "cmt_excess_payments"
                  If vPanelItem.ParameterName = "CmtExcessPaymentType" Then
                    Dim vCountList As New ParameterList(True, True)
                    vCountList("CmtExcessPayment") = epl.GetValue("CmtExcessPayment")
                    vCountList(vPanelItem.ParameterName) = vParamValue
                    If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctCMTExcessPayments, vCountList) > 0 Then
                      pValid = epl.SetErrorField(vPanelItem.ParameterName, InformationMessages.ImCMTExcessPaymentTypeAlreadyused)
                    End If
                  End If

                Case "membership_entitlement"
                  If vPanelItem.ParameterName = "SequenceNumber" AndAlso vParamValue.Length > 0 Then
                    epl.SetErrorField(vPanelItem.ParameterName, "")
                    Dim vCountList As New ParameterList(True, True)
                    vCountList("MembershipType") = epl.GetValue("MembershipType")
                    vCountList(vPanelItem.ParameterName) = vParamValue
                    vCountList("Product") = epl.GetValue("Product")
                    If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctMembershipEntitlements, vCountList) > 0 Then
                      pValid = epl.SetErrorField(vPanelItem.ParameterName, String.Format(InformationMessages.ImMemberEntitlementSequenceAlreadyUsed, vParamValue))
                    End If
                  End If

              End Select              'Select on mvTableName
            End If                    'If valid
          End If                      'If valid

          If pValid Then
            epl.SetErrorField(vPanelItem.ParameterName, "")
          Else
            Exit For
          End If

        End If                        'If Editable
      Next

      If pValid Then
        'This needs to be seperate line as if vValid = False on item 0, this then fails
        If mvTableName = "internal_resources" AndAlso Not mvInternalResourceMesgShown Then
          mvInternalResourceMesgShown = True
          If vIRProductSet = False AndAlso vIRRateSet = False AndAlso vIRContactSet = False Then
            pValid = False
          ElseIf vIRProductSet = True AndAlso vIRRateSet = True AndAlso vIRContactSet = True Then
            pValid = False
          ElseIf (vIRProductSet = True AndAlso vIRRateSet = False) OrElse (vIRProductSet = False AndAlso vIRRateSet = True) Then
            pValid = False
          End If
          If Not pValid Then
            ShowWarningMessage(InformationMessages.ImInternalResources)    'Either Product & Rate OR Resource Contact must be specified
          End If
        End If
      End If
    End If
  End Sub

  Private Function GetOriginalValue(ByVal pParameterName As String) As String
    Dim vValue As String = String.Empty
    If mvParams.ContainsKey(pParameterName) Then vValue = mvParams(pParameterName)
    Return vValue
  End Function

  Private Sub epl_GetInitialCodeRestrictions(ByVal sender As System.Object, ByVal pParameterName As System.String, ByRef pList As CDBNETCL.ParameterList) Handles epl.GetInitialCodeRestrictions
    Dim vControl As Control = epl.FindPanelControl(pParameterName)
    Dim vPanelItem As PanelItem = DirectCast(vControl.Tag, PanelItem)

    If vPanelItem.ValidationTable <> String.Empty Then
      If pList Is Nothing Then pList = New ParameterList(True)
      'Extra stuff to handle descriptions
      If vPanelItem.RestrictionAttribute.Length <> 0 Then
        If mvTableName = "membership_types" AndAlso vPanelItem.AttributeName = "branch_rate" Then
          pList("Product") = AppValues.ControlValue(AppValues.ControlTables.membership_controls, AppValues.ControlValues.product)
        ElseIf mvTableName = "membership_prices" AndAlso vPanelItem.AttributeName = "rate" Then
          If mvCriteria IsNot Nothing Then
            Dim vList As New ParameterList(True)
            vList("MembershipType") = mvCriteria("MembershipType")
            Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtMembershipTypes, vList)
            pList(ProperName(vPanelItem.RestrictionAttribute)) = vDataRow("FirstPeriodsProduct").ToString
          End If
        End If
      Else
        If pList Is Nothing Then pList = New ParameterList(True)

        Select Case mvTableName
          Case "activity_cpd_points"
            If vPanelItem.ParameterName.Equals("CpdCycleType") Then pList("CpdType") = "P"

          Case "cancellation_reasons"
            If vPanelItem.AttributeName = "status" Then pList("ContactGroupNull") = "Y"

          Case "cpd_cycle_category_types", "cpd_categories"
            If vPanelItem.AttributeName = "cpd_category_type" Then pList("ForMaintenance") = "Y" ' NFPCARE-435

          Case "contact_controls"
            Select Case vPanelItem.AttributeName
              Case "position_activity_group", "position_relationship_group"
                pList("UsageCode") = "P"
            End Select

          Case "event_pricing_matrices"
            If vPanelItem.AttributeName = "child_exempt_vat_rate" Then
              pList("Percentage") = "0"
              pList("ForMaintenance") = "Y" 'NFPCARE-32
            End If

          Case "exam_controls"
            If vPanelItem.AttributeName = "exemption_org_activity" Then pList("OrganisationGroup") = EntityGroup.DefaultOrganisationGroupCode

          Case "financial_controls"
            Select Case vPanelItem.AttributeName
              Case "despatch_transaction_type"
                pList("TransactionSign") = "C"
              Case "first_claim_transaction_type", "one_off_claim_transaction_type"
                pList("TransactionSign") = "C"
                pList("NegativesAllowed") = "N"
              Case "preview_invoice_std_document"
                pList("MailmergeHeader") = "INV"
                pList("Active") = "Y"
              Case "receipt_print_std_document"
                pList("MailmergeHeader") = "RECPT"
                pList("Active") = "Y"
            End Select

          Case "lookup_groups"
            If vPanelItem.AttributeName = "table_name" Then
              pList("RestrictTables") = "'sources','mailings','standard_documents','topics','statuses','departments', 'ownership_groups','mailing_suppressions','relationships','position_seniorities','position_functions','roles','titles','users','pis_numbers','distribution_codes','devices','membership_types'"
            End If

          Case "mailing_template_documents"
            If vPanelItem.AttributeName = "standard_document" Then
              If Not mvCriteria Is Nothing Then
                pList("MailingTemplate") = mvCriteria("MailingTemplate")
              Else
                pList("CMDHeaderCodes") = "Y"
              End If
            End If

          Case "mailing_templates"
            If vPanelItem.AttributeName = "standard_document" Then pList("CMDHeaderCodes") = "Y"

          Case "marketing_controls", "vat_rate_history"
            If vPanelItem.AttributeName = "vat_rate" Then pList("ForMaintenance") = "Y" 'NFPCARE-32

          Case "membership_controls"
            If vPanelItem.AttributeName = "card_default_standard_document" Then pList("MailmergeHeader") = "MCMM"

          Case "products"
            If vPanelItem.AttributeName = "nominal_account" Then pList("Active") = "Y"

          Case "rate_nominal_accounts"
            If vPanelItem.AttributeName = "product_nominal_account" Then pList("Active") = "Y"

          Case "rates"
            Select Case vPanelItem.AttributeName
              Case "membership_lookup_group"
                pList("TableName") = "membership_types"
                pList("Active") = "Y"
              Case "nominal_account_suffix", "subsequent_nominal_suffix"
                If Not mvCriteria Is Nothing Then
                  If mvCriteria.ContainsKey("Product") AndAlso mvCriteria("Product").Length > 0 Then
                    Dim vList As New ParameterList(True)
                    vList("Product") = mvCriteria("Product")
                    Dim vDataRow As DataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtProducts, vList)
                    If vPanelItem.AttributeName = "nominal_account_suffix" Then
                      pList("ProductNominalAccount") = vDataRow("NominalAccount").ToString
                    Else
                      pList("ProductNominalAccount") = vDataRow("SubsequentNominalAccount").ToString
                    End If
                  End If
                End If
            End Select

          Case "scoring_details"
            If vPanelItem.AttributeName = "search_area" Then pList("ApplicationName") = "SM"

          Case "sources"
            If vPanelItem.AttributeName = "incentive_scheme" Then pList("Active") = "Y"

          Case "statuses"
            If vPanelItem.AttributeName = "activity_group" Then pList("UsageCode") = "S"

          Case "vat_rate_identification"
            If vPanelItem.AttributeName = "vat_rate" Then pList("ForMaintenance") = "Y" 'NFPCARE-526
          Case "config"
            If vPanelItem.AttributeName.Equals("config_value", StringComparison.InvariantCultureIgnoreCase) Then
              Select Case vPanelItem.ValidationTable.ToLower
                Case "document_types"
                  pList("IncludeEmailDocSource") = "Y"

                Case "fp_applications"
                  If mvCriteria IsNot Nothing AndAlso mvCriteria.ContainsKey("ConfigName") Then
                    If mvCriteria("ConfigName").ToLower.Equals("trader_application_edit_trans") Then
                      pList("Filter") = "fp_application_type = 'TRANS'"
                    End If
                  End If
              End Select
            End If
        End Select

        Select Case vPanelItem.ValidationTable
          Case "distribution_codes"
            pList("Active") = "Y"
          Case "users"
            If vPanelItem.AttributeName = "logname" Then pList("Active") = "Y"
        End Select

      End If
    Else
      If mvTableName = "organisation_groups" And pParameterName = "OrganisationNumber" Then
        If mvParams.Contains("ViewInContactCard") AndAlso mvParams("ViewInContactCard") = "Y" Then
          If mvParams.Contains("OrganisationGroup") Then
            pList("OrganisationGroupCode") = mvParams("OrganisationGroup")
          End If
        End If
      End If
    End If
  End Sub

  Private Sub frmTableEntry_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    Try
      If Not mvEditMode = CareNetServices.XMLTableMaintenanceMode.xtmmSelect AndAlso epl.DataChanged AndAlso Not mvCancel Then
        mvInternalResourceMesgShown = False
        If ConfirmSave() Then
          e.Cancel = Not ProcessSave()
          Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End If
      End If
    Catch vException As CareException
      If vException.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        ShowInformationMessage(InformationMessages.ImRecordAlreadyExists)
      Else
        DataHelper.HandleException(vException)
      End If
      e.Cancel = True
    End Try
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    If pParameterName = "GeographicalRegion" Then
      If pList.ContainsKey("CollectionNumber") Then pList.Remove("CollectionNumber")
      pList("GeographicalRegionType") = epl.GetValue("GeographicalRegionType")
    ElseIf mvTableName = "prize_draws" AndAlso pParameterName = "Product" Then
      pList("FindProductType") = "R"
    ElseIf mvTableName = "gaye_controls" Then
      Select Case pParameterName
        Case "AdminFeeProduct", "DonorProduct", "EmployerProduct", "GovernmentProduct", "OtherMatchedProduct"
          'Pre-tax Payroll Giving Products
          pList("FindProductType") = "F"    'All flags except donation should be N
        Case "AdminFeeRate", "DonorRate", "EmployerRate", "GovernmentRate", "OtherMatchedRate"
          'Pre-tax Payroll Giving Rates
          pList("CurrentPrice") = "0"
        Case "PostTaxDonorProduct"
          pList("FindProductType") = "X"    'Post Tax Donor Product
        Case "PostTaxEmployerProduct"
          pList("FindProductType") = "T"    'Post Tax Employer Product
        Case "PostTaxDonorRate", "PostTaxEmployerRate"
          pList("CurrentPrice") = "0"
      End Select
    ElseIf mvTableName = "event_pricing_matrices" Then
      Select Case pParameterName
        Case "AdultFeeProduct", "ChildFeeProduct", "ExtraSessionFeeProduct"
          pList("FindProductType") = "V"
        Case "AdultFeeRate", "ChildFeeRate"
          pList("CurrentPrice") = "0"
          pList("VatExclusive") = "Y"
        Case "ExtraSessionFeeRate"
          pList("ExtraSessionFeeRate") = "Y"
      End Select
    ElseIf pParameterName = "VatRates" Then
      pList("ForMaintenance") = "Y"
    ElseIf pParameterName = "CapitalProduct" Then
      pList("FindProductType") = "L"
    ElseIf pParameterName = "InterestProduct" OrElse pParameterName = "InvoiceUnderPaymentProduct" OrElse pParameterName = "InvoiceOverPaymentProduct" _
    OrElse pParameterName = "FirstPeriodsRefundProduct" OrElse pParameterName = "SubsequentRefundProduct" OrElse pParameterName = "CmtRefundProduct" Then
      pList("FindProductType") = "Z"
    ElseIf pParameterName = "FirstPeriodsRefundRate" OrElse pParameterName = "SubsequentRefundRate" OrElse pParameterName = "CmtRefundRate" Then
      pList("CurrentPrice") = "0"
    Else
      'This could be some sort of finder. check if there are any restrictions
      'that need to be applied. eg. in activity_groups table the campaign and appeal
      'are both displayed as finders. we need to restrict the appeal to the campaign 
      'that is selected.
      Dim vControl As Control = epl.FindPanelControl(pParameterName, False)
      If vControl IsNot Nothing Then 'ideally this should never be nothing
        Dim vPanelItem As PanelItem = DirectCast(vControl.Tag, PanelItem)
        If vPanelItem.RestrictionAttribute.Length > 0 Then
          Dim vAttr As String = vPanelItem.RestrictionAttribute
          Dim vRestControl As Control = epl.FindPanelControl(ProperName(vAttr))
          Dim vValue As String = epl.GetValue(ProperName(vAttr))
          If vRestControl IsNot Nothing Then
            Dim vRestItem As PanelItem = DirectCast(vRestControl.Tag, PanelItem)
            If vRestItem.ValidationAttribute.Length > 0 Then vAttr = vRestItem.ValidationAttribute
          End If
          If vValue.Length > 0 Then pList(ProperName(vAttr)) = vValue
        End If
      End If
      'These need to be applied as well as the restriction attributes (above)
      Select Case pParameterName
        Case "CapitalRate"
          If mvTableName = "loan_types" Then
            pList("CurrentPrice") = "0"
            pList("VatExclusive") = "N"
          End If
        Case "InterestProductRate"
          If mvTableName = "loan_types" Then
            pList("LoanInterest") = "Y"
            pList("VatExclusive") = "N"
          End If
      End Select
    End If
  End Sub

  Private Sub epl_ValidateItem(ByVal sender As Object, ByVal pParameterName As String, ByVal pValue As String, ByRef vValid As Boolean) Handles epl.ValidateItem
    Select Case mvTableName
      Case "financial_controls"
        If pParameterName = "LoanCapitalisationDate" Then
          epl.SetErrorField(pParameterName, "")
          If pValue.Length <> 4 Then
            vValid = False
          Else
            Dim vDate As Date = DateSerial(Year(Now), IntegerValue(pValue.Substring(2)), IntegerValue(pValue.Substring(0, 2)))
            If vDate.Day <> IntegerValue(pValue.Substring(0, 2)) OrElse vDate.Month <> IntegerValue(pValue.Substring(2)) OrElse vDate.Year <> Year(Now) Then vValid = False 'Value was not in correct dd/mm format
          End If
          If vValid = False Then epl.SetErrorField(pParameterName, InformationMessages.ImControlTableDateInvalid, False)
        End If

      Case "gift_aid_controls"
        Select Case pParameterName
          Case "AccountingPeriodStart", "TaxYearStart"
            epl.SetErrorField(pParameterName, "")
            If pValue.Length <> 4 Then
              vValid = False
            Else
              Dim vDate As Date = DateSerial(Year(Now), IntegerValue(pValue.Substring(2)), IntegerValue(pValue.Substring(0, 2)))
              If vDate.Day <> IntegerValue(pValue.Substring(0, 2)) OrElse vDate.Month <> IntegerValue(pValue.Substring(2)) OrElse vDate.Year <> Year(Now) Then vValid = False 'Value was not in correct dd/mm format
            End If
            If vValid = False Then epl.SetErrorField(pParameterName, InformationMessages.ImControlTableDateInvalid, False)
          Case "SubmitterContact"
            If pValue = "O" AndAlso epl.PanelInfo.PanelItems.Exists("SubmitterContact") AndAlso _
              epl.GetValue("SubmitterContact").Length = 0 Then
              epl.SetErrorField("SubmitterContact", InformationMessages.ImFieldMandatory)
              vValid = False
            Else
              epl.SetErrorField("SubmitterContact", "")
            End If
        End Select

      Case "membership_type_categories"
        Select Case pParameterName
          Case "MembershipType"
            If pValue.Length = 0 Then
              epl.SetErrorField(pParameterName, InformationMessages.ImFieldMandatory)    'Field is mandatory and cannot be left empty
              vValid = False
            End If
          Case "Activity"
            If pValue.Length = 0 Then
              epl.SetErrorField(pParameterName, InformationMessages.ImFieldMandatory)    'Field is mandatory and cannot be left empty
              vValid = False
            End If
        End Select

      Case "contact_group_users"
        If pParameterName = "Department" Then
          If pValue.Length = 0 Then
            epl.SetErrorField(pParameterName, InformationMessages.ImFieldMandatory)    'Field is mandatory and cannot be left empty
            vValid = False
          End If
        End If

      Case "rate_modifiers"
        If pParameterName = "SequenceNumber" OrElse pParameterName = "NextSequenceNumber" Then
          Dim vSequenceNumber As Integer = 0
          Dim vNextSequenceNumber As Integer = 0
          Dim vStopIfNoModifiers As Boolean = False
          Select Case pParameterName
            Case "SequenceNumber"
              Integer.TryParse(pValue, vSequenceNumber)
              Integer.TryParse(epl.GetValue("NextSequenceNumber"), vNextSequenceNumber)
            Case "NextSequenceNumber"
              Integer.TryParse(epl.GetValue("SequenceNumber"), vSequenceNumber)
              Integer.TryParse(pValue, vNextSequenceNumber)
          End Select
          If vNextSequenceNumber > 0 Then
            If vSequenceNumber = 0 Then
              epl.SetErrorField(pParameterName, InformationMessages.ImSequenceNumberBlankError)
              vValid = False
            ElseIf vSequenceNumber >= vNextSequenceNumber Then
              epl.SetErrorField(pParameterName, InformationMessages.ImSequenceNumberError)
              vValid = False
            End If
          End If
        ElseIf pParameterName = "StopIfNoModifiers" AndAlso pValue = "Y" AndAlso String.IsNullOrWhiteSpace(epl.GetValue("SequenceNumber")) Then
          epl.SetErrorField(pParameterName, InformationMessages.ImStopIfNoModifiersInvalid)
          vValid = False
        End If
      Case "bank_accounts"
        If pParameterName = "IbanNumber" AndAlso pValue.Length > 0 Then
          Try
            epl.SetErrorField("IbanNumber", "")
            DataHelper.CheckIbanNumber(pValue)
          Catch vException As Exception
            epl.SetErrorField("IbanNumber", GetInformationMessage(vException.Message))
            vValid = False
          End Try
        End If

      Case "bank_account_claim_days"
        If pParameterName.Equals("ClaimDay") Then
          epl.SetErrorField(pParameterName, String.Empty)
          If epl.GetValue("ClaimType").Equals("DD", StringComparison.InvariantCultureIgnoreCase) AndAlso AppValues.ConfigurationOption(AppValues.ConfigurationOptions.fp_dd_fixed_claim_date, False) Then
            If IntegerValue(pValue) > 28 Then
              vValid = epl.SetErrorField(pParameterName, InformationMessages.ImInvalidDDClaimDay, True)
            End If
          End If
        End If
    End Select
  End Sub

  Private Sub HandlePackageTableChange(sender As Object, pParameterName As String, pValue As String)
    If pParameterName = "DocumentSource" Then
      SetPackageTableFieldAvailabilty()
    End If
  End Sub

  Private Sub SetPackageTableFieldAvailabilty()
    Dim vDocumentSource As TextLookupBox = epl.FindTextLookupBox("DocumentSource", False)
    If vDocumentSource IsNot Nothing Then
      Dim vPackagePath As TextBox = DirectCast(FindControl(epl, "PackagePath", False), TextBox)
      Dim vDocfileExtension As TextBox = DirectCast(FindControl(epl, "DocfileExtension", False), TextBox)
      Dim vStorageType As TextLookupBox = epl.FindTextLookupBox("StorageType", False)
      Dim vStoragePath As TextBox = DirectCast(FindControl(epl, "StoragePath", False), TextBox)
      Dim vPackageRunstring As TextBox = DirectCast(FindControl(epl, "PackageRunstring", False), TextBox)
      Dim vCommunicationType As ComboBox = DirectCast(FindControl(epl, "CommunicationType", False), ComboBox)
      Dim vCommunicationVersion As ComboBox = DirectCast(FindControl(epl, "CommunicationVersion", False), ComboBox)
      Dim vTopicOrObjectName As TextBox = DirectCast(FindControl(epl, "TopicOrObjectName", False), TextBox)
      Dim vClassName As TextBox = DirectCast(FindControl(epl, "ClassName", False), TextBox)
      If vDocumentSource.Text = "E" Then
        If vPackagePath IsNot Nothing Then
          vPackagePath.Text = String.Empty
          vPackagePath.Enabled = False
        End If
        If vDocfileExtension IsNot Nothing Then
          vDocfileExtension.Text = String.Empty
          vDocfileExtension.Enabled = False
        End If
        If vStorageType IsNot Nothing Then
          vStorageType.Text = "I"
          vStorageType.Enabled = False
        End If
        If vStoragePath IsNot Nothing Then
          vStoragePath.Text = String.Empty
          vStoragePath.Enabled = False
        End If
        If vPackageRunstring IsNot Nothing Then
          vPackageRunstring.Text = String.Empty
          vPackageRunstring.Enabled = False
        End If
        If vCommunicationType IsNot Nothing Then
          vCommunicationType.SelectedIndex = -1
          vCommunicationType.Enabled = False
        End If
        If vCommunicationVersion IsNot Nothing Then
          vCommunicationVersion.SelectedIndex = -1
          vCommunicationVersion.Enabled = False
        End If
        If vTopicOrObjectName IsNot Nothing Then
          vTopicOrObjectName.Text = String.Empty
          vTopicOrObjectName.Enabled = False
        End If
        If vClassName IsNot Nothing Then
          vClassName.Text = String.Empty
          vClassName.Enabled = False
        End If
      Else
        If vPackagePath IsNot Nothing Then
          vPackagePath.Enabled = True
        End If
        If vDocfileExtension IsNot Nothing Then
          vDocfileExtension.Enabled = True
        End If
        If vStorageType IsNot Nothing Then
          vStorageType.Enabled = True
        End If
        If vStoragePath IsNot Nothing Then
          vStoragePath.Enabled = True
        End If
        If vPackageRunstring IsNot Nothing Then
          vPackageRunstring.Enabled = True
        End If
        If vCommunicationType IsNot Nothing Then
          vCommunicationType.Enabled = True
        End If
        If vCommunicationVersion IsNot Nothing Then
          vCommunicationVersion.Enabled = True
        End If
        If vTopicOrObjectName IsNot Nothing Then
          vTopicOrObjectName.Enabled = True
        End If
        If vClassName IsNot Nothing Then
          vClassName.Enabled = True
        End If
      End If
    End If

  End Sub

  Private Sub CheckFutureMembershipChange()
    Dim vNewProductRate As String = String.Format("{0}.{1}", epl.GetValue("FirstPeriodsProduct"), epl.GetValue("FirstPeriodsRate"))
    Dim vOldProductRate As String = String.Format("{0}.{1}", mvParams("FirstPeriodsProduct"), mvParams("FirstPeriodsRate"))
    If Not String.Equals(vOldProductRate, vNewProductRate, StringComparison.CurrentCultureIgnoreCase) Then
      Dim vMembList As New ParameterList(True, True)
      vMembList.Add("MembershipType", mvParams("MembershipType"))
      vMembList.Add("Product", mvParams("FirstPeriodsProduct"))
      vMembList.Add("Rate", mvParams("FirstPeriodsRate"))
      If DataHelper.GetCount(CareNetServices.XMLGetCountTypes.xgctFutureMembershipTypes, vMembList) > 0 Then
        Dim vResponse As DialogResult = ShowQuestion(QuestionMessages.QmMembershipTypeProductRateInUse, MessageBoxButtons.YesNo)
        If vResponse = System.Windows.Forms.DialogResult.Yes Then
          CascadeUpdateFMTChange()
        End If
      End If
    End If

  End Sub

  Private Sub CascadeUpdateFMTChange()
    Dim vParams As New ParameterList(True, True)
    vParams.Add("MembershipType", epl.GetValue("MembershipType"))
    vParams.Add("Product", mvParams("FirstPeriodsProduct"))
    vParams.Add("Rate", mvParams("FirstPeriodsRate"))
    FormHelper.ProcessTask(CareNetServices.TaskJobTypes.tjtUpdateFutureMembershipType, vParams)
  End Sub

  Private Sub SetPaymentFrequencyOffset(ByVal pSetDisabledValue As Boolean)
    epl.SetErrorField("OffsetMonths", "")
    If mvTableName.Equals("payment_frequencies", StringComparison.InvariantCultureIgnoreCase) AndAlso FindControl(epl, "OffsetMonths", False) IsNot Nothing Then
      Dim vEnableOffset As Boolean = False
      Dim vFrequency As Integer = IntegerValue(epl.GetValue("Frequency"))
      Dim vInterval As Integer = IntegerValue(epl.GetValue("Interval"))
      Dim vPeriod As String = epl.GetValue("Period")
      If Not String.IsNullOrWhiteSpace(vPeriod) Then
        If vPeriod.Equals("M", StringComparison.InvariantCultureIgnoreCase) Then
          If (vFrequency * vInterval) <= 12 Then vEnableOffset = True
          If (vFrequency * vInterval) = 12 AndAlso vFrequency = 12 Then vEnableOffset = False 'Disable for monthly instalments
          If (vFrequency * vInterval) < 12 AndAlso vFrequency = 1 Then vEnableOffset = False 'Disable for regular instalments
        End If
        If vEnableOffset = False AndAlso pSetDisabledValue = True Then epl.SetValue("OffsetMonths", "0")
        epl.EnableControl("OffsetMonths", vEnableOffset)

        If vEnableOffset Then
          Dim vMaxOffset As Integer = (vInterval - 1)
          If (vFrequency * vInterval) < 12 Then vMaxOffset = (12 - (vFrequency * vInterval))
          Dim vOffsetControl As TextBox = epl.FindTextBox("OffsetMonths")
          If vOffsetControl.Tag IsNot Nothing AndAlso TypeOf (vOffsetControl.Tag) Is PanelItem Then
            CType(vOffsetControl.Tag, PanelItem).MaximumValue = vMaxOffset.ToString
          End If
        End If
      End If
    End If
  End Sub

End Class