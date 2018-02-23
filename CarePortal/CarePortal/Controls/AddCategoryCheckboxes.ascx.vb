Public Partial Class AddCategoryCheckboxes
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvHandlesActivities = True
    mvUsesHiddenContactNumber = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddCategoryCheckboxes, tblDataEntry)
      AddHiddenField("OldActivityValues")  'To check if there is an existing record
      AddHiddenField("DisplayedActivityDate")
      AddHiddenField("DisplayedQuantity")
      AddHiddenField("DisplayedNotes")
      AddHiddenField("DisplayedValidFrom")
      AddHiddenField("DisplayedValidTo")
      AddHiddenField("OldValidFromDates")
      AddHiddenField("OldValidToDates")
      AddHiddenField("OldSources")
      AddHiddenField("DelegateActivityNumber")
      SetDefaultDates()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessActivitySelection(ByVal pTable As DataTable)
    Dim vDisplayActivity As Boolean
    If pTable IsNot Nothing Then
      Dim vActivityValues As New StringBuilder
      Dim vValidFromDates As New StringBuilder
      Dim vValidToDates As New StringBuilder
      Dim vSources As New StringBuilder
      Dim vDelegateActivityNumbers As New StringBuilder
      Dim vAddValue As Boolean
      For Each vRow As DataRow In pTable.Rows
        Dim vValidFrom As Date = Date.Parse(vRow("ValidFrom").ToString)
        Dim vValidTo As Date = Date.Parse(vRow("ValidTo").ToString)
        Dim vAmendedOn As Date = Date.Parse(vRow("AmendedOn").ToString)
        If vRow("ActivityCode").ToString = DefaultParameters("Activity").ToString AndAlso _
          vValidFrom <= Date.Today AndAlso vValidTo >= Date.Today AndAlso Not (vValidTo = vAmendedOn) Then

          vAddValue = True
          For Each vValue As String In vActivityValues.ToString.Split(","c)
            If vValue = vRow("ActivityValueCode").ToString Then
              vAddValue = False
              Exit For
            End If
          Next
          If vAddValue Then
            If vActivityValues.Length > 0 Then
              vActivityValues.Append(",")
              vValidFromDates.Append(",")
              vValidToDates.Append(",")
              vSources.Append(",")
              If vDelegateActivityNumbers.Length > 0 Then vDelegateActivityNumbers.Append(",")
            End If
            SetCheckBoxChecked("ActivityValue_" & vRow("ActivityValueCode").ToString)
            SetTextBoxText("ActivityDate", vRow("ActivityDate").ToString)
            SetTextBoxText("Quantity", vRow("Quantity").ToString)
            SetTextBoxText("Notes", vRow("Notes").ToString)
            SetTextBoxText("ValidFrom", vRow("ValidFrom").ToString)
            SetTextBoxText("ValidTo", vRow("ValidTo").ToString)

            vActivityValues.Append(vRow("ActivityValueCode").ToString)
            vValidFromDates.Append(vRow("ValidFrom").ToString)
            vValidToDates.Append(vRow("ValidTo").ToString)
            vSources.Append(vRow("SourceCode").ToString)
            If pTable.Columns.Contains("DelegateActivityNumber") AndAlso vRow("DelegateActivityNumber") IsNot Nothing AndAlso Not String.IsNullOrEmpty(vRow("DelegateActivityNumber").ToString) Then
              vDelegateActivityNumbers.Append(vRow("DelegateActivityNumber").ToString)
            End If

            SetHiddenText("DisplayedActivityDate", vRow("ActivityDate").ToString)
            SetHiddenText("DisplayedQuantity", vRow("Quantity").ToString)
            SetHiddenText("DisplayedNotes", vRow("Notes").ToString)
            SetHiddenText("DisplayedValidFrom", vRow("ValidFrom").ToString)
            SetHiddenText("DisplayedValidTo", vRow("ValidTo").ToString)
            vDisplayActivity = True
          End If
        End If
      Next
      If vActivityValues.Length > 0 Then
        SetHiddenText("OldActivityValues", vActivityValues.ToString)
        SetHiddenText("OldValidFromDates", vValidFromDates.ToString)
        SetHiddenText("OldValidToDates", vValidToDates.ToString)
        SetHiddenText("OldSources", vSources.ToString)
        SetHiddenText("DelegateActivityNumber", vDelegateActivityNumbers.ToString)
      End If
      ShowActivityOrSuppression(vDisplayActivity)
    End If
  End Sub

  Public Overrides Sub DisplayActivitySuppressionModule(ByVal pValue As Boolean)
    ShowActivityOrSuppression(pValue)
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    Dim vContactFound As Boolean = vContactNumber > 0
    If Not vContactFound Then vContactNumber = IntegerValue(pList("ContactNumber").ToString)

    Dim vOldValues() As String = GetHiddenText("OldActivityValues").Split(","c)
    Dim vOldValidFromDates() As String = GetHiddenText("OldValidFromDates").Split(","c)
    Dim vOldValidToDates() As String = GetHiddenText("OldValidToDates").Split(","c)
    Dim vOldSources() As String = GetHiddenText("OldSources").Split(","c)
    Dim vDelegateActivityNumber() As String = GetHiddenText("DelegateActivityNumber").Split(","c)
    Dim vIndex As Integer
    For Each vCheckBox As CheckBox In mvCheckBoxes
      Dim vSetHistoric As Boolean = BooleanValue(DefaultParameters("SetHistoric").ToString)
      Dim vAddNew As Boolean = True
      vIndex = Array.IndexOf(vOldValues, vCheckBox.ID.Substring(vCheckBox.ID.IndexOf("_") + 1))
      If vContactFound Then
        Dim vDataChanged As Boolean = False
        If vIndex < 0 AndAlso vCheckBox.Checked Then
          'Add New Record
        ElseIf vIndex >= 0 AndAlso vCheckBox.Checked = False Then
          'The record is removed
          vAddNew = False
          vDataChanged = True
        ElseIf Not (ValueChanged("DisplayedActivityDate", "ActivityDate") _
            OrElse ValueChanged("DisplayedValidFrom", "ValidFrom") OrElse ValueChanged("DisplayedValidTo", "ValidTo") _
            OrElse ValueChanged("DisplayedQuantity", "Quantity") OrElse ValueChanged("DisplayedNotes", "Notes")) OrElse vCheckBox.Checked = False Then
          'Nothing is changed. Don't add/delete anything - unless it was set by default for delegate activities
          If GetHiddenText("DelegateActivityNumber").Length = 0 AndAlso ParentGroup = "DelegateActivities" Then
            vAddNew = True
          Else
            vAddNew = False
          End If
        Else
          'Data is changed. Set the existing as historic and add new record
          vDataChanged = True
          'contact related data changed which is shown default to current contact not event delegate data
          If GetHiddenText("DelegateActivityNumber").Length = 0 AndAlso ParentGroup = "DelegateActivities" Then vDataChanged = False
        End If
        If vDataChanged Then
          If vSetHistoric Then
            'Check if there is a similar record with ValidTo and AmendedOn date set as current date
            Dim vList As New ParameterList(HttpContext.Current)
            vList("ContactNumber") = vContactNumber
            AddUserParameters(vList)
            vList("Activity") = DefaultParameters("Activity")
            vList("ActivityValue") = vOldValues.GetValue(vIndex)
            vList("Source") = vOldSources.GetValue(vIndex)
            vList("ValidTo") = TodaysDate() '
            vList("AmendedOn") = TodaysDate() '
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vList)
            If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
              'If a similar record is found then delete the current record (not the one we just found)
              vList("ValidFrom") = vOldValidFromDates.GetValue(vIndex)
              vList("ValidTo") = vOldValidToDates.GetValue(vIndex)
              vList.Remove("AmendedOn")
              If ParentGroup = "DelegateActivities" Then
                If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
                If vDelegateActivityNumber.Length > vIndex Then
                  vList("DelegateActivityNumber") = vDelegateActivityNumber(vIndex)
                  If IntegerValue(vDelegateActivityNumber(vIndex)) > 0 Then DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
                End If
              Else
                DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
              End If
              vDataChanged = False
            End If
            If vDataChanged Then
              'Set the current as historic
              vList = New ParameterList(HttpContext.Current)
              vList("OldContactNumber") = vContactNumber
              If Convert.ToString(vList("UserLogname")).Length = 0 Then
                vList("UserID") = vContactNumber
              End If
              vList("OldActivity") = DefaultParameters("Activity")
              vList("OldActivityValue") = vOldValues.GetValue(vIndex).ToString
              vList("OldSource") = vOldSources.GetValue(vIndex)
              vList("OldValidFrom") = vOldValidFromDates.GetValue(vIndex)
              vList("OldValidTo") = vOldValidToDates.GetValue(vIndex)
              vList("ValidTo") = TodaysDate()
              vList("AmendedOn") = TodaysDate()
              vList("CarePortal") = "Y" 'BR14165: Pass this to not validate the dates
              If ParentGroup = "DelegateActivities" Then
                If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
                If vDelegateActivityNumber.Length > vIndex Then
                  vList("DelegateActivityNumber") = vDelegateActivityNumber(vIndex)
                  If IntegerValue(vDelegateActivityNumber(vIndex)) > 0 Then DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
                End If
              Else
                DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
              End If
            End If
          Else
            'Delete
            Dim vList As New ParameterList(HttpContext.Current)
            vList("ContactNumber") = vContactNumber
            AddUserParameters(vList)
            vList("Activity") = DefaultParameters("Activity")
            vList("ActivityValue") = vOldValues.GetValue(vIndex)
            vList("KeepHistoricActivity") = "Y"
            If ParentGroup = "DelegateActivities" Then
              If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
              If vDelegateActivityNumber.Length > vIndex Then
                vList("DelegateActivityNumber") = vDelegateActivityNumber(vIndex)
                If IntegerValue(vDelegateActivityNumber(vIndex)) > 0 Then DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
              End If
            Else
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
            End If
          End If
        End If
      End If
      If vAddNew AndAlso vCheckBox.Checked Then
        'Add New
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vContactNumber
        AddUserParameters(vList)
        vList("Activity") = DefaultParameters("Activity")
        vList("ActivityValue") = vCheckBox.ID.Substring(vCheckBox.ID.IndexOf("_") + 1)
        vList("Source") = DefaultParameters("Source")
        vList("ValidFrom") = SetDate(DateType.ValidFrom) ' Today.ToShortDateString
        'valid to : check if activity has duration, if not set it to Today.AddYears(100).ToShortDateString
        vList("ValidTo") = SetCategoryValidToDate(DateType.ValidTo, vList("ValidFrom").ToString, vList("Activity").ToString, vList("ActivityValue").ToString)
        AddOptionalTextBoxValue(vList, "ActivityDate")
        AddOptionalTextBoxValue(vList, "Quantity")
        AddOptionalTextBoxValue(vList, "Notes")
        If vSetHistoric Then vList("CarePortal") = "Y" 'BR14165: Pass this to not extend the historic record with this new record 
        If ParentGroup = "DelegateActivities" Then
          If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
          DataHelper.AddDelegateActivity(vList)
        Else
          DataHelper.AddActivity(vList)
        End If
      End If
    Next
  End Sub

  Public Overrides Sub ClearControls()
    ClearControls(False)
  End Sub

  Public Overrides Sub ClearControls(ByVal pClearLabels As Boolean)
    MyBase.ClearControls(pClearLabels)
    ClearHiddenControls()
    SetDefaultDates()
  End Sub

  Private Sub ClearHiddenControls()
    SetHiddenText("OldActivityValues", String.Empty)
    SetHiddenText("DisplayedActivityDate", String.Empty)
    SetHiddenText("DisplayedQuantity", String.Empty)
    SetHiddenText("DisplayedNotes", String.Empty)
    SetHiddenText("DisplayedValidFrom", String.Empty)
    SetHiddenText("DisplayedValidTo", String.Empty)
    SetHiddenText("OldValidFromDates", String.Empty)
    SetHiddenText("OldValidToDates", String.Empty)
    SetHiddenText("OldSources", String.Empty)
    SetHiddenText("DelegateActivityNumber", String.Empty)
  End Sub

  Protected Overrides Sub AddCustomValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)
    AddCustomValidator(pHTMLCell, pID, "Required Field")
  End Sub

  Public Overrides Sub ServerValidate(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
    Dim vCFV As CustomValidator = DirectCast(sender, CustomValidator)
    If vCFV IsNot Nothing AndAlso vCFV.ID.Length > 0 Then
      Dim vChkBox As CheckBox = DirectCast(FindControlByName(Me, vCFV.ID.Substring(3)), CheckBox)
      e.IsValid = vChkBox.Checked
    Else
      e.IsValid = False
    End If
  End Sub

End Class