Public Partial Class AddCategoryOptions
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvHandlesActivities = True
    mvUsesHiddenContactNumber = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddCategoryOptions, tblDataEntry)
      AddHiddenField("OldActivityValue")  'To check if there is an existing record
      AddHiddenField("OldActivityDate")
      AddHiddenField("OldQuantity")
      AddHiddenField("OldNotes")
      AddHiddenField("OldValidFrom")
      AddHiddenField("OldValidTo")
      AddHiddenField("OldSource")
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
      For Each vRow As DataRow In pTable.Rows
        Dim vValidFrom As Date = Date.Parse(vRow("ValidFrom").ToString)
        Dim vValidTo As Date = Date.Parse(vRow("ValidTo").ToString)
        Dim vAmendedOn As Date = Date.Parse(vRow("AmendedOn").ToString)
        If vRow("ActivityCode").ToString = DefaultParameters("Activity").ToString AndAlso _
          vValidFrom <= Date.Today AndAlso vValidTo >= Date.Today AndAlso Not (vValidTo = vAmendedOn) Then
          SetRadioButtonChecked("ActivityValue_" & vRow("ActivityValueCode").ToString)
          SetTextBoxText("ActivityDate", vRow("ActivityDate").ToString)
          SetTextBoxText("Quantity", vRow("Quantity").ToString)
          SetTextBoxText("Notes", vRow("Notes").ToString)
          SetTextBoxText("ValidFrom", vRow("ValidFrom").ToString)
          SetTextBoxText("ValidTo", vRow("ValidTo").ToString)

          SetHiddenText("OldActivityValue", vRow("ActivityValueCode").ToString)
          SetHiddenText("OldActivityDate", vRow("ActivityDate").ToString)
          SetHiddenText("OldQuantity", vRow("Quantity").ToString)
          SetHiddenText("OldNotes", vRow("Notes").ToString)
          SetHiddenText("OldValidFrom", vRow("ValidFrom").ToString)
          SetHiddenText("OldValidTo", vRow("ValidTo").ToString)
          SetHiddenText("OldSource", vRow("SourceCode").ToString)
          If pTable.Columns.Contains("DelegateActivityNumber") AndAlso vRow("DelegateActivityNumber") IsNot Nothing AndAlso Not String.IsNullOrEmpty(vRow("DelegateActivityNumber").ToString) Then
            SetHiddenText("DelegateActivityNumber", vRow("DelegateActivityNumber").ToString)
          End If
          vDisplayActivity = True
          Exit For
        End If
      Next
    End If
    ShowActivityOrSuppression(vDisplayActivity)
  End Sub

  Public Overrides Sub DisplayActivitySuppressionModule(ByVal pValue As Boolean)
    ShowActivityOrSuppression(pValue)
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    Dim vContactFound As Boolean = vContactNumber > 0
    If Not vContactFound Then vContactNumber = IntegerValue(pList("ContactNumber").ToString)

    Dim vOldValues() As String = GetHiddenText("OldActivityValue").Split(","c)
    For Each vRadioButton As RadioButton In mvRadioButtons
      Dim vSetHistoric As Boolean = BooleanValue(DefaultParameters("SetHistoric").ToString)
      Dim vAddNew As Boolean = True
      Dim vIndex As Integer = Array.IndexOf(vOldValues, vRadioButton.ID.Substring(vRadioButton.ID.IndexOf("_") + 1))
      If vContactFound Then
        Dim vDataChanged As Boolean = False
        If vIndex < 0 AndAlso vRadioButton.Checked Then
          'Add New Record
        ElseIf vIndex >= 0 AndAlso vRadioButton.Checked = False Then
          'The record is removed
          vAddNew = False
          vDataChanged = True
        ElseIf Not (ValueChanged("OldActivityDate", "ActivityDate") _
            OrElse ValueChanged("OldValidFrom", "ValidFrom") OrElse ValueChanged("OldValidTo", "ValidTo") _
            OrElse ValueChanged("OldQuantity", "Quantity") OrElse ValueChanged("OldNotes", "Notes")) OrElse vRadioButton.Checked = False Then
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
            vList("Source") = GetHiddenText("OldSource")
            vList("ValidTo") = TodaysDate() '
            vList("AmendedOn") = TodaysDate() '
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vList)
            If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
              'If a similar record is found then delete the current record (not the one we just found)
              vList("ValidFrom") = GetHiddenText("OldValidFrom")
              vList("ValidTo") = GetHiddenText("OldValidTo")
              vList.Remove("AmendedOn")
              If ParentGroup = "DelegateActivities" Then
                If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
                If Not String.IsNullOrEmpty(GetHiddenText("DelegateActivityNumber")) Then
                  vList("DelegateActivityNumber") = GetHiddenText("DelegateActivityNumber")
                  DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
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
              vList("OldSource") = GetHiddenText("OldSource")
              vList("OldValidFrom") = GetHiddenText("OldValidFrom")
              vList("OldValidTo") = GetHiddenText("OldValidTo")
              vList("ValidTo") = TodaysDate()
              vList("AmendedOn") = TodaysDate()
              vList("CarePortal") = "Y" 'BR14165: Pass this to not validate the dates
              If ParentGroup = "DelegateActivities" Then
                If pList.Contains("EventDelegateNumber") Then vList("EventDelegateNumber") = pList("EventDelegateNumber")
                If Not String.IsNullOrEmpty(GetHiddenText("DelegateActivityNumber")) Then vList("DelegateActivityNumber") = GetHiddenText("DelegateActivityNumber")
                DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
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
              If Not String.IsNullOrEmpty(GetHiddenText("DelegateActivityNumber")) Then
                vList("DelegateActivityNumber") = GetHiddenText("DelegateActivityNumber")
                DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctEventDelegateActivity, vList)
              End If
            Else
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
            End If
          End If
        End If
      End If
      If vAddNew AndAlso vRadioButton.Checked Then
        'Add New
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vContactNumber
        AddUserParameters(vList)
        vList("Activity") = DefaultParameters("Activity")
        vList("ActivityValue") = vRadioButton.ID.Substring(vRadioButton.ID.IndexOf("_") + 1)
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
    SetHiddenText("OldActivityValue", String.Empty)
    SetHiddenText("OldActivityDate", String.Empty)
    SetHiddenText("OldQuantity", String.Empty)
    SetHiddenText("OldNotes", String.Empty)
    SetHiddenText("OldValidFrom", String.Empty)
    SetHiddenText("OldValidTo", String.Empty)
    SetHiddenText("OldSource", String.Empty)
    SetHiddenText("DelegateActivityNumber", String.Empty)
  End Sub
End Class