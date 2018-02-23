Partial Public Class AddSuppression
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvHandlesSuppressions = True
    mvUsesHiddenContactNumber = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddSuppression, tblDataEntry)
      AddHiddenField("OldSuppression")  'To check if there is an existing record
      AddHiddenField("OldNotes")
      AddHiddenField("OldValidFrom")
      AddHiddenField("OldValidTo")
      SetDefaultDates()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessSuppressionSelection(ByVal pTable As DataTable)
    Dim vDisplaySuppression As Boolean
    If pTable IsNot Nothing Then
      For Each vRow As DataRow In pTable.Rows
        Dim vValidFrom As Date = Date.Parse(vRow("ValidFrom").ToString)
        Dim vValidTo As Date = Date.Parse(vRow("ValidTo").ToString)
        Dim vAmendedOn As Date = Date.Parse(vRow("AmendedOn").ToString)
        If vRow("SuppressionCode").ToString = DefaultParameters("Suppression").ToString AndAlso _
          vValidFrom <= Date.Today AndAlso vValidTo >= Date.Today AndAlso Not (vValidTo = vAmendedOn) Then
          SetCheckBoxChecked("Suppression")
          SetTextBoxText("Notes", vRow("Notes").ToString)
          SetTextBoxText("ValidFrom", vRow("ValidFrom").ToString)
          SetTextBoxText("ValidTo", vRow("ValidTo").ToString)

          SetHiddenText("OldSuppression", vRow("SuppressionCode").ToString)
          SetHiddenText("OldNotes", vRow("Notes").ToString)
          SetHiddenText("OldValidFrom", vRow("ValidFrom").ToString)
          SetHiddenText("OldValidTo", vRow("ValidTo").ToString)
          vDisplaySuppression = True
          Exit For
        End If
      Next
    End If
    ShowActivityOrSuppression(vDisplaySuppression)
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vSetHistoric As Boolean = BooleanValue(DefaultParameters("SetHistoric").ToString)
    Dim vAddNew As Boolean = True
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    If vContactNumber > 0 Then
      Dim vDataChanged As Boolean = False
      Dim vRemoved As Boolean = False
      vAddNew = False
      If GetHiddenText("OldSuppression").Length = 0 Then
        'Add New Record
        vAddNew = True
      ElseIf FindControlByName(Me, "Suppression") Is Nothing OrElse GetCheckBoxChecked("Suppression") = False Then
        'The record is removed
        vDataChanged = True
        vRemoved = True
      ElseIf Not (ValueChanged("OldNotes", "Notes") _
          OrElse ValueChanged("OldValidFrom", "ValidFrom") OrElse ValueChanged("OldValidTo", "ValidTo")) Then
        'Nothing is changed. Don't add/delete anything
      Else
        'Data is changed so just update the record
        vDataChanged = True
      End If
      If vDataChanged Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vContactNumber
        AddUserParameters(vList)
        If vSetHistoric Then 'set the old record to historic and then create a new one
          vList("Suppression") = DefaultParameters("Suppression")
          vList("ValidFrom") = GetHiddenText("OldValidFrom")
          vList("ValidTo") = TodaysDate() 'Set the current as historic
          vList("AmendedOn") = TodaysDate()
          AddOptionalTextBoxValue(vList, "Notes")
          'check if theres already a historic suppression like this 
          Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactSuppressions, vList)
          If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
            'record returned so delete the current record
            vList("ValidTo") = GetHiddenText("OldValidTo")
            DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctSuppression, vList)
          Else
            vList("OldSuppression") = DefaultParameters("Suppression")
            vList("OldValidFrom") = GetHiddenText("OldValidFrom")
            vList("OldValidTo") = GetHiddenText("OldValidTo")
            vList("CarePortal") = "Y"
            DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctSuppression, vList)
          End If
          If Not vRemoved Then ' not unchecked suppression
            'then add new suppression for the new values.
            vAddNew = True
          End If
        Else 'delete the old record
          Try
            vList("Suppression") = DefaultParameters("Suppression")
            vList("ValidFrom") = GetHiddenText("OldValidFrom")
            vList("ValidTo") = GetHiddenText("OldValidTo")
            DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctSuppression, vList)
            If Not vRemoved Then
              vAddNew = True
            End If
          Catch vEx As CareException
            Debug.Print(vEx.Message)
          End Try
        End If
      End If
    Else
      vContactNumber = IntegerValue(pList("ContactNumber").ToString)
    End If
    If vAddNew AndAlso GetCheckBoxChecked("Suppression") Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = vContactNumber
      AddUserParameters(vList)
      vList("Suppression") = DefaultParameters("Suppression")
      AddOptionalTextBoxValue(vList, "Notes")
      vList("ValidFrom") = SetDate(DateType.ValidFrom) ' Today.ToShortDateString
      vList("ValidTo") = SetDate(DateType.ValidTo) 'Today.AddYears(100).ToShortDateString
      vList("CarePortal") = "Y"
      If DefaultParameters.ContainsKey("Source") AndAlso DefaultParameters("Source").ToString.Length > 0 Then
        vList("Source") = DefaultParameters("Source")
      End If
      DataHelper.AddSuppresion(vList)
    End If
  End Sub

  Public Overrides Sub ClearControls()
    MyBase.ClearControls()
    SetDefaultDates()
  End Sub
End Class