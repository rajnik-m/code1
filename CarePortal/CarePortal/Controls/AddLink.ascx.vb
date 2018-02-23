Public Partial Class AddLink
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    mvHandlesLinks = True
    mvUsesHiddenContactNumber = True
    mvHideHistoricLinks = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddLink, tblDataEntry, "", "")
      AddHiddenField("OldRelationship")  'To check if there is an existing record
      AddHiddenField("OldValidFrom")
      AddHiddenField("OldValidTo")
      AddHiddenField("OldNotes")
      AddHiddenField("OldRelationshipStatus")
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ClearControls()
    MyBase.ClearControls()
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vAddLink As Boolean = True
    Dim vContactNumber As Integer = GetHiddenContactNumber()
    Dim vContactNumber2 As Integer = IntegerValue(pList("ContactNumber").ToString)
    Dim vList As ParameterList
    Dim vUseCurrentDate As Boolean = BooleanValue(DefaultParameters("UseCurrentDate").ToString)
    Dim vSetHistoric As Boolean = BooleanValue(If(IsNothing(DefaultParameters("SetHistoric")), "N", DefaultParameters("SetHistoric").ToString))
    Dim vDataChanged As Boolean

    If vContactNumber > 0 Then
      'This is an existing record
      If Not (ValueChanged("OldNotes", "Notes") _
          OrElse ValueChanged("OldValidFrom", "ValidFrom") OrElse ValueChanged("OldValidTo", "ValidTo") OrElse ValueChanged("OldRelationshipStatus", "RelationshipStatus")) AndAlso _
        vContactNumber = IntegerValue(GetDropDownValue("ContactNumber1")) Then
        'Nothing is changed. Don't add/delete anything
        vAddLink = False
      Else
        vDataChanged = True
        If vSetHistoric AndAlso IsVisible() Then
          'set old record as historic with validto as todays date
          'but first see if there is a similar record with ValidTo and AmendedOn date set as current date
          vList = New ParameterList(HttpContext.Current)
          vList("ContactNumber") = vContactNumber
          vList("ContactNumber2") = vContactNumber2
          AddUserParameters(vList)
          vList("Relationship") = GetHiddenText("OldRelationship")
          vList("ValidTo") = TodaysDate() '
          vList("AmendedOn") = TodaysDate() '
          If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
            vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
          End If

          Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, vList)
          If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
            'a record is returned therefore do not add a new record, just delete the current record
            'vList("ValidFrom") = GetHiddenText("OldValidFrom")
            vList("ValidTo") = GetHiddenText("OldValidTo")
            vList.Remove("AmendedOn")
            If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
              vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
            End If
            Dim vReturnList As New ParameterList(HttpContext.Current)
            vReturnList = DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
            If vContactNumber = GetHiddenContactNumber() Then SetHiddenText("HiddenContactNumber", "0")

            'Delete the complimentary record
            If vReturnList.ContainsKey("ComplimentaryRelationship") Then
              vList("Relationship") = vReturnList.Item("ComplimentaryRelationship")
              vList("ContactNumber") = vContactNumber2
              vList("ContactNumber2") = vContactNumber
              DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
            End If
            'now add the new link
            vAddLink = True
          Else
            If IsVisible() Then
              'no existing historic record so update the old record to be historic
              vList = New ParameterList(HttpContext.Current)
              vList("ContactNumber2") = vContactNumber2
              If Convert.ToString(vList("UserLogname")).Length = 0 Then
                vList("UserID") = vContactNumber
              End If
              vList("ContactNumber") = vContactNumber
              vList("OldRelationship") = GetHiddenText("OldRelationship")
              vList("OldValidFrom") = GetHiddenText("OldValidFrom")
              vList("OldValidTo") = GetHiddenText("OldValidTo")
              vList("ValidTo") = TodaysDate()
              vList("Relationship") = vList("OldRelationship")
              vList("CarePortal") = "Y"
              AddOptionalTextBoxValue(vList, "ValidFrom", True)
              AddOptionalTextBoxValue(vList, "Notes")
              vList("CarePortal") = "Y" ' to avoid the merge mechanism in Update method
              If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
                vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
              End If
              Dim vReturnList As New ParameterList(HttpContext.Current)
              vReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList) 'just update valid_to to Todays Date
              'update the valid to date for the complimentary record too
              If vReturnList.ContainsKey("ComplimentaryRelationship") Then
                vList("Relationship") = vReturnList.Item("ComplimentaryRelationship")
                vList("OldRelationship") = vReturnList.Item("ComplimentaryRelationship")
                vList("ContactNumber2") = vContactNumber
                vList("ContactNumber") = vContactNumber2
                DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
              End If
              ''now add a new record 
              If GetHiddenContactNumber() = IntegerValue(GetDropDownValue("ContactNumber1")) Then
                vAddLink = False
              Else
                vAddLink = True
              End If
            End If
          End If
        Else
          If IsVisible() Then
            'Set historic not set so delete existing and then add a new
            vList = New ParameterList(HttpContext.Current)
            vList("ContactNumber") = vContactNumber
            AddUserParameters(vList)
            vList("ContactNumber2") = vContactNumber2
            vList("Relationship") = GetHiddenText("OldRelationship")
            vList("ValidFrom") = GetHiddenText("OldValidFrom")
            vList("ValidTo") = GetHiddenText("OldValidTo")
            If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
              vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
            End If
            Dim vReturnList As New ParameterList(HttpContext.Current)
            'also need to remove the complimentary relationship
            vReturnList = DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
            If vReturnList.ContainsKey("ComplimentaryRelationship") Then
              vList("Relationship") = vReturnList.Item("ComplimentaryRelationship")
              vList("ContactNumber") = vContactNumber2
              vList("ContactNumber2") = vContactNumber
              vReturnList = DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctLink, vList)
            End If
            'now add a new record 
            vAddLink = True
          End If

        End If
      End If
    End If
    If vAddLink AndAlso IsVisible() Then
      'new link therefore add a new link
      vContactNumber = IntegerValue(GetDropDownValue("ContactNumber1"))
      If vContactNumber > 0 Then
        vList = New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vContactNumber
        vList("ContactNumber2") = vContactNumber2
        If GetTextBoxText("ValidFrom").Length > 0 Then
          vList("ValidFrom") = GetTextBoxText("ValidFrom")
        Else
          vList("ValidFrom") = SetDate(DateType.ValidFrom) ' Today.ToShortDateString
        End If
        If GetTextBoxText("ValidTo").Length > 0 Then
          vList("ValidTo") = GetTextBoxText("ValidTo")
        Else
          vList("ValidTo") = ""
        End If
        If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
          vList("RelationshipStatus") = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList).SelectedValue
        End If
        AddOptionalTextBoxValue(vList, "Notes")
        vList("CarePortal") = "Y"
        If vUseCurrentDate AndAlso Not vList.Contains("ValidFrom") Then vList("ValidFrom") = TodaysDate()
        vList("Relationship") = DefaultParameters("Relationship").ToString
        Dim vReturnList As New ParameterList(HttpContext.Current)
        vReturnList = DataHelper.AddLink(vList)
        If vReturnList.ContainsKey("ComplimentaryRelationship") Then
          vList("Relationship") = vReturnList.Item("ComplimentaryRelationship")
          vList("ContactNumber") = vContactNumber2
          vList("ContactNumber2") = vContactNumber
          DataHelper.AddLink(vList)
        End If
      End If
    End If
  End Sub

  Public Overrides Sub ProcessLinkSelection(ByVal pRow As DataRow)
    SetHiddenText("OldRelationship", pRow("RelationshipCode").ToString)
    SetHiddenText("OldValidFrom", pRow("ValidFrom").ToString)
    SetHiddenText("OldValidTo", pRow("ValidTo").ToString)
    SetHiddenText("OldNotes", pRow("Notes").ToString)
    SetHiddenText("OldRelationshipStatus", pRow("RelationshipStatus").ToString)

    If Me.FindControl("ContactNumber1") IsNot Nothing AndAlso Me.FindControl("ContactNumber1").Visible Then
      Dim vContactDropDownList As DropDownList = DirectCast(Me.FindControl("ContactNumber1"), DropDownList)
      Dim vFoundItem As Boolean = False
      For Each vItem As ListItem In vContactDropDownList.Items
        If vItem.Value.ToString = pRow("ContactNumber").ToString Then
          vFoundItem = True
          Exit For
        End If
      Next
      'if no item is found, then it must be historic, therefore add it to the drop down
      If vFoundItem = False Then
        Dim vContactMainInfo As New ParameterList(HttpContext.Current)
        vContactMainInfo("ContactNumber") = pRow("ContactNumber").ToString
        Dim vContactInfo As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactMainInfo)
        Dim vContactInfoTable As DataTable = GetDataTable(vContactInfo)
        Dim vContactSalutation As String = String.Empty
        Dim vContactName As String = String.Empty
        If vContactInfoTable IsNot Nothing Then
          vContactName = vContactInfoTable.Rows(0)("LabelName").ToString
        End If
        vContactDropDownList.Items.Add(New ListItem(vContactName, pRow("ContactNumber").ToString))
      End If
    End If

    SetDropDownText("ContactNumber1", pRow("ContactNumber").ToString)
    SetTextBoxText("ValidFrom", pRow("ValidFrom").ToString)
    SetTextBoxText("ValidTo", pRow("ValidTo").ToString)
    SetTextBoxText("Notes", pRow("Notes").ToString)
    If Me.FindControl("RelationshipStatus") IsNot Nothing AndAlso Me.FindControl("RelationshipStatus").Visible Then
      Dim vDropDownList As DropDownList = DirectCast(Me.FindControl("RelationshipStatus"), DropDownList)
      If pRow("RelationshipCode").ToString.Trim.Length > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("Relationship") = pRow("RelationshipCode").ToString
        vDropDownList.DataTextField = "RelationshipStatusDesc"
        vDropDownList.DataValueField = "RelationshipStatus"
        DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtRelationshipStatuses, vDropDownList, True, vList, "Relationship Is Null OR Relationship = '" & pRow("RelationshipCode").ToString & "'")
        SetDropDownText("RelationshipStatus", pRow("RelationshipStatus").ToString)
      Else
        'ToDo clear values
      End If
    End If
  End Sub

  ''' <summary>
  ''' Returns false if any of the controls is set to visible = true else
  ''' returns false
  ''' </summary>
  ''' <returns>Returns True if the controls are visible else False </returns>
  ''' <remarks></remarks>
  Private Function IsVisible() As Boolean
    If tblDataEntry.Controls.Count > 0 Then
      Return True
    Else
      Return False
    End If
  End Function
End Class