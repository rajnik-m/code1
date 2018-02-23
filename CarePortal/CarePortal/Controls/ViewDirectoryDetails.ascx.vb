Public Class ViewDirectoryDetails
  Inherits CareWebControl
  Dim mvContactNumber As String
  Dim mvContactGroup As String
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vDirectoryName As String = String.Empty
      InitialiseControls(CareNetServices.WebControlTypes.wctViewDirectoryDetails, tblDataEntry)
      If InitialParameters.ContainsKey("DirectoryName") Then vDirectoryName = InitialParameters("DirectoryName").ToString
      If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
        mvContactNumber = Request.QueryString("CN")
      Else
        If InitialParameters.ContainsKey("ContactNumber") Then mvContactNumber = InitialParameters("ContactNumber").ToString
      End If
      If mvContactNumber IsNot Nothing Then
        Dim vList As New ParameterList(HttpContext.Current)
        Dim vResult As String
        Dim vDataTable As DataTable
        Dim vRow As DataRow

        vList("ContactNumber") = mvContactNumber
        vRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
        mvContactGroup = vRow.Item("GroupCode").ToString
        vList("ViewName") = vDirectoryName
        If mvContactGroup = "CON" Then
          vList("ContactType") = "C"
        ElseIf mvContactGroup = "ORG" Then
          vList("ContactType") = "O"
        End If
        'Search contact in specified directory
        vList("UserID") = UserContactNumber()
        vList("AddressNumber") = vRow.Item("AddressNumber").ToString
        vList("ViewDetails") = "Y"
        vResult = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebDirectoryEntries, vList)
        vDataTable = GetDataTable(vResult)
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
          Dim vTextBox As TextBox = TryCast(Me.FindControl("ContactNumber"), TextBox)
          If vTextBox IsNot Nothing Then
            Dim vLabel As Label = TryCast(Me.FindControl(vTextBox.ID & "_Desc"), Label)
            If vLabel IsNot Nothing Then vLabel.Visible = False
          End If
          SetTextBoxText("ContactNumber", mvContactNumber)
          SetContactAndOrgDetails()
          FindControlByName(Me, "WarningMessage").Visible = False
        Else
          SetLabelText("WarningMessage", String.Format("Contact not found in directory {0}", vDirectoryName))
          SetControlVisibleForDirectory(False)
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

  Private Sub SetControlVisibleForDirectory(ByVal pVisible As Boolean)
    'Set the textbox and label invisible
    If FindControlByName(Me, "ContactNumber") IsNot Nothing Then
      FindControlByName(Me, "MemberNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "MemberNumber").Visible = pVisible
      FindControlByName(Me, "ContactNumber").Visible = pVisible
      FindControlByName(Me, "ContactNumber").Parent.Parent.Parent.Parent.Visible = pVisible
      FindControlByName(Me, "Surname").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "Surname").Parent.Visible = pVisible
      FindControlByName(Me, "Address").Visible = pVisible
      FindControlByName(Me, "Address").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "DefaultContactNumber").Visible = pVisible
      FindControlByName(Me, "DefaultContactNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "Position").Visible = pVisible
      FindControlByName(Me, "Position").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "Name").Visible = pVisible
      FindControlByName(Me, "Name").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "DirectDial").Visible = pVisible
      FindControlByName(Me, "DirectDial").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "NiNumber").Visible = pVisible
      FindControlByName(Me, "NiNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "DOB").Visible = pVisible
      FindControlByName(Me, "DOB").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "MembershipType").Visible = pVisible
      FindControlByName(Me, "MembershipType").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "EmailAddress").Visible = pVisible
      FindControlByName(Me, "EmailAddress").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "WebAddress").Visible = pVisible
      FindControlByName(Me, "WebAddress").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "SwitchboardNumber").Visible = pVisible
      FindControlByName(Me, "SwitchboardNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "FaxNumber").Visible = pVisible
      FindControlByName(Me, "FaxNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "MobileNumber").Visible = pVisible
      FindControlByName(Me, "MobileNumber").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "DirectoryAddress").Visible = pVisible
      FindControlByName(Me, "DirectoryAddress").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "CommunicationUsage1").Visible = pVisible
      FindControlByName(Me, "CommunicationUsage1").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "CommunicationUsage2").Visible = pVisible
      FindControlByName(Me, "CommunicationUsage2").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "CommunicationUsage3").Visible = pVisible
      FindControlByName(Me, "CommunicationUsage3").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "CommunicationUsage4").Visible = pVisible
      FindControlByName(Me, "CommunicationUsage4").Parent.Parent.Visible = pVisible
    End If
    If FindControlByName(Me, "ActivityValue1") IsNot Nothing Then
      FindControlByName(Me, "ActivityValue1").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue1").Visible = pVisible
      FindControlByName(Me, "ActivityValue2").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue2").Visible = pVisible
      FindControlByName(Me, "ActivityValue3").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue3").Visible = pVisible
      FindControlByName(Me, "ActivityValue4").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue4").Visible = pVisible
      FindControlByName(Me, "ActivityValue5").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue5").Visible = pVisible
      FindControlByName(Me, "ActivityValue6").Parent.Parent.Visible = pVisible
      FindControlByName(Me, "ActivityValue6").Visible = pVisible
    End If
  End Sub

  Private Sub SetContactAndOrgDetails()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vRow As DataRow
    Dim vDataTable As DataTable
    vList("ContactNumber") = mvContactNumber
    vRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
    If vRow IsNot Nothing Then
      SetTextBoxText("Address", vRow("AddressLine").ToString)
      SetTextBoxText("NiNumber", vRow("NiNumber").ToString)
      SetTextBoxText("DOB", vRow("DateOfBirth").ToString)
      SetTextBoxText("Position", vRow("Position").ToString)
      SetTextBoxText("Name", vRow("OrganisationName").ToString)

      If mvContactGroup = "CON" Then
        SetTextBoxText("Surname", vRow("surname").ToString)
      ElseIf vRow("GroupCode").ToString = "ORG" Then
        SetTextBoxText("DefaultContactNumber", vRow("DefaultContactName").ToString)
        SetTextBoxText("Surname", vRow("OrganisationName").ToString)
      End If
    End If
    'Set the communications items
    vList("AddressNumber") = vRow("AddressNumber")
    vRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsInformation, vList))
    If vRow IsNot Nothing Then
      SetTextBoxText("DirectDial", vRow("DirectNumber").ToString)
      SetTextBoxText("MobileNumber", vRow("MobileNumber").ToString)
      SetTextBoxText("EmailAddress", vRow("EMailAddress").ToString)
      SetTextBoxText("WebAddress", vRow("WebAddress").ToString)
      SetTextBoxText("SwitchboardNumber", vRow("SwitchboardNumber").ToString)
      SetTextBoxText("FaxNumber", vRow("FaxNumber").ToString)
    End If

    If InitialParameters("AddressUsage") IsNot Nothing Then vList("AddressUsage") = InitialParameters("AddressUsage").ToString
    If InitialParameters("CommunicationUsage1") IsNot Nothing Then vList("CommunicationUsage1") = InitialParameters("CommunicationUsage1").ToString
    If InitialParameters("CommunicationUsage2") IsNot Nothing Then vList("CommunicationUsage2") = InitialParameters("CommunicationUsage2").ToString
    If InitialParameters("CommunicationUsage3") IsNot Nothing Then vList("CommunicationUsage3") = InitialParameters("CommunicationUsage3").ToString
    If InitialParameters("CommunicationUsage4") IsNot Nothing Then vList("CommunicationUsage4") = InitialParameters("CommunicationUsage4").ToString

    Dim vDirectoryData As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtDirectoryUsage, vList)

    SetDirectoryValue("AddressUsage", vDirectoryData, "DirectoryAddress")
    SetDirectoryValue("CommunicationUsage1", vDirectoryData, "CommunicationUsage1")
    SetDirectoryValue("CommunicationUsage2", vDirectoryData, "CommunicationUsage2")
    SetDirectoryValue("CommunicationUsage3", vDirectoryData, "CommunicationUsage3")
    SetDirectoryValue("CommunicationUsage4", vDirectoryData, "CommunicationUsage4")
    
    'Set the Membership Number and Membership Type.
    vDataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactMemberships, vList)
    If vDataTable IsNot Nothing Then
      For Each vDataRow As DataRow In vDataTable.Rows
        If vDataRow.Item("CancelledOn").ToString.Length = 0 Then
          SetTextBoxText("MemberNumber", vDataRow("MemberNumber").ToString)
          SetTextBoxText("MembershipType", vDataRow("MembershipTypeDesc").ToString)
          Exit For
        End If
      Next
    End If
    SetContactAndOrgControl()
    SetActivityValues()
  End Sub
  Private Sub SetContactAndOrgControl()
    If mvContactGroup = "CON" Then
      'For Contact, Default contact should be invisible
      If FindControlByName(Me, "DefaultContactNumber") IsNot Nothing Then
        FindControlByName(Me, "DefaultContactNumber").Parent.Parent.Visible = False
        FindControlByName(Me, "DefaultContactNumber").Visible = False
      End If
    ElseIf mvContactGroup = "ORG" Then
      'For organisation Ni Number, Date of birth,Position and organsiation name should be invisible.
      If FindControlByName(Me, "NiNumber") IsNot Nothing Then
        FindControlByName(Me, "NiNumber").Parent.Parent.Visible = False
        FindControlByName(Me, "NiNumber").Visible = False
        FindControlByName(Me, "DOB").Parent.Parent.Visible = False
        FindControlByName(Me, "DOB").Visible = False
        FindControlByName(Me, "Position").Parent.Parent.Visible = False
        FindControlByName(Me, "Position").Visible = False
        FindControlByName(Me, "Name").Parent.Parent.Visible = False
        FindControlByName(Me, "Name").Visible = False
      End If
    End If
  End Sub

  Private Sub SetActivityValues()
    Dim vDataTable As DataTable
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vActivityValue As New StringBuilder
    vList("ContactNumber") = mvContactNumber
    vDataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vList)
    If vDataTable IsNot Nothing Then
      'For Activity1
      SetActivityValue("Activity1", vDataTable, "ActivityValue1")
      'For Activity2
      SetActivityValue("Activity2", vDataTable, "ActivityValue2")
      'For Activity3
      SetActivityValue("Activity3", vDataTable, "ActivityValue3")
      'For Activity4
      SetActivityValue("Activity4", vDataTable, "ActivityValue4")
      'For Activity5
      SetActivityValue("Activity5", vDataTable, "ActivityValue5")
      'For Activity6
      SetActivityValue("Activity6", vDataTable, "ActivityValue6")
    End If
  End Sub

  Private Sub SetActivityValue(ByVal pParameterName As String, ByVal pDataTable As DataTable, ByVal pTextBox As String)
    Dim vActivityValue As New StringBuilder
    If InitialParameters.ContainsKey(pParameterName) AndAlso InitialParameters(pParameterName).ToString.Length > 0 Then
      For Each vRow As DataRow In pDataTable.Rows
        If vRow.Item("ActivityCode").ToString = InitialParameters(pParameterName).ToString Then
          If CDate(vRow("ValidFrom").ToString) <= Date.Today AndAlso CDate(vRow("ValidTo").ToString) >= Date.Today Then
            vActivityValue.Append(",")
            vActivityValue.Append(vRow("ActivityValueDesc").ToString)
          End If
        End If
      Next
      If vActivityValue.Length > 1 Then vActivityValue.Remove(0, 1)
      If vActivityValue.ToString.Length > 0 Then
        SetTextBoxText(pTextBox, vActivityValue.ToString)
      Else
        If FindControlByName(Me, pTextBox) IsNot Nothing Then
          FindControlByName(Me, pTextBox).Parent.Parent.Visible = False
          FindControlByName(Me, pTextBox).Visible = False
        End If
      End If
    Else
      If FindControlByName(Me, pTextBox) IsNot Nothing Then
        FindControlByName(Me, pTextBox).Parent.Parent.Visible = False
        FindControlByName(Me, pTextBox).Visible = False
      End If
    End If
  End Sub

  Private Sub SetDirectoryValue(ByVal pParameterName As String, ByVal pTable As DataTable, ByVal pTextBox As String)
    If pTable IsNot Nothing AndAlso InitialParameters.ContainsKey(pParameterName) AndAlso InitialParameters(pParameterName).ToString.Length > 0 Then
      Dim vEndText As String = ""
      Dim vParamValue As String = InitialParameters(pParameterName).ToString

      For Each vDr As DataRow In pTable.Rows
        If vDr("CommunicationUsage").ToString = vParamValue Then
          vEndText = vDr("Value").ToString
          Exit For
        ElseIf vDr("AddressUsage").ToString = vParamValue Then
          vEndText = vDr("Value").ToString
          Exit For
        End If
      Next

      If vEndText.Length > 0 Then
        SetTextBoxText(pTextBox, vEndText)
      ElseIf FindControlByName(Me, pTextBox) IsNot Nothing Then
        FindControlByName(Me, pTextBox).Parent.Parent.Visible = False
        FindControlByName(Me, pTextBox).Visible = False
      End If
    Else
      FindControlByName(Me, pTextBox).Parent.Parent.Visible = False
      FindControlByName(Me, pTextBox).Visible = False
    End If
  End Sub
End Class