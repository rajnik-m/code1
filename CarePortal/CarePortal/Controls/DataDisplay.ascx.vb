Partial Public Class DataDisplay
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If HTML.Length > 0 Then
        Dim vContactNumber As Integer = UserContactNumber()
        If vContactNumber > 0 Then
          If HTMLContains("ContactName") OrElse HTMLContains("Forenames") OrElse HTMLContains("Surname") OrElse _
             HTMLContains("LabelName") OrElse HTMLContains("DateOfBirth") OrElse HTMLContains("PreferredForename") OrElse _
             HTMLContains("Address_D") OrElse HTMLContains("Address_C") OrElse HTMLContains("Address_O") Then
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber))
            If vRow IsNot Nothing Then
              HTMLMerge("ContactName", vRow("ContactName").ToString)
              HTMLMerge("Forenames", vRow("Forenames").ToString)
              HTMLMerge("Surname", vRow("Surname").ToString)
              HTMLMerge("LabelName", vRow("LabelName").ToString)
              HTMLMerge("DateOfBirth", vRow("DateOfBirth").ToString)
              HTMLMerge("PreferredForename", vRow("PreferredForename").ToString)
              HTMLMerge("Address_D", vRow("AddressMultiLine").ToString.Replace(vbCrLf, "<BR/>"))
              If vRow("AddressType").ToString = "C" Then
                HTMLMerge("Address_C", vRow("AddressMultiLine").ToString.Replace(vbCrLf, "<BR/>"))  'Display C type of address
              Else
                HTMLMerge("Address_O", vRow("AddressMultiLine").ToString.Replace(vbCrLf, "<BR/>"))  'Display O type of address
              End If
            End If
          End If
          If HTMLContains("Address_C") OrElse HTMLContains("Address_O") OrElse _
             HTMLContains("Address_C_", False) OrElse HTMLContains("Address_O_", False) OrElse HTMLContains("Address_B_", False) Then
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesWithUsages, vContactNumber)
            Dim vRows() As DataRow
            If vTable IsNot Nothing Then
              If HTMLContains("Address_C") Then 'If the default address of the contact is of type C, this field is already populated
                vRows = vTable.Select("AddressType = 'C'")
                HTMLMerge("Address_C", If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
              End If
              If HTMLContains("Address_O") Then 'If the default address of the contact is of type O, this field is already populated
                vRows = vTable.Select("AddressType = 'O'")
                HTMLMerge("Address_O", If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
              End If
            Else
              HTMLMerge("Address_C", "")
              HTMLMerge("Address_O", "")
            End If
            Dim vUsageCode As String = ""
            While HTMLContains("Address_C_", False) 'Display C type of address with a given Address Usage Code
              vUsageCode = HTMLGetMailMergeCode("Address_C_")
              If vTable IsNot Nothing Then
                vRows = vTable.Select("AddressType = 'C' AND AddressUsage = '" & vUsageCode & "'")
                HTMLMerge("Address_C_" & vUsageCode, If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
              Else
                HTMLMerge("Address_C_" & vUsageCode, "")
              End If
            End While
            While HTMLContains("Address_O_", False) 'Display O type of address with a given Address Usage Code
              vUsageCode = HTMLGetMailMergeCode("Address_O_")
              If vTable IsNot Nothing Then
                vRows = vTable.Select("AddressType = 'O' AND AddressUsage = '" & vUsageCode & "'")
                HTMLMerge("Address_O_" & vUsageCode, If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
              Else
                HTMLMerge("Address_O_" & vUsageCode, "")
              End If
            End While
            While HTMLContains("Address_B_", False) 'Display any address with a given Address Usage Code
              vUsageCode = HTMLGetMailMergeCode("Address_B_")
              If vTable IsNot Nothing Then
                vRows = vTable.Select("AddressUsage = '" & vUsageCode & "'")
                HTMLMerge("Address_B_" & vUsageCode, If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
              Else
                HTMLMerge("Address_B_" & vUsageCode, "")
              End If
            End While
          End If
          If HTMLContains("Communication_P") OrElse _
             HTMLContains("Communication_D", True) OrElse _
             HTMLContains("Communication_DC_", False) OrElse _
             HTMLContains("Communication_DUC_", False) Then
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbersWithUsages, vContactNumber)
            Dim vRows() As DataRow
            If vTable IsNot Nothing Then
              vRows = vTable.Select("Default LIKE 'Y*'")  'Display default phone number
              HTMLMerge("Communication_D", If(vRows.Length > 0, vRows(0)("PhoneNumber").ToString, ""))

              vRows = vTable.Select("PreferredMethod LIKE 'Y*'")  'Display preferred phone number
              HTMLMerge("Communication_P", If(vRows.Length > 0, vRows(0)("PhoneNumber").ToString, ""))
            Else
              HTMLMerge("Communication_D", "")
              HTMLMerge("Communication_P", "")
            End If
            Dim vDeviceCode As String = ""
            While HTMLContains("Communication_DC_", False)  'Display number for given Device Code
              vDeviceCode = HTMLGetMailMergeCode("Communication_DC_")
              If vTable IsNot Nothing Then
                vRows = vTable.Select("DeviceCode = '" & vDeviceCode & "' AND Default LIKE 'Y*'")
                If vRows.Length = 0 Then vRows = vTable.Select("DeviceCode = '" & vDeviceCode & "'")
                HTMLMerge("Communication_DC_" & vDeviceCode, If(vRows.Length > 0, GetMostRecentRow(vRows)("PhoneNumber").ToString, ""))
              Else
                HTMLMerge("Communication_DC_" & vDeviceCode, "")
              End If
            End While
            Dim vUsageCode As String = ""
            While HTMLContains("Communication_DUC_", False) 'Display number for given Device and Communication Usage Codes
              vDeviceCode = HTMLGetMailMergeCode("Communication_DUC_")
              vUsageCode = vDeviceCode.Split("_"c)(1)
              vDeviceCode = vDeviceCode.Split("_"c)(0)
              If vTable IsNot Nothing Then
                vRows = vTable.Select("DeviceCode = '" & vDeviceCode & "' AND CommunicationUsage = '" & vUsageCode & "' AND Default LIKE 'Y*'")
                If vRows.Length = 0 Then vRows = vTable.Select("DeviceCode = '" & vDeviceCode & "' AND CommunicationUsage = '" & vUsageCode & "'")
                HTMLMerge("Communication_DUC_" & vDeviceCode & "_" & vUsageCode, If(vRows.Length > 0, GetMostRecentRow(vRows)("PhoneNumber").ToString, ""))
              Else
                HTMLMerge("Communication_DUC_" & vDeviceCode & "_" & vUsageCode, "")
              End If
            End While
          End If
          If HTMLContains("Activity_C_", False) OrElse HTMLContains("Activity_D_", False) OrElse HTMLContains("Activity_N_", False) OrElse HTMLContains("Activity_Q_", False) Then
            Dim vList As New ParameterList(HttpContext.Current)
            vList("Current") = "Y"
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vContactNumber, vList)
            Dim vRows() As DataRow
            Dim vActivityCode As String = ""
            Dim vValue As New StringBuilder
            While HTMLContains("Activity_C_", False)  'Display all Activity Value Codes for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("Activity_C_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("ActivityValueCode").ToString)
                  Next
                End If
              End If
              HTMLMerge("Activity_C_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("Activity_D_", False)  'Display all Activity Value Descriptions for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("Activity_D_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("ActivityValueDesc").ToString)
                  Next
                End If
              End If
              HTMLMerge("Activity_D_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("Activity_N_", False)  'Display all Activity Value Notes for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("Activity_N_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("Notes").ToString)
                  Next
                End If
              End If
              HTMLMerge("Activity_N_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("Activity_Q_", False)  'Display all Activity Value Quantities for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("Activity_Q_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("Quantity").ToString)
                  Next
                End If
              End If
              HTMLMerge("Activity_Q_" & vActivityCode, vValue.ToString)
            End While
          End If

          If HTMLContains("MemberNumber") Then
            Dim vList As New ParameterList(HttpContext.Current)
            vList("CancellationReason") = ""
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactMemberships, vContactNumber, vList)
            If vTable IsNot Nothing Then
              HTMLMerge("MemberNumber", vTable.Rows(0)("MemberNumber").ToString)
              HTMLMerge("MembershipTypeDesc", vTable.Rows(0)("MembershipTypeDesc").ToString)
              HTMLMerge("RenewalDate", vTable.Rows(0)("RenewalDate").ToString)
              HTMLMerge("Joined", vTable.Rows(0)("Joined").ToString)
            End If
          End If

        End If

        If Request.QueryString("EN") IsNot Nothing AndAlso Request.QueryString("EN").Length > 0 Then
          If HTMLContains("EventDesc") Then
            Dim vEventNumber As Integer = IntegerValue(Request.QueryString("EN"))
            If vEventNumber > 0 Then
              Dim vList As New ParameterList(HttpContext.Current)
              vList("EventNumber") = vEventNumber.ToString
              Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetEventDataTable(CarePortal.CareNetServices.XMLEventDataSelectionTypes.xedtEventInformation, vList))
              If vRow IsNot Nothing Then
                HTMLMerge("EventDesc", vRow("EventDesc").ToString)
              End If
            End If
          End If
          If HTML.Contains("<<EventNumber>>") Then
            Dim vEventNumber As Integer = IntegerValue(Request.QueryString("EN"))
            If vEventNumber > 0 Then
              HTML = HTML.Replace("<<EventNumber>>", vEventNumber.ToString)
            End If
          End If
        End If
        If Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("PR").Length > 0 Then
          If HTMLContains("ProductDesc") Then
            Dim vProduct As String = Request.QueryString("PR")
            Dim vList As New ParameterList(HttpContext.Current)
            vList("Product") = vProduct
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtProducts, vList))
            If vRow IsNot Nothing Then
              HTMLMerge("ProductDesc", vRow("ProductDesc").ToString)
            End If
          End If
        End If
        If Session("SelectedOrganisationNumber") IsNot Nothing Then
          vContactNumber = IntegerValue(Session("SelectedOrganisationNumber").ToString)
          If HTMLContains("SelectedOrganisationName") Then
            Dim vList As New ParameterList(HttpContext.Current)
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber))
            If vRow IsNot Nothing Then
              HTMLMerge("SelectedOrganisationName", vRow("ContactName").ToString)
            End If
          End If
          If HTMLContains("SelectedOrganisationAddress") Then
            Dim vList As New ParameterList(HttpContext.Current)
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesWithUsages, vContactNumber)
            If vTable IsNot Nothing Then
              Dim vRows() As DataRow = vTable.Select("Default = 'Yes'")
              HTMLMerge("SelectedOrganisationAddress", If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
            Else
              HTMLMerge("SelectedOrganisationAddress", "")
            End If
          End If
          If HTMLContains("SelectedOrganisationActivity_C_", False) OrElse HTMLContains("SelectedOrganisationActivity_D_", False) OrElse HTMLContains("SelectedOrganisationActivity_N_", False) OrElse HTMLContains("SelectedOrganisationActivity_Q_", False) Then
            Dim vList As New ParameterList(HttpContext.Current)
            vList("Current") = "Y"
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, vContactNumber, vList)
            Dim vRows() As DataRow
            Dim vActivityCode As String = ""
            Dim vValue As New StringBuilder
            While HTMLContains("SelectedOrganisationActivity_C_", False)  'Display all Selected Organisation Activity Value Codes for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("SelectedOrganisationActivity_C_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("ActivityValueCode").ToString)
                  Next
                End If
              End If
              HTMLMerge("SelectedOrganisationActivity_C_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("SelectedOrganisationActivity_D_", False)  'Display all Selected Organisation Activity Value Descriptions for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("SelectedOrganisationActivity_D_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("ActivityValueDesc").ToString)
                  Next
                End If
              End If
              HTMLMerge("SelectedOrganisationActivity_D_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("SelectedOrganisationActivity_N_", False)  'Display all Selected Organisation Activity Notes for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("SelectedOrganisationActivity_N_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("Notes").ToString)
                  Next
                End If
              End If
              HTMLMerge("SelectedOrganisationActivity_N_" & vActivityCode, vValue.ToString)
            End While
            vValue = New StringBuilder
            While HTMLContains("SelectedOrganisationActivity_Q_", False)  'Display all Selected Organisation Activity Quantities for given Activity Code
              vActivityCode = HTMLGetMailMergeCode("SelectedOrganisationActivity_Q_")
              vValue.Clear()
              If vTable IsNot Nothing Then
                vRows = vTable.Select("ActivityCode = '" & vActivityCode & "'")
                If vRows.Length > 0 Then
                  For Each vRow As DataRow In vRows
                    If vValue.Length > 0 Then vValue.Append("<BR/>")
                    vValue.Append(vRow("Quantity").ToString)
                  Next
                End If
              End If
              HTMLMerge("SelectedOrganisationActivity_Q_" & vActivityCode, vValue.ToString)
            End While
          End If
        End If
        If Session("SelectedContactNumber") IsNot Nothing Then
          If HTMLContains("SelectedContactName") Then
            Dim vList As New ParameterList(HttpContext.Current)
            vContactNumber = IntegerValue(Session("SelectedContactNumber").ToString)
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber))
            If vRow IsNot Nothing Then
              HTMLMerge("SelectedContactName", vRow("ContactName").ToString)
            End If
          End If
          If HTMLContains("SelectedContactAddress") Then
            Dim vList As New ParameterList(HttpContext.Current)
            vContactNumber = IntegerValue(Session("SelectedContactNumber").ToString)
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesWithUsages, vContactNumber)
            If vTable IsNot Nothing Then
              Dim vRows() As DataRow = vTable.Select("Default = 'Yes'")
              HTMLMerge("SelectedContactAddress", If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
            Else
              HTMLMerge("SelectedContactAddress", "")
            End If
          End If
        End If
        If Session("PayerContactNumber") IsNot Nothing Then
          If HTMLContains("PayerContactName") Then
            Dim vList As New ParameterList(HttpContext.Current)
            vContactNumber = IntegerValue(Session("PayerContactNumber").ToString)
            Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vContactNumber))
            If vRow IsNot Nothing Then
              HTMLMerge("PayerContactName", vRow("ContactName").ToString)
            End If
          End If
          If HTMLContains("PayerContactAddress") Then
            Dim vList As New ParameterList(HttpContext.Current)
            vContactNumber = IntegerValue(Session("PayerContactNumber").ToString)
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesWithUsages, vContactNumber)
            If vTable IsNot Nothing Then
              Dim vRows() As DataRow = vTable.Select("AddressNumber = " & Session("PayerAddressNumber").ToString)
              HTMLMerge("PayerContactAddress", If(vRows.Length > 0, vRows(0)("AddressMultiLine").ToString, "").Replace(vbCrLf, "<BR/>"))
            Else
              HTMLMerge("PayerContactAddress", "")
            End If
          End If
        End If


        'Direct Debit Changes
        If Session("DirectDebitNumber") IsNot Nothing AndAlso CInt(Session("DirectDebitNumber")) > 0 Then
          DisplayDirectDebitDetails()
        End If

        If Request.QueryString("MT") IsNot Nothing AndAlso Request.QueryString("MT").Length > 0 Then
          If HTMLContains("SelectedMembershipType") Then
            HTMLMerge("SelectedMembershipType", Request.QueryString("MT").ToString)
          End If
        End If
        If Request.QueryString("SD") IsNot Nothing AndAlso Request.QueryString("SD").Length > 0 Then
          If HTMLContains("SelectedStartDate") Then
            HTMLMerge("SelectedStartDate", Request.QueryString("SD").ToString)
          End If
        End If

        Dim vHTMLRow As New HtmlTableRow
        Dim vHTMLCell As New HtmlTableCell
        vHTMLCell.InnerHtml = HTML
        vHTMLRow.Cells.Add(vHTMLCell)
        tblContent.Rows.Add(vHTMLRow)
        tblContent.Attributes("Class") = GetClass()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Function GetMostRecentRow(ByVal pRows() As DataRow) As DataRow
    Dim vRow As DataRow = Nothing
    If pRows.Length > 0 Then
      If pRows(0).Item("PreferredMethod").ToString.Length > 0 Then Return pRows(0)
      If pRows(0).Item("DeviceDefault").ToString.Length > 0 Then Return pRows(0)
    End If
    For Each vCheckRow As DataRow In pRows
      If vRow Is Nothing Then
        vRow = vCheckRow
      Else
        If IntegerValue(vCheckRow("CommunicationNumber").ToString) > IntegerValue(vRow("CommunicationNumber").ToString) Then
          vRow = vCheckRow
        End If
      End If
    Next
    Return vRow
  End Function

  Private Function HTMLContains(ByVal pMailMergeField As String) As Boolean
    Return HTMLContains(pMailMergeField, True)
  End Function
  Private Function HTMLContains(ByVal pMailMergeField As String, ByVal pExactMatch As Boolean) As Boolean
    Dim vMailMergeSuffix As String = If(pExactMatch, "&gt;&gt;", "")
    Return HTML.Contains("&lt;&lt;" & pMailMergeField & vMailMergeSuffix)
  End Function

  Private Sub HTMLMerge(ByVal pMailMergeField As String, ByVal pValue As String)
    If pValue.Length = 0 Then pValue = "&nbsp;"
    HTML = HTML.Replace("&lt;&lt;" & pMailMergeField & "&gt;&gt;", Server.HtmlEncode(pValue))
  End Sub

  Private Function HTMLGetMailMergeCode(ByVal pMailMergeField As String) As String
    Dim vSearchField As String = "&lt;&lt;" & pMailMergeField
    Dim vTemp As String = HTML.Substring(HTML.IndexOf(vSearchField) + vSearchField.Length)
    Return vTemp.Substring(0, vTemp.IndexOf("&gt;&gt;"))
  End Function
  ''' <summary>
  ''' Gets all the details required to display Direct debit details from the database 
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DisplayDirectDebitDetails()
    Dim vContactNumber As Integer = UserContactNumber()
    If vContactNumber > 0 Then
      Dim vDirectDebitNumber As Integer = IntegerValue(Session("DirectDebitNumber").ToString)
      Dim vList As New ParameterList(HttpContext.Current)
      vList.Add("ContactNumber", vContactNumber)
      vList.Add("DirectDebitNumber", vDirectDebitNumber)
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CarePortal.CareNetServices.XMLContactDataSelectionTypes.xcdtContactDirectDebits, vList))
      If vRow IsNot Nothing Then
        If HTMLContains("DirectDebitSortCode") Then HTMLMerge("DirectDebitSortCode", vRow("SortCode").ToString)
        If HTMLContains("DirectDebitAccountNumber") Then HTMLMerge("DirectDebitAccountNumber", vRow("AccountNumber").ToString)
        If HTMLContains("DirectDebitMandateNumber") Then HTMLMerge("DirectDebitMandateNumber", vRow("Reference").ToString)
        If HTMLContains("DirectDebitAccountHolderName") Then HTMLMerge("DirectDebitAccountHolderName", vRow("AccountName").ToString)

        Dim vBankParams As New ParameterList(HttpContext.Current)
        vBankParams.Add("SortCode", vRow("SortCode").ToString.Replace("-", ""))
        Dim vBankDetailRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtBanks, vBankParams))
        If vBankDetailRow IsNot Nothing Then
          If HTMLContains("DirectDebitBankName") Then HTMLMerge("DirectDebitBankName", vBankDetailRow("Bank").ToString)

          If HTMLContains("DirectDebitBankAddress") Then
            HTMLMerge("DirectDebitBankAddress", GetBankAddress(vBankDetailRow, False))
          End If

          If HTMLContains("DirectDebitBankAddressMultiLine") Then
            HTMLMerge("DirectDebitBankAddressMultiLine", GetBankAddress(vBankDetailRow, True))
          End If
        End If

        'display payment plan details from direct debit
        DisplayPaymentPlanDetails()
       
      End If
    End If
  End Sub
  ''' <summary>
  ''' This method will format the bank address depending on the pMultiLine
  ''' </summary>
  ''' <param name="pDataRow">Data row with bank details from bank table</param>
  ''' <param name="pMultiLine">Flag to decide on the Bank Address format. If true will put all details on new line else seprates with ,</param>
  ''' <returns>Formatted bank address</returns>
  ''' <remarks></remarks>
  Private Function GetBankAddress(ByVal pDataRow As DataRow, ByVal pMultiLine As Boolean) As String
    Dim vBankAddress As New StringBuilder()
    vBankAddress.Append(pDataRow("Address").ToString)
    If vBankAddress.Length > 0 Then
      If pMultiLine Then vBankAddress.Append("<BR/>") Else vBankAddress.Append(",")
    End If
    vBankAddress.Append(pDataRow("Town").ToString)
    If vBankAddress.Length > 0 Then
      If pMultiLine Then vBankAddress.Append("<BR/>") Else vBankAddress.Append(",")
    End If
    vBankAddress.Append(pDataRow("County").ToString)
    Return vBankAddress.ToString
  End Function
  ''' <summary>
  ''' Prints the Payment Plan details for Direct debit payment plans
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DisplayPaymentPlanDetails()
    If Session("DirectDebitPaymentFrequency") IsNot Nothing AndAlso HTMLContains("DirectDebitPaymentFrequency") Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList.Add("PaymentFrequency", Session("DirectDebitPaymentFrequency").ToString)
      Dim vDataRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtPaymentFrequencies, vList))
      If vDataRow IsNot Nothing Then HTMLMerge("DirectDebitPaymentFrequency", vDataRow("PaymentFrequencyDesc").ToString)
    End If
    If Session("DirectDebitFrequencyAmount") IsNot Nothing AndAlso HTMLContains("DirectDebitFrequencyAmount") Then HTMLMerge("DirectDebitFrequencyAmount", Session("DirectDebitFrequencyAmount").ToString)
    If Session("DirectDebitClaimDate") IsNot Nothing AndAlso HTMLContains("DirectDebitNextClaimDate") Then HTMLMerge("DirectDebitNextClaimDate", Session("DirectDebitClaimDate").ToString)
  End Sub
End Class