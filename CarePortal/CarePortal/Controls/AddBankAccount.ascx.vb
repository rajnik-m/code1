Public Partial Class AddBankAccount
  Inherits CareWebControl
  Implements ICareChildWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      If Not IsPostBack Then SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If IsPostBack Then SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub SetDefaults()
    mvNeedsParent = True
    mvHandlesBankAccounts = True
    InitialiseControls(CareNetServices.WebControlTypes.wctAddBankAccount, tblDataEntry)
    AddHiddenField("BankDetailsNumber")
    AddHiddenField("OldAccountNumber")
    AddHiddenField("OldSortCode")
  End Sub

  Public Overrides Sub ProcessBankAccountSelection(ByVal pTable As DataTable)
    If pTable IsNot Nothing AndAlso pTable.Rows.Count > 0 Then
      Dim vRow As DataRow = pTable.Rows(0)
      SetHiddenText("BankDetailsNumber", vRow("BankDetailsNumber").ToString)
      SetHiddenText("OldAccountNumber", vRow("AccountNumber").ToString)
      SetHiddenText("OldSortCode", vRow("SortCode").ToString.Replace("-", ""))
      'Encrypt Account Number
      Dim vAccountNumber As String = vRow("AccountNumber").ToString
      Dim vEncryptedAN As New StringBuilder
      Dim vEncryptedDigits As Integer = 4
      If vAccountNumber.Length = 1 Then
        vEncryptedDigits = 0                                          'No Encryption required
      ElseIf vAccountNumber.Length <= 6 Then
        vEncryptedDigits = CInt(vAccountNumber.Length / 2)            'Display half digits
      Else
        vEncryptedDigits = vAccountNumber.Length - vEncryptedDigits   'Display the last 4 digits only
      End If
      For vIndex As Integer = 1 To vEncryptedDigits
        vEncryptedAN.Append("*")
      Next
      vEncryptedAN.Append(Substring(vAccountNumber, vEncryptedDigits, vAccountNumber.Length))
      SetTextBoxText("AccountNumber", vEncryptedAN.ToString)
      SetTextBoxText("AccountName", vRow("AccountName").ToString)
      SetDropDownText("Bank", vRow("BankName").ToString, True)
      SetDropDownText("BranchName", vRow("BranchName").ToString)
      'Encrypt Sort Code
      Dim vSortCode As String = vRow("SortCode").ToString
      Dim vEncryptedSC As New StringBuilder
      vEncryptedSC.Append(vSortCode.Substring(0, 2))
      For vIndex As Integer = 2 To vSortCode.Length - 1
        If vSortCode.Substring(vIndex, 1) = "-" Then
          vEncryptedSC.Append("-")
        Else
          vEncryptedSC.Append("*")
        End If
      Next
      SetTextBoxText("SortCode", vEncryptedSC.ToString)
      SetTextBoxText("BankPayerName", vRow("BankPayerName").ToString)
      SetTextBoxText("Notes", vRow("Notes").ToString)
    End If
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vAccountNo As String = GetTextBoxText("AccountNumber")
    Dim vSortCode As String = GetTextBoxText("SortCode")
    If vAccountNo.Length > 0 AndAlso vSortCode.Length > 0 AndAlso GetDropDownValue("Bank").Length > 0 AndAlso GetDropDownValue("BranchName").Length > 0 Then
      If vAccountNo.Contains("*") Then vAccountNo = GetHiddenText("OldAccountNumber")
      If vSortCode.Contains("*") Then vSortCode = GetHiddenText("OldSortCode")

      'Dim vList As New ParameterList(HttpContext.Current)
      Dim vBDN As String = GetHiddenText("BankDetailsNumber")
      If vBDN.Length > 0 Then pList("BankDetailsNumber") = vBDN
      AddOptionalTextBoxValue(pList, "AccountName")
      pList("BankPayerName") = GetTextBoxText("BankPayerName")
      pList("Notes") = GetTextBoxText("Notes")

      If vAccountNo = GetHiddenText("OldAccountNumber") AndAlso vSortCode = GetHiddenText("OldSortCode") Then
        'AccountNumber and SortCode are not changed- Just update the record with other details
        DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts, pList)
      ElseIf GetTextBoxText("AccountName").Length > 0 Then
        'Add New Bank Account Details
        pList("AccountNumber") = vAccountNo
        pList("SortCode") = vSortCode.Replace("-", "")
        pList("DefaultAccount") = "Y"
        DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContactAccounts, pList)
      End If
    End If
  End Sub

End Class