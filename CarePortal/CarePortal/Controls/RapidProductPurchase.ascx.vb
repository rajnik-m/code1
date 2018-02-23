Public Partial Class RapidProductPurchase
  Inherits CareWebControl

  Private mvBaseAmount As Double
  Private mvPageNumber As Integer
  Private mvProduct As String = ""
  Private mvRate As String = ""

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      SupportsOnlineCCAuthorisation = True
      InitialiseControls(CareNetServices.WebControlTypes.wctRapidProductPurchase, tblDataEntry, "", "Source,DonationAmount")
      If Request.QueryString("pn") IsNot Nothing Then mvPageNumber = IntegerValue(Request.QueryString("pn"))
      If Request.QueryString("bn") IsNot Nothing Then
        SetTextBoxText("BatchNumber", Request.QueryString("bn"))
        'Now set the focus to the ContactNumber control
        Dim vControl As Control = FindControlByName(tblDataEntry, "ContactNumber")
        If vControl IsNot Nothing Then vControl.Focus()
      End If
      If (Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("RA") IsNot Nothing) AndAlso (Request.QueryString("PR").Length > 0 AndAlso Request.QueryString("RA").Length > 0) Then
        'We expect both Product & Rate to be passed in, or neither
        mvProduct = Request.QueryString("PR")
        mvRate = Request.QueryString("RA")
      End If
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        SetErrorLabel("")
        Dim vList As New ParameterList(HttpContext.Current)
        Dim vAddressNumber As Long
        Dim vBatchComplete As Boolean
        Dim vBatchNumber As Integer
        Dim vBatchTotal As Double
        Dim vBatchValid As Boolean
        Dim vMsg As New StringBuilder
        Dim vNoEntries As Integer
        Dim vNoTransactions As Integer
        Dim vQuantity As Integer
        Dim vSource As String = ""
        Dim vTransTotal As Double

        vBatchNumber = IntegerValue(GetTextBoxText("BatchNumber"))
        vList("ContactNumber") = GetTextBoxText("ContactNumber")

        'Get AddressNumber
        Dim vDT As DataTable = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftContacts, vList)
        If vDT IsNot Nothing Then
          For Each vRow As DataRow In vDT.Rows
            If vRow("AddressNumber").ToString.Length > 0 Then
              vAddressNumber = IntegerValue(vRow("AddressNumber").ToString)
              Exit For
            End If
          Next
        Else
          Throw New Exception("Invalid Contact Number")
        End If

        'GetSourceCode
        If DefaultParameters.ContainsKey("Source") Then vSource = DefaultParameters("Source").ToString
        If vSource.Length = 0 Then
          If DefaultParameters.ContainsKey("PartSource") Then
            'Get the Source from the last mailing
            Dim vPartSource As String = DefaultParameters("PartSource").ToString
            If vPartSource.Length > 0 AndAlso vPartSource.Contains("+") Then vPartSource = vPartSource.Replace("+", ",")
            vList("PartSource") = vPartSource
            Dim vReturnList As New ParameterList(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactSourceFromLastMailing, vList))
            If vReturnList IsNot Nothing AndAlso vReturnList.Count > 0 Then vSource = vReturnList("Source").ToString
          End If
        End If
        If vSource.Length = 0 Then
          Throw New Exception("Source Code can not be null")
        End If

        vQuantity = IntegerValue((Val(GetTextBoxText("Amount")) / mvBaseAmount).ToString)

        vList = New ParameterList(HttpContext.Current)
        vList("ContactNumber") = GetTextBoxText("ContactNumber")
        vList("AddressNumber") = vAddressNumber.ToString
        vList("Product") = mvProduct
        vList("Rate") = mvRate
        vList("Quantity") = vQuantity
        vList("BankAccount") = DefaultParameters("BankAccount").ToString
        vList("Source") = vSource
        vList("BatchNumber") = vBatchNumber.ToString
        vList("Amount") = GetTextBoxText("Amount")
        If GetTextBoxText("DonationAmount").Length > 0 Then
          vList("DonationAmount") = GetTextBoxText("DonationAmount")
          If DefaultParameters("DonationProduct") Is Nothing Then Throw New Exception("Donation Product not specified") Else vList("DonationProduct") = DefaultParameters("DonationProduct").ToString
          If DefaultParameters("DonationRate") Is Nothing Then Throw New Exception("Donation Rate not specified") Else vList("DonationRate") = DefaultParameters("DonationRate").ToString
        End If

        'Check the Batch
        Dim vBatchList As ParameterList = New ParameterList(HttpContext.Current)
        vBatchList("BatchNumber") = vBatchNumber.ToString
        vDT = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftBatches, vBatchList)
        If vDT IsNot Nothing Then
          For Each vRow As DataRow In vDT.Rows
            If vRow.Item("BatchNumber").ToString.Length > 0 Then
              If vRow.Item("BatchType").ToString = "CC" Then
                vBatchValid = True
                'For CC batches, pass in CreditCardNumber & ExpiryDate OR NoClaimRequired
                If GetTextBoxText("CreditCardNumber").Length > 0 OrElse GetTextBoxText("CardExpiryDate").Length > 0 Then
                  vList("CreditCardNumber") = GetTextBoxText("CreditCardNumber")
                  vList("CardExpiryDate") = GetTextBoxText("CardExpiryDate")
                  vList("NoClaimRequired") = "N"
                  If GetTextBoxText("CreditCardNumber").Length > 0 AndAlso GetTextBoxText("SecurityCode").Length > 0 Then
                    vList("SecurityCode") = GetTextBoxText("SecurityCode")
                    vList("GetAuthorisation") = "Y"
                  End If
                Else
                  vList("NoClaimRequired") = "Y"
                End If
                If GetTextBoxText("AuthorisationCode").Length > 0 Then vList("AuthorisationCode") = GetTextBoxText("AuthorisationCode")
              ElseIf vRow.Item("BatchType").ToString = "CA" Then
                vBatchValid = True
              End If
              Exit For
            End If
          Next
        End If
        If vBatchValid = False Then
          Throw New Exception(String.Format("Batch Number {0} is invalid", vBatchNumber))
        End If

        Dim vSkipProcessing As Boolean
        Try
          'Add the ProductSale
          'Debug.Print(vList.XMLParameterString)
          DataHelper.AddProductSale(vList)
        Catch vEx As ThreadAbortException
          Throw vEx
        Catch vEx As CareException
          SetErrorLabel(vEx.Message)
          vSkipProcessing = True
        End Try
        If vSkipProcessing = False Then
          'Check Batch totals etc.
          vBatchComplete = False
          vDT = DataHelper.FindDataTable(CareNetServices.XMLDataFinderTypes.xdftBatches, vBatchList)
          If vDT IsNot Nothing Then
            For Each vRow As DataRow In vDT.Rows
              If vRow.Item("BatchNumber").ToString.Length > 0 Then
                vNoEntries = IntegerValue(vRow.Item("NumberOfEntries").ToString)
                vNoTransactions = IntegerValue(vRow.Item("NumberOfTransactions").ToString)
                vBatchTotal = Val(vRow.Item("BatchTotal").ToString)
                vTransTotal = Val(vRow.Item("TransactionTotal").ToString)
                If vNoEntries > 0 OrElse vBatchTotal > 0 Then
                  If vNoEntries > 0 AndAlso (vNoTransactions >= vNoEntries) Then
                    vBatchComplete = True
                  ElseIf vBatchTotal > 0 AndAlso (vTransTotal >= vBatchTotal) Then
                    vBatchComplete = True
                  End If
                End If
                If vBatchComplete Then
                  vMsg.AppendLine("Batch {0} has now been completed; no more transactions can be added.")
                  If vNoEntries > 0 Then
                    vMsg.AppendLine(String.Format("Expected Number of Transactions = {0}, Transactions Entered = {1}.", vNoEntries.ToString, vNoTransactions.ToString))
                  End If
                  If vBatchTotal > 0 Then
                    vMsg.AppendLine(String.Format("Expected Total Amount = {0}, Amount Entered = {1}.", vBatchTotal.ToString("#.00"), vTransTotal.ToString("#.00")))
                  End If
                End If
                Exit For
              End If
            Next
          End If

          If vBatchComplete Then
            'Display message to user
            ShowMessageOnly(String.Format(vMsg.ToString, vBatchNumber))
          Else
            'Loop back round and redisplay this page, pre-populating the BatchNumber
            ProcessRedirect("Default.aspx?pn=" & mvPageNumber.ToString & "&bn=" & vBatchNumber)
          End If
        End If
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub SetDefaults()
    If mvProduct.Length = 0 Then
      mvProduct = InitialParameters("Product").ToString
      mvRate = InitialParameters("Rate").ToString
    End If
    SetAmountOrBalance("Amount")
    mvBaseAmount = Val(GetTextBoxText("Amount"))
  End Sub
End Class