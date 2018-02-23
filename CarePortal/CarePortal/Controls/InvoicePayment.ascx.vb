Public Class InvoicePayment
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctInvoicePayment, tblDataEntry)
      GetInvoices()
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vBaseList As BaseDataList = DirectCast(Me.FindControl("InvoiceData"), BaseDataList)
        'Get the total number of items in the BaseDataList
        Dim vItemsCount As Integer = If(TypeOf (vBaseList) Is DataGrid, DirectCast(vBaseList, DataGrid).Items.Count, DirectCast(vBaseList, DataList).Items.Count)
        Dim vFound As Boolean = False
        'At least one record must be selected for submission
        For vIndex As Integer = 0 To vItemsCount - 1
          Dim vCheckBox As CheckBox = GetBaseDataListCheckBox(vBaseList, vIndex)
          If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then
            vFound = True
            Exit For
          End If
        Next
        If vFound = False Then
          DirectCast(Me.FindControl("WarningMessage2"), Label).Visible = True           'Show warning message 2 when no record is selected
          Exit Sub
        End If

        Dim vPayList As New ParameterList(HttpContext.Current)
        GetShoppingBasketTransaction(UserContactNumber, vPayList)        'Find any existing Provisional Batch and Transaction
        vPayList("ContactNumber") = UserContactNumber()
        vPayList("AddressNumber") = UserAddressNumber()
        vPayList("BankAccount") = DefaultParameters("BankAccount")
        vPayList("Source") = DefaultParameters("Source")
        AddUserParameters(vPayList)

        Dim vSkipProcessing As Boolean
        Dim vPaidInvoices As New List(Of Integer)
        Dim vSkippedInvoice As String = ""
        For vIndex As Integer = 0 To vItemsCount - 1
          Dim vCheckBox As CheckBox = GetBaseDataListCheckBox(vBaseList, vIndex)
          If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then 'Only proceed for selected items
            DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Rows(vIndex)("PayCheck") = "Y" 'Set this to be used by DataBind (below)
            If vSkipProcessing = False Then 'Only proceed if no error has occurred
              vPayList("InvoiceNumber") = GetBaseDataListItem(vBaseList, vIndex, "InvoiceNumber")
              vPayList("Amount") = GetBaseDataListItem(vBaseList, vIndex, "Outstanding")
              Try
                Dim vReturnList As ParameterList = DataHelper.AddInvoicePayment(vPayList)
                If vPayList.Contains("BatchNumber") = False Then
                  'New Batch and Transaction number, use this for later payments
                  vPayList("BatchNumber") = vReturnList("BatchNumber")
                  vPayList("TransactionNumber") = vReturnList("TransactionNumber")
                End If
              Catch vEx As ThreadAbortException
                Throw vEx
              Catch vEx As CareException
                'Error on adding a payment. Display the error message and skip processing.
                SetErrorLabel(vEx.Message)
                vSkipProcessing = True
              End Try
              If vSkipProcessing Then
                If vPaidInvoices.Count = 0 Then
                  'The error occurred when not a single Invoice has been paid. Just highlight the error record.
                  If TypeOf (vBaseList) Is DataGrid Then
                    DirectCast(vBaseList, DataGrid).SelectedIndex = vIndex
                  Else
                    DirectCast(vBaseList, DataList).SelectedIndex = vIndex
                  End If
                Else
                  'Save the invoice number with error to be used later for highlighting the record.
                  vSkippedInvoice = vPayList("InvoiceNumber").ToString
                End If
              Else
                'All paid payment plan records should be removed
                vPaidInvoices.Add(IntegerValue(vPayList("InvoiceNumber").ToString))
              End If
            End If
          End If
        Next
        If vPaidInvoices.Count > 0 Then
          Dim vInvoiceTable As DataTable = DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow")
          For Each vPaidInvoiceNumber As Integer In vPaidInvoices
            For vIndex As Integer = vInvoiceTable.Rows.Count - 1 To 0 Step -1
              If vInvoiceTable.Rows(vIndex)("InvoiceNumber").ToString = vPaidInvoiceNumber.ToString Then
                vInvoiceTable.Rows.Remove(vInvoiceTable.Rows(vIndex))
                Exit For
              End If
            Next
          Next
          vBaseList.DataBind()  'Bind the data again to get rid of paid invoices and to keep check boxes checked state
          If vSkippedInvoice.Length > 0 Then
            For vSkippedIndex As Integer = vInvoiceTable.Rows.Count - 1 To 0 Step -1
              If vInvoiceTable.Rows(vSkippedIndex)("InvoiceNumber").ToString = vSkippedInvoice Then
                If TypeOf (vBaseList) Is DataGrid Then
                  DirectCast(vBaseList, DataGrid).SelectedIndex = vSkippedIndex
                Else
                  DirectCast(vBaseList, DataList).SelectedIndex = vSkippedIndex
                End If
                Exit For
              End If
            Next
          End If
        End If
        If vSkipProcessing = False Then GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub GetInvoices()
    'First see if there is a shopping basket transaction
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vInvoiceList As New StringList("")
    If GetShoppingBasketTransaction(UserContactNumber, vList) Then
      'If so lets go and read the data from it
      Dim vTransactionTable As DataTable = DataHelper.GetDataTable(DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, vList))
      For Each vRow As DataRow In vTransactionTable.Rows
        If vRow("InvoiceNumber").ToString.Length > 0 Then vInvoiceList.Add(vRow("InvoiceNumber").ToString)
      Next
    End If
    DirectCast(Me.FindControl("WarningMessage2"), Label).Visible = False
    Dim vDGR As DataGrid = CType(Me.FindControl("InvoiceData"), DataGrid)
    vList = New ParameterList(HttpContext.Current)
    vList("Company") = InitialParameters("Company")
    Dim vContactNumber As Integer = GetContactNumberFromParentGroup()
    vList("ContactNumber") = vContactNumber.ToString
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCustomers, vList)
    If vTable IsNot Nothing Then
      vTable.DefaultView.RowFilter = "Company = '" & InitialParameters("Company").ToString & "'"
      vList("SalesLedgerAccount") = vTable.Rows(0)("SalesLedgerAccount").ToString
      vList("SystemColumns") = "Y"
      vList.Add("WPD", "Y")
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      vList("InvoiceNumbersAdded") = ""
      Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactOutstandingInvoices, vList)
      Dim vRestriction As String = ""
      If vInvoiceList.Count > 0 Then vRestriction = String.Format("InvoiceNumber NOT IN ({0})", vInvoiceList.ItemList)
      DataHelper.FillGrid(vResult, vDGR, vRestriction)
      If vDGR.Items.Count <= 0 Then
        ShowMessageOnly(Me.FindControl("WarningMessage1"))
      Else
        DirectCast(Me.FindControl("WarningMessage1"), Label).Visible = False
        vDGR.Visible = True
      End If
    Else
      ShowMessageOnly(Me.FindControl("WarningMessage1"))
    End If
  End Sub

End Class



