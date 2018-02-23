Public Class PayMultiplePaymentPlans
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vShowOrganisationPP As Boolean = False
      InitialiseControls(CareNetServices.WebControlTypes.wctPayMultiplePaymentPlans, tblDataEntry)
      If InitialParameters.ContainsKey("MembershipFor") AndAlso InitialParameters("MembershipFor").ToString.Length > 0 AndAlso InitialParameters("MembershipFor").ToString.ToUpper = "O" Then vShowOrganisationPP = True

      'Hide any warning messages
      Dim vWarning1 As Control = FindControlByName(Me, "WarningMessage1")
      If vWarning1 IsNot Nothing Then vWarning1.Visible = False
      If FindControlByName(Me, "WarningMessage2") IsNot Nothing Then FindControlByName(Me, "WarningMessage2").Visible = False

      'Get all due Payment Plan Payments
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = UserContactNumber()
      Dim vResult As String = String.Empty
      Dim vOrganisationPP As New DataSet


      If Not vShowOrganisationPP Then
        vList("SystemColumns") = "Y"
        vList("WebPageItemNumber") = Me.WebPageItemNumber
        vResult = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlanPayments, vList)
      Else
        vList("Current") = "Y"
        Dim vDT As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vList)

        For Each vDataRow As DataRow In vDT.Rows
          Dim vOrganisationList As New ParameterList(HttpContext.Current)
          vOrganisationList("ContactNumber") = vDataRow("ContactNumber").ToString
          vOrganisationList("SystemColumns") = "Y"
          vOrganisationList("WebPageItemNumber") = Me.WebPageItemNumber
          vResult = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPaymentPlanPayments, vOrganisationList)
          If vResult.Length > 0 Then
            If vOrganisationPP.Tables.Count = 0 Then
              Dim vStringReader As New System.IO.StringReader(vResult)
              vOrganisationPP.ReadXml(vStringReader, XmlReadMode.Auto)
            Else
              Dim vSr As New System.IO.StringReader(vResult)
              Dim vDataSet As New DataSet
              vDataSet.ReadXml(vSr, XmlReadMode.Auto)
              If vDataSet.Tables.Count > 0 AndAlso vDataSet.Tables.Contains("DataRow") Then
                For Each vDR As DataRow In vDataSet.Tables("DataRow").Rows
                  If vOrganisationPP.Tables.Contains("DataRow") Then
                    vOrganisationPP.Tables("DataRow").ImportRow(vDR)
                  Else
                    vOrganisationPP.Tables.Add(vDataSet.Tables("DataRow").Copy())
                  End If
                Next
              End If
            End If
          End If
        Next
      End If

      If vShowOrganisationPP Then
        Dim vStringWriter As New System.IO.StringWriter()
        vOrganisationPP.WriteXml(vStringWriter)
        vResult = vStringWriter.ToString
      End If

      Dim vBaseList As BaseDataList = TryCast(Me.FindControl("PaymentPlans"), BaseDataList)
      If vBaseList IsNot Nothing AndAlso vResult.Length > 0 Then
        'For Column Format,set the DataKeyField to generate DataKeys for each item(row)
        If TypeOf vBaseList Is DataList Then vBaseList.DataKeyField = "PaymentPlanNumber"
        DataHelper.FillGrid(vResult, vBaseList)
        'Show warning message 1 when no records are found
        If (TypeOf vBaseList Is DataList AndAlso DirectCast(vBaseList, DataList).Items.Count = 0) OrElse _
          (TypeOf vBaseList Is DataGrid AndAlso DirectCast(vBaseList, DataGrid).Items.Count = 0) Then
          'BR19328
          If vWarning1 IsNot Nothing Then
            vWarning1.Visible = True
            Dim vSubmitButton As Control = FindControlByName(Me, "Submit")
            If vSubmitButton IsNot Nothing Then vSubmitButton.Visible = False
            vBaseList.Visible = False
          End If
        End If
      Else
        'Show warning message 1 when grid is not visible
        ShowMessageOnly(vWarning1)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        'Clear error and warning messages
        SetErrorLabel("")
        Dim vWarning2 As Control = FindControlByName(Me, "WarningMessage2")
        If vWarning2 IsNot Nothing Then vWarning2.Visible = False

        Dim vBaseList As BaseDataList = DirectCast(Me.FindControl("PaymentPlans"), BaseDataList)
        'Get the total number of items in the BaseDataList
        Dim vItemsCount As Integer = If(TypeOf (vBaseList) Is DataGrid, DirectCast(vBaseList, DataGrid).Items.Count, DirectCast(vBaseList, DataList).Items.Count)
        Dim vContinue As Boolean = False
        'At least one record must be selected for submission
        For vIndex As Integer = 0 To vItemsCount - 1
          Dim vCheckBox As CheckBox = GetBaseDataListCheckBox(vBaseList, vIndex)
          If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then
            vContinue = True
            Exit For
          End If
        Next
        If vContinue = False Then
          'Show warning message 2 when no record is selected
          If vWarning2 IsNot Nothing Then vWarning2.Visible = True
          Exit Sub
        End If
        Dim vPayList As New ParameterList(HttpContext.Current)
        GetShoppingBasketTransaction(UserContactNumber, vPayList)        'Find any existing Provisional Batch and Transaction
        vPayList("ContactNumber") = UserContactNumber()
        vPayList("AddressNumber") = UserAddressNumber()
        vPayList("BankAccount") = InitialParameters("BankAccount")
        vPayList("Source") = InitialParameters("Source")
        AddUserParameters(vPayList)

        Dim vSkipProcessing As Boolean
        Dim vPaidPaymentPlans As New StringBuilder
        Dim vSkippedPaymentPlan As String = ""
        For vIndex As Integer = 0 To vItemsCount - 1
          Dim vCheckBox As CheckBox = GetBaseDataListCheckBox(vBaseList, vIndex)
          If vCheckBox IsNot Nothing AndAlso vCheckBox.Checked Then 'Only proceed for selected items
            DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Rows(vIndex)("CheckColumn") = "Y" 'Set this to be used by DataBind (below)
            If vSkipProcessing = False Then 'Only proceed if no error has occurred
              vPayList("PaymentPlanNumber") = GetBaseDataListItem(vBaseList, vIndex, "PaymentPlanNumber")
              vPayList("Amount") = GetBaseDataListItem(vBaseList, vIndex, "NextPaymentAmount")
              Try
                Dim vReturnList As ParameterList = DataHelper.AddPaymentPlanPayment(vPayList)
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
                If vPaidPaymentPlans.Length = 0 Then
                  'The error occurred when not a single Payment Plan has been paid. Just high light the error record.
                  If TypeOf (vBaseList) Is DataGrid Then
                    DirectCast(vBaseList, DataGrid).SelectedIndex = vIndex
                  Else
                    DirectCast(vBaseList, DataList).SelectedIndex = vIndex
                  End If
                Else
                  'Save the payment plan number with error to be used later for high lighting the record.
                  vSkippedPaymentPlan = vPayList("PaymentPlanNumber").ToString
                End If
              Else
                'All paid payment plan records should be removed
                If vPaidPaymentPlans.Length > 0 Then vPaidPaymentPlans.Append(",")
                vPaidPaymentPlans.Append(vPayList("PaymentPlanNumber"))
              End If
            End If
          End If
        Next
        If vPaidPaymentPlans.Length > 0 Then
          Dim vNumbers() As String = vPaidPaymentPlans.ToString.Split(","c)
          For vIndex As Integer = 0 To vNumbers.Length - 1
            'Find the data row of the payment plan in the data source of the base list
            Dim vRow As DataRow = DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Select("PaymentPlanNumber = " & vNumbers(vIndex))(0)
            'Remove this row from the data source
            DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Rows.Remove(vRow)
          Next
          vBaseList.DataBind()  'Bind the data again to get rid of paid payment plans and to keep check boxes checked state
          If vSkippedPaymentPlan.Length > 0 Then
            'If an error occurred then find the row and high light it
            Dim vRow As DataRow = DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Select("PaymentPlanNumber = " & vSkippedPaymentPlan)(0)
            Dim vIndex As Integer = DirectCast(vBaseList.DataSource, DataSet).Tables("DataRow").Rows.IndexOf(vRow)
            If TypeOf (vBaseList) Is DataGrid Then
              DirectCast(vBaseList, DataGrid).SelectedIndex = vIndex
            Else
              DirectCast(vBaseList, DataList).SelectedIndex = vIndex
            End If
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

End Class