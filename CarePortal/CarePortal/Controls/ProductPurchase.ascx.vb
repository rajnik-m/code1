Public Class ProductPurchase
  Inherits CareWebControl
  Implements ICareParentWebControl
  Private mvProduct As String = String.Empty
  Private mvRate As String = String.Empty
  Private mvContactNumber As Integer
  Private mvProductDesc As String = String.Empty
  Private mvAmount As String = String.Empty
  Private mvTotalAmount As String = String.Empty
  Private mvBatchNumber As String = String.Empty
  Private mvTransactionNumber As String = String.Empty
  Private mvLineNumber As String = String.Empty
  Private mvIsUpdate As Boolean = False

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      SupportsOnlineCCAuthorisation = False
      InitialiseControls(CareNetServices.WebControlTypes.wctProductPurchase, tblDataEntry)
      If (Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("RA") IsNot Nothing) AndAlso (Request.QueryString("PR").Length > 0 AndAlso Request.QueryString("RA").Length > 0) Then
        'We expect both Product & Rate to be passed in, or neither
        mvProduct = Request.QueryString("PR")
        mvRate = Request.QueryString("RA")
      End If
      If (Request.QueryString("BN") IsNot Nothing AndAlso Request.QueryString("BN").Length > 0) AndAlso _
        (Request.QueryString("TN") IsNot Nothing AndAlso Request.QueryString("TN").Length > 0) AndAlso _
        (Request.QueryString("LN") IsNot Nothing AndAlso Request.QueryString("LN").Length > 0) Then
        mvBatchNumber = Request.QueryString("BN")
        mvTransactionNumber = Request.QueryString("TN")
        mvLineNumber = Request.QueryString("LN")
        mvIsUpdate = True
      End If
      AddHiddenField("HiddenCurrentPrice")
      AddHiddenField("HiddenVatExclusive")
      AddHiddenField("HiddenPercentage")
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
        If Not mvIsUpdate Then
          SetErrorLabel("")
          'Finding whether Existing Transaction is in Progress
          Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
          Dim vList As New ParameterList(HttpContext.Current)
          GetShoppingBasketTransaction(IntegerValue(vReturnList("ContactNumber").ToString), vList)               'now check if there is an existing transaction
          vList("ContactNumber") = vReturnList("ContactNumber")
          vList("AddressNumber") = vReturnList("AddressNumber")
          'Now need to take the quantity and total Amount
          vList("Product") = mvProduct
          vList("Rate") = mvRate
          vList("Quantity") = GetTextBoxText("Quantity")
          vList("Amount") = GetTextBoxText("TotalAmount")
          vList("UserID") = UserContactNumber.ToString
          AddUserParameters(vList)
          AddDefaultParameters(vList)
          Dim vSkipProcessing As Boolean
          Try
            DataHelper.AddProductSale(vList)
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vEx As CareException
            SetErrorLabel(vEx.Message)
            SetHiddenText("HiddenContactNumber", vReturnList("ContactNumber").ToString)
            SetHiddenText("HiddenAddressNumber", vReturnList("AddressNumber").ToString)
            vSkipProcessing = True
          End Try
          If vSkipProcessing = False Then
            ProcessChildControls(vReturnList)
            GoToSubmitPage()
          End If
        Else
          Try
            Dim vList As New ParameterList(HttpContext.Current)
            vList("Product") = mvProduct
            vList("Rate") = mvRate
            vList("Quantity") = GetTextBoxText("Quantity")
            vList("Amount") = GetTextBoxText("TotalAmount")
            vList("UserID") = UserContactNumber.ToString
            vList("ContactNumber") = UserContactNumber.ToString
            vList("BatchNumber") = mvBatchNumber
            vList("TransactionNumber") = mvTransactionNumber
            vList("LineNumber") = mvLineNumber
            DataHelper.UpdateProvisionalTransaction(vList)
            GoToSubmitPage()
          Catch vEx As Exception
            SetErrorLabel(vEx.Message)
          End Try
        End If
      Catch vEX As ThreadAbortException
        Throw vEX
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub
  Private Sub SetDefaults()
    If Not mvIsUpdate Then
      Dim vAttrs As New StringBuilder
      If Not InitialParameters.ContainsKey("Product") AndAlso mvProduct.Length = 0 Then vAttrs.Append("Product")
      If Not InitialParameters.ContainsKey("Rate") AndAlso mvRate.Length = 0 Then
        If vAttrs.Length > 0 Then vAttrs.Append(",")
        vAttrs.Append(" Rate")
      End If
      If vAttrs.Length > 0 Then
        SetErrorLabel(vAttrs.ToString & " has not been set")
      Else
        If mvProduct.Length = 0 Or mvRate.Length = 0 Then
          mvProduct = InitialParameters("Product").ToString
          mvRate = InitialParameters("Rate").ToString
        End If
      End If
      If Not vAttrs.Length > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("SystemColumns") = "N"
        mvContactNumber = UserContactNumber()
        If mvContactNumber > 0 Then
          vList("ContactNumber") = mvContactNumber
        End If
        vList("Product") = mvProduct
        Dim vRow As DataRow
        vRow = DataHelper.GetRowFromDataTable(GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebProducts, vList)))
        If vRow IsNot Nothing Then
          mvProductDesc = vRow.Item("ProductDesc").ToString()
          mvAmount = vRow.Item("GrossPrice").ToString
          mvTotalAmount = mvAmount
          SetHiddenText("HiddenCurrentPrice", vRow("CurrentPrice").ToString)
          SetHiddenText("HiddenVatExclusive", vRow("VatExclusive").ToString)
          SetHiddenText("HiddenPercentage", vRow("Percentage").ToString)
        End If
        DirectCast(FindControlByName(Me, "Product"), TextBox).ReadOnly = True
      End If
      If Not mvIsUpdate Then
        If Not IsPostBack Then
          SetTextBoxText("Amount", mvAmount)
          SetTextBoxText("TotalAmount", mvTotalAmount)
          SetTextBoxText("Product", mvProductDesc)
          SetTextBoxText("Quantity", "1")
        End If
      End If
    Else
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vResult As String
      Dim vDataTable As DataTable
      vList("BatchNumber") = mvBatchNumber
      vList("TransactionNumber") = mvTransactionNumber
      vList("LineNumber") = mvLineNumber
      vResult = DataHelper.GetTransactionData(CareNetServices.XMLTransactionDataSelectionTypes.xtdtTransactionAnalysis, vList)
      vDataTable = GetDataTable(vResult)
      If vDataTable IsNot Nothing And vDataTable.Rows.Count > 0 Then
        SetErrorLabel("")
        If Not IsPostBack Then
          SetTextBoxText("Amount", FixTwoPlaces((CDbl(vDataTable.Rows(0).Item("Amount").ToString) / CDbl(vDataTable.Rows(0).Item("Quantity").ToString)).ToString).ToString)
          SetTextBoxText("TotalAmount", CStr(FixTwoPlaces(vDataTable.Rows(0).Item("Amount").ToString)))
          SetTextBoxText("Product", vDataTable.Rows(0).Item("ProductDesc").ToString)
          SetTextBoxText("Quantity", vDataTable.Rows(0).Item("Quantity").ToString)
          Dim vRow As DataRow
          vList = New ParameterList(HttpContext.Current)
          vList("Product") = vDataTable.Rows(0).Item("Product").ToString
          vList("Rate") = vDataTable.Rows(0).Item("Rate").ToString
          If UserContactNumber() > 0 Then
            vList("ContactNumber") = UserContactNumber()
          End If
          vRow = DataHelper.GetRowFromDataTable(GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebProducts, vList)))
          If vRow IsNot Nothing Then
            SetHiddenText("HiddenCurrentPrice", vRow("CurrentPrice").ToString)
            SetHiddenText("HiddenVatExclusive", vRow("VatExclusive").ToString)
            SetHiddenText("HiddenPercentage", vRow("Percentage").ToString)
          End If
        End If
        mvProduct = vDataTable.Rows(0).Item("Product").ToString
        mvRate = vDataTable.Rows(0).Item("Rate").ToString
      End If
    End If
  End Sub
End Class