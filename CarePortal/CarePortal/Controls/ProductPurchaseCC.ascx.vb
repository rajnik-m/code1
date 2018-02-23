Partial Public Class ProductPurchaseCC
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvProduct As String = ""
  Private mvRate As String = ""

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      SupportsOnlineCCAuthorisation = True
      InitialiseControls(CareNetServices.WebControlTypes.wctProductPurchaseCC, tblDataEntry, "CreditCardNumber,CardExpiryDate", "DirectNumber,MobileNumber")
      If (Request.QueryString("PR") IsNot Nothing AndAlso Request.QueryString("RA") IsNot Nothing) AndAlso (Request.QueryString("PR").Length > 0 AndAlso Request.QueryString("RA").Length > 0) Then
        'We expect both Product & Rate to be passed in, or neither
        mvProduct = Request.QueryString("PR")
        mvRate = Request.QueryString("RA")
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
        SetErrorLabel("")
        Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vReturnList("ContactNumber")
        vList("AddressNumber") = vReturnList("AddressNumber")
        'Now need to take the payment
        vList("Product") = mvProduct
        vList("Rate") = mvRate
        vList("Quantity") = GetTextBoxText("Quantity")
        vList("Amount") = GetTextBoxText("TotalAmount")
        AddCCParameters(vList)
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
    If mvProduct.Length = 0 Then
      mvProduct = InitialParameters("Product").ToString
      mvRate = InitialParameters("Rate").ToString
    End If
    Dim vList As New ParameterList(HttpContext.Current)
    vList("SystemColumns") = "N"
    If UserContactNumber() > 0 Then
      vList("ContactNumber") = UserContactNumber()
    End If
    vList("Product") = mvProduct
    Dim vProductDesc As String = String.Empty
    Dim vAmount As String = String.Empty
    Dim vTotalAmount As String = String.Empty
    Dim vRow As DataRow
    vRow = DataHelper.GetRowFromDataTable(GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebProducts, vList)))
    If vRow IsNot Nothing Then
      vProductDesc = vRow.Item("ProductDesc").ToString()
      vAmount = vRow.Item("GrossPrice").ToString
      vTotalAmount = vAmount
      SetHiddenText("HiddenCurrentPrice", vRow("CurrentPrice").ToString)
      SetHiddenText("HiddenVatExclusive", vRow("VatExclusive").ToString)
      SetHiddenText("HiddenPercentage", vRow("Percentage").ToString)
    End If
    If Not IsPostBack Then
      SetTextBoxText("Amount", vAmount)
      SetTextBoxText("TotalAmount", vTotalAmount)
      SetTextBoxText("Product", vProductDesc)
      SetTextBoxText("Quantity", "1")
    End If
  End Sub
End Class