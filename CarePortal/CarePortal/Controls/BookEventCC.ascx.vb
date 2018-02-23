Partial Public Class BookEventCC
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvEventNumber As Integer
  Private mvOptionNumber As Integer
  Private mvProduct As String
  Private mvRate As String
  Private mvEventSource As String = String.Empty

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber"
      SupportsOnlineCCAuthorisation = True
      InitialiseControls(CareNetServices.WebControlTypes.wctBookEventCC, tblDataEntry, "CreditCardNumber,CardExpiryDate", "DirectNumber,MobileNumber")
      If (Request.QueryString("EN") IsNot Nothing AndAlso Request.QueryString("OP") IsNot Nothing) AndAlso (Request.QueryString("EN").Length > 0 AndAlso Request.QueryString("OP").Length > 0) Then
        mvEventNumber = IntegerValue(Request.QueryString("EN").ToString)
        mvOptionNumber = IntegerValue(Request.QueryString("OP").ToString)
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
        Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = vReturnList("ContactNumber")
        AddUserParameters(vList)
        vList("AddressNumber") = vReturnList("AddressNumber")
        'Now need to take the payment
        vList("EventNumber") = mvEventNumber.ToString
        vList("OptionNumber") = mvOptionNumber.ToString
        AddCCParameters(vList)
        AddDefaultParameters(vList)
        vList("Rate") = mvRate
        vList("Quantity") = GetTextBoxText("Quantity")
        vList("Amount") = GetTextBoxText("TotalAmount")
        vList("Notes") = GetTextBoxText("Notes")
        If vList.ContainsKey("UseSourceFromEvent") Then
          If BooleanValue(vList("UseSourceFromEvent").ToString) AndAlso mvEventSource.Length > 0 Then
            vList("Source") = mvEventSource
          End If
          vList.Remove("UseSourceFromEvent")
        End If
        Dim vSkipProcessing As Boolean
        Try
          DataHelper.AddEventBooking(vList)
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
    If mvEventNumber = 0 Then
      mvEventNumber = IntegerValue(InitialParameters("EventNumber").ToString)
      mvOptionNumber = IntegerValue(InitialParameters("OptionNumber").ToString)
    End If
    Dim vList As New ParameterList(HttpContext.Current)
    vList("EventNumber") = mvEventNumber.ToString     'InitialParameters("EventNumber")
    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventInformation, vList))
    If vRow IsNot Nothing Then SetTextBoxText("EventDesc", vRow("EventDesc").ToString)
    vList("OptionNumber") = mvOptionNumber.ToString   'InitialParameters("OptionNumber")
    mvEventSource = vRow("Source").ToString
    vRow = DataHelper.GetRowFromDataTable(DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingOptions, vList))
    mvProduct = vRow("ProductCode").ToString
    mvRate = vRow("RateCode").ToString
    vList = New ParameterList(HttpContext.Current)
    vList("Product") = mvProduct
    vList("Rate") = mvRate
    vRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtRates, vList))
    If vRow IsNot Nothing Then SetTextBoxText("Amount", CDbl(vRow("CurrentPrice")).ToString("#.00"))
    SetTextBoxText("TotalAmount", GetTextBoxText("Amount"))
    SetTextBoxText("Quantity", "1")
  End Sub
End Class