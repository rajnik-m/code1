Partial Public Class BookEvent
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvEventNumber As Integer
  Private mvOptionNumber As Integer
  Private mvProduct As String
  Private mvRate As String
  Private mvSessionList As String
  Private mvMinimumBookings As Integer
  Private mvMaximumBookings As Integer
  Private mvEventSource As String = String.Empty

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      mvUsesHiddenContactNumber = True
      mvHiddenFields = "HiddenAddressNumber,HiddenCurrentPrice,HiddenVatExclusive,HiddenPercentage"
      SupportsOnlineCCAuthorisation = False
      InitialiseControls(CareNetServices.WebControlTypes.wctBookEvent, tblDataEntry)
      SetControlVisible("WarningMessage1", False)
      SetControlVisible("WarningMessage2", False)
      SetControlVisible("WarningMessage3", False)
      SetControlVisible("WarningMessage4", False)
      SetControlVisible("WarningMessage5", False)
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not InWebPageDesigner() Then
      If IsValid() AndAlso ValidateQuantity(IntegerValue(GetTextBoxText("Quantity")), False) Then
        Try
          SetErrorLabel("")
          Dim vReturnList As ParameterList = AddNewContact(GetHiddenContactNumber() > 0)
          Dim vList As New ParameterList(HttpContext.Current)
          Dim vParamList As New ParameterList(HttpContext.Current)
          GetShoppingBasketTransaction(IntegerValue(vReturnList("ContactNumber").ToString), vList)               'now check if there is an existing transaction
          vList("ContactNumber") = vReturnList("ContactNumber")
          vList("AddressNumber") = vReturnList("AddressNumber")
          AddUserParameters(vList)
          If mvSessionList.Length > 0 Then vList("SessionNumbers") = mvSessionList
          'Now need to take the payment
          vList("EventNumber") = mvEventNumber.ToString
          vList("OptionNumber") = mvOptionNumber.ToString
          AddDefaultParameters(vList)
          vList("Rate") = mvRate
          If vList.ContainsKey("UseSourceFromEvent") Then
            If BooleanValue(vList("UseSourceFromEvent").ToString) AndAlso mvEventSource.Length > 0 Then
              vList("Source") = mvEventSource
            End If
            vList.Remove("UseSourceFromEvent")
          End If
          Dim vControlExists As Boolean
          vList("Quantity") = GetTextBoxText("Quantity", vControlExists)
          If Not vControlExists AndAlso mvMaximumBookings = mvMinimumBookings Then
            vList("Quantity") = mvMinimumBookings.ToString
          End If
          vList("Amount") = GetTextBoxText("TotalAmount")
          vList("Notes") = GetTextBoxText("Notes")
          If Me.FindControl("AdultQuantity") IsNot Nothing Then vList("AdultQuantity") = GetTextBoxText("AdultQuantity")
          If Me.FindControl("ChildQuantity") IsNot Nothing Then vList("ChildQuantity") = GetTextBoxText("ChildQuantity")
          vList("UserID") = vList("ContactNumber")
          Dim vSkipProcessing As Boolean
          Try
            vParamList = DataHelper.AddEventBooking(vList)
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
            GoToSubmitPage(String.Format("&BN={0}", vParamList("BookingNumber")))
          End If

        Catch vEX As ThreadAbortException
          Throw vEX
        Catch vException As Exception
          ProcessError(vException)
        End Try
      End If
    End If
  End Sub

  Public Sub ProcessChildControls(ByVal pList As ParameterList) Implements ICareParentWebControl.ProcessChildControls
    SubmitChildControls(pList)
  End Sub

  Private Function QueryStringHasValue(ByVal pItem As String) As Boolean
    If Request.QueryString(pItem) IsNot Nothing AndAlso Request.QueryString(pItem).Length > 0 Then Return True
  End Function

  Private Sub SetDefaults()
    If QueryStringHasValue("EN") AndAlso QueryStringHasValue("RA") AndAlso _
      (QueryStringHasValue("OP") OrElse QueryStringHasValue("ON")) Then
      mvEventNumber = IntegerValue(Request.QueryString("EN"))
      If QueryStringHasValue("OP") Then
        mvOptionNumber = IntegerValue(Request.QueryString("OP"))    'Handle either OP (correct) 
      Else
        mvOptionNumber = IntegerValue(Request.QueryString("ON"))    'Or ON (Incorrect but what was being passed before)
      End If
      mvRate = Request.QueryString("RA")
      If QueryStringHasValue("SL") Then
        mvSessionList = Request.QueryString("SL")
      Else
        mvSessionList = ""
      End If
    Else
      If InitialParameters.ContainsKey("EventNumber") Then mvEventNumber = IntegerValue(InitialParameters("EventNumber").ToString)
      If InitialParameters.ContainsKey("OptionNumber") Then mvOptionNumber = IntegerValue(InitialParameters("OptionNumber").ToString)
      If InitialParameters.ContainsKey("SessionList") Then
        mvSessionList = InitialParameters("SessionList").ToString
      Else
        mvSessionList = ""
      End If
      If InitialParameters.ContainsKey("Rate") Then mvRate = InitialParameters("Rate").ToString
    End If

    If mvEventNumber > 0 AndAlso mvOptionNumber > 0 Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("EventNumber") = mvEventNumber.ToString
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventInformation, vList))
      If BooleanValue(vRow("Booking").ToString) Then
        SetTextBoxText("EventDesc", vRow("EventDesc").ToString)
        vList("OptionNumber") = mvOptionNumber.ToString
        vList("SystemColumns") = "N"
        If UserContactNumber() > 0 Then vList("ContactNumber") = UserContactNumber()
        mvEventSource = vRow("Source").ToString
        vRow = DataHelper.GetRowFromDataTable(GetDataTable(DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebBookingOptions, vList)))
        If vRow IsNot Nothing Then
          mvProduct = vRow("Product").ToString
          mvRate = vRow("Rate").ToString
          mvMinimumBookings = IntegerValue(vRow.Item("MinimumBookings").ToString)
          mvMaximumBookings = IntegerValue(vRow.Item("MaximumBookings").ToString)
          Dim vProductMaxQuantity As Integer = IntegerValue(vRow.Item("ProductMaxQuantity").ToString)
          Dim vProductMinQuantity As Integer = IntegerValue(vRow.Item("ProductMinQuantity").ToString)
          If vProductMinQuantity > mvMinimumBookings AndAlso vProductMinQuantity <= mvMaximumBookings Then mvMinimumBookings = vProductMinQuantity
          If vProductMaxQuantity < mvMaximumBookings AndAlso vProductMaxQuantity >= mvMinimumBookings Then mvMaximumBookings = vProductMaxQuantity
          'If the Rate is VatExclusive then set the NetPrice on hidden current price field otherwise set the GrossPrice
          If BooleanValue(vRow("VatExclusive").ToString) Then
            SetHiddenText("HiddenCurrentPrice", vRow("NetPrice").ToString)
          Else
            SetHiddenText("HiddenCurrentPrice", vRow("GrossPrice").ToString)
          End If
          SetHiddenText("HiddenVatExclusive", vRow("VatExclusive").ToString)
          SetHiddenText("HiddenPercentage", vRow("Percentage").ToString)

          If Not IsPostBack Then
            SetTextBoxText("TotalAmount", vRow.Item("GrossPrice").ToString)
            SetTextBoxText("Amount", vRow.Item("GrossPrice").ToString)
            SetTextBoxText("Quantity", mvMinimumBookings.ToString)
            SetTotalAmount(mvMinimumBookings)
          End If
          SetTextBoxText("OptionDesc", vRow.Item("OptionDesc").ToString)
          'Check for Pick Sessions
          If BooleanValue(vRow.Item("PickSessions").ToString) AndAlso mvSessionList.Length = 0 Then
            SetControlVisible("WarningMessage2", True)
          End If
        Else
          SetControlVisible("WarningMessage3", True)
        End If
      Else
        SetControlVisible("WarningMessage1", True)
      End If
    End If
  End Sub

  Public Function ValidateQuantity(pTextBox As TextBox) As Boolean
    Dim vQuantity As Integer = IntegerValue(pTextBox.Text)
    Dim vSetQuantity As Boolean
    Select Case pTextBox.ID
      Case "AdultQuantity"
        vQuantity += IntegerValue(GetTextBoxText("ChildQuantity"))
        vSetQuantity = True
      Case "ChildQuantity"
        vQuantity += IntegerValue(GetTextBoxText("AdultQuantity"))
        vSetQuantity = True
    End Select
    Return ValidateQuantity(vQuantity, vSetQuantity)
  End Function

  Public Function ValidateQuantity(pQuantity As Integer, pSetQuantity As Boolean) As Boolean
    SetControlVisible("WarningMessage3", False)
    SetControlVisible("WarningMessage4", False)
    SetControlVisible("WarningMessage5", False)

    'If the quantity field is not visible but min and max are set the same then it is ok
    If pQuantity = 0 AndAlso mvMinimumBookings = mvMaximumBookings Then
      Return True
    End If

    If pQuantity > mvMaximumBookings Then
      Dim vMessage As String = GetLabelText("WarningMessage4")
      If vMessage.Length = 0 Then     'Handle a previously customised module
        SetLabelText("WarningMessage3", String.Format("Quantity cannot exceed {0} for this Booking Option", mvMaximumBookings))
        SetControlVisible("WarningMessage3", True)
      Else
        If vMessage.Contains("{0}") Then vMessage = String.Format(vMessage, mvMaximumBookings)
        SetLabelText("WarningMessage4", vMessage)
        SetControlVisible("WarningMessage4", True)
      End If
      Return False
    ElseIf pQuantity < mvMinimumBookings Then
      Dim vMessage As String = GetLabelText("WarningMessage5")
      If vMessage.Length = 0 Then     'Handle a previously customised module
        SetLabelText("WarningMessage3", String.Format("Quantity cannot be less than {0} for this Booking Option", mvMinimumBookings))
        SetControlVisible("WarningMessage3", True)
      Else
        If vMessage.Contains("{0}") Then vMessage = String.Format(vMessage, mvMinimumBookings)
        SetLabelText("WarningMessage5", vMessage)
        SetControlVisible("WarningMessage5", True)
      End If
      Return False
    Else
      If pSetQuantity Then
        SetTextBoxText("Quantity", pQuantity.ToString)
        SetControlEnabled("Quantity", False)
      End If
      Return True
    End If
  End Function

End Class