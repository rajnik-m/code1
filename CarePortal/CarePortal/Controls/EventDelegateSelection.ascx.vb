Public Class EventDelegateSelection
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectEventDelegates, tblDataEntry)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      FindEventDelegate(vList, False, True)
      If Me.FindControl("PageError") IsNot Nothing Then Me.FindControl("PageError").Visible = False
      If Me.FindControl("WarningMessage") IsNot Nothing Then Me.FindControl("WarningMessage").Visible = False
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub FindEventDelegate(ByVal pList As ParameterList, ByVal pShowCount As Boolean, ByVal pAddCheckBox As Boolean)
    pList("ContactNumber") = UserContactNumber()
    pList("WebPageItemNumber") = Me.WebPageItemNumber
    If InitialParameters.ContainsKey("BookingNumber") Then
      pList("BookingNumber") = InitialParameters("BookingNumber")
    ElseIf Request.QueryString("BN") IsNot Nothing Then
      pList("BookingNumber") = Request.QueryString("BN")
    Else
      pList("DocumentColumns") = "Y"
    End If
    If Request.QueryString("BN") IsNot Nothing Then pList("BookingNumber") = Request.QueryString("BN")
    Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactEventBookingDelegates, pList)
    Dim vDGR As DataGrid = TryCast(Me.FindControl("EventDelegate"), DataGrid)
    If vDGR IsNot Nothing Then
      DataHelper.FillGrid(vResult, vDGR)
      If pAddCheckBox Then
        Dim vTempColumn As New TemplateColumn()
        vTempColumn.HeaderText = ""
        vTempColumn.ItemTemplate = New CheckBoxTemplate("Select")
        vTempColumn.Visible = True
        vDGR.Columns.AddAt(0, vTempColumn)
        vDGR.DataBind()
      End If
      Dim vSelectPos As Integer = 0
      For vCount As Integer = 1 To vDGR.Columns.Count - 1
        Dim vBoundColumn As TemplateColumn = TryCast(vDGR.Columns(vCount), TemplateColumn)
        If vBoundColumn IsNot Nothing AndAlso vBoundColumn.HeaderText = "Select" Then
          vSelectPos = vCount
        End If
      Next
      For vRow As Integer = 0 To vDGR.Items.Count - 1
        If DirectCast(vDGR.Items(vRow).Cells(vSelectPos).Controls(0), ITextControl).Text = "Y" Then
          CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = True
        Else
          CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = False
        End If
        CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Text = ""
      Next
    End If
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If CType(sender, Button).Text = "Submit" Then
        UpdateDelegates(False)
      ElseIf CType(sender, Button).Text = "Add" Then
        UpdateDelegates()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enDuplicateRecord Then
        Me.FindControl("PageError").Visible = True
        SetLabelText("PageError", "Contact is already a delegate")
      ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enAppointmentConflict Then
        Me.FindControl("PageError").Visible = True
        SetLabelText("PageError", vEx.Message)
      Else
        ProcessError(vEx)
      End If
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub UpdateDelegates(Optional ByVal pAdd As Boolean = True)
    Dim vTotalDelegate As Integer
    Dim vBookingNumber As String = ""
    Dim vTotalCheck As Integer
    Dim vDelegateToBeAdded As New Hashtable
    Dim vDelegateToBeDeleted As New Hashtable
    Dim vExistingDelegate As New Hashtable
    Dim vQuantity As Integer
    If InitialParameters("BookingNumber") IsNot Nothing AndAlso InitialParameters("BookingNumber").ToString.Trim.Length > 0 Then
      vBookingNumber = InitialParameters("BookingNumber").ToString
    ElseIf Request.QueryString("BN") IsNot Nothing Then
      vBookingNumber = Request.QueryString("BN")
    End If
    Dim vList As New ParameterList(HttpContext.Current)
    If vBookingNumber IsNot Nothing AndAlso vBookingNumber.Trim.Length > 0 Then
      vList("EventNumber") = (IntegerValue(vBookingNumber) \ 10000).ToString
      vList("BookingNumber") = vBookingNumber
    Else
      vList("DocumentColumns") = "Y"
    End If
    Dim vReturnEventData As New ParameterList(DataHelper.SelectEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookings, vList))
    vQuantity = IntegerValue(vReturnEventData("Quantity").ToString)
    Dim vDGR As DataGrid = TryCast(Me.FindControl("EventDelegate"), DataGrid)
    'finding total checked text boxes
    For vCount As Integer = 0 To vDGR.Items.Count - 1
      If CType(vDGR.Items(vCount).Cells(0).Controls(0), CheckBox).Checked Then
        vTotalCheck = vTotalCheck + 1
      End If
    Next

    Dim vDataTable As DataTable = CType(vDGR.DataSource, DataSet).Tables("dataRow")

    Dim vTempDelegate As TempDelegate
    For vRow As Integer = 0 To vDataTable.Rows.Count - 1
      If vDataTable.Rows(vRow).Item("Select").ToString = "Y" Then
        vTotalDelegate = vTotalDelegate + 1
        vTempDelegate = New TempDelegate("", vDataTable.Rows(vRow).Item("EventDelegateNumber").ToString, "")
        vExistingDelegate.Add(vExistingDelegate.Count + 1, vTempDelegate)
        If CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = False Then
          vTempDelegate = New TempDelegate("", vDataTable.Rows(vRow).Item("EventDelegateNumber").ToString, "")
          vDelegateToBeDeleted.Add(vDelegateToBeDeleted.Count + 1, vTempDelegate)
        End If
      ElseIf vDataTable.Rows(vRow).Item("Select").ToString = "N" AndAlso CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = True Then
        vTempDelegate = New TempDelegate(vDataTable.Rows(vRow).Item("ContactNumber").ToString, "", vDataTable.Rows(vRow).Item("AddressNumber").ToString)
        vDelegateToBeAdded.Add(vDelegateToBeAdded.Count + 1, vTempDelegate)
      End If
    Next
    If GetTextBoxText("Surname").Length > 0 Then
      vTempDelegate = New TempDelegate(UserContactNumber().ToString, "", "1", GetTextBoxText("MemberNumber"), GetDropDownValue("Title"), GetTextBoxText("Forenames"), GetTextBoxText("Surname"), GetTextBoxText("EmailAddress"))
      vDelegateToBeAdded.Add(vDelegateToBeAdded.Count + 1, vTempDelegate)
    End If

    If pAdd Then
      If vQuantity < vTotalDelegate + vDelegateToBeAdded.Count - vDelegateToBeDeleted.Count Then
        Me.FindControl("WarningMessage").Visible = True
        SetLabelText("WarningMessage", "You can add maximum " & vQuantity.ToString & " delegates.")
        Exit Sub
      Else
        Me.FindControl("WarningMessage").Visible = False
      End If
    End If
    'Checking if total selected items match with quantity.
    If (vQuantity < vTotalDelegate + vDelegateToBeAdded.Count - vDelegateToBeDeleted.Count) OrElse _
      (vTotalDelegate + vDelegateToBeAdded.Count - vDelegateToBeDeleted.Count <= 0) Then
      Me.FindControl("WarningMessage").Visible = True
      SetLabelText("WarningMessage", "You must select between 1 and " & vQuantity.ToString & " Delegates")
      Exit Sub
    Else
      Me.FindControl("WarningMessage").Visible = False

    End If

    If vQuantity >= vTotalDelegate AndAlso vExistingDelegate.Count = vDelegateToBeAdded.Count AndAlso vDelegateToBeAdded.Count = vDelegateToBeDeleted.Count Then
      'when existing delegate count and count of delegates to be added and 
      'count of delegates to be deleted are same.
      Dim vItem As TempDelegate
      For Each vkey As Object In vExistingDelegate.Keys
        vItem = CType(vExistingDelegate(vkey), TempDelegate)
        vList = New ParameterList(HttpContext.Current)
        vList("EventDelegateNumber") = vItem.EventDelegateNumber
        vItem = CType(vDelegateToBeAdded(vkey), TempDelegate)
        If vItem.Surname.Length = 0 Then
          vList("ContactNumber") = vItem.ContactNumber
          vList("AddressNumber") = vItem.AddressNumber
        Else
          vList("ContactNumber") = UserContactNumber()
          vList("BookingNumber") = vBookingNumber
          vList("MemberNumber") = GetTextBoxText("MemberNumber")
          vList("AddressNumber") = vItem.AddressNumber
          vList("Title") = GetDropDownValue("Title")
          vList("Forenames") = GetTextBoxText("Forenames")
          vList("Surname") = GetTextBoxText("Surname")
          vList("EMailAddress") = GetTextBoxText("EmailAddress")
          vList("Source") = DefaultParameters("Source")
          vList("UserID") = UserContactNumber.ToString
        End If
        DataHelper.UpdateEventDelegate(vList)
      Next
      SubmitPage(vBookingNumber, pAdd)
      Exit Sub
    ElseIf vDelegateToBeAdded.Count >= vQuantity AndAlso vDelegateToBeDeleted.Count = vExistingDelegate.Count Then
      'when existing delegates are to be deleted are less than
      'delegates to be added
      Dim vItem As TempDelegate
      For Each vkey As Object In vDelegateToBeDeleted.Keys
        vItem = CType(vDelegateToBeDeleted(vkey), TempDelegate)
        vList = New ParameterList(HttpContext.Current)
        vList("EventDelegateNumber") = vItem.EventDelegateNumber
        vItem = CType(vDelegateToBeAdded(vkey), TempDelegate)
        vList("ContactNumber") = vItem.ContactNumber
        vList("AddressNumber") = vItem.AddressNumber
        DataHelper.UpdateEventDelegate(vList)
      Next
      'adding remaining delegate
      For Each vkey As Object In vDelegateToBeAdded.Keys
        If Not vDelegateToBeDeleted.ContainsKey(vkey) Then
          vItem = CType(vDelegateToBeAdded(vkey), TempDelegate)
          vList = New ParameterList(HttpContext.Current)
          vList("ContactNumber") = vItem.ContactNumber
          vList("BookingNumber") = vBookingNumber
          vList("AddressNumber") = vItem.AddressNumber
          vList("MemberNumber") = GetTextBoxText("MemberNumber")
          vList("Title") = GetDropDownValue("Title")
          vList("Forenames") = GetTextBoxText("Forenames")
          vList("Surname") = GetTextBoxText("Surname")
          vList("EmailAddress") = GetTextBoxText("EmailAddress")
          vList("UserID") = UserContactNumber.ToString
          vList("Source") = DefaultParameters("Source")
          DataHelper.AddEventDelegate(vList)
        End If
      Next
      SubmitPage(vBookingNumber, pAdd)
      Exit Sub
    ElseIf vDelegateToBeDeleted.Count = vQuantity AndAlso vDelegateToBeAdded.Count > 0 Then
      'when count of delegates to be deleted is equal to quantity
      'and count of delegate to be added is greater than 0
      Dim vItem As TempDelegate
      For Each vkey As Object In vDelegateToBeAdded.Keys
        vItem = CType(vDelegateToBeAdded(vkey), TempDelegate)
        vList = New ParameterList(HttpContext.Current)
        If vItem.Surname.Length = 0 Then
          vList("ContactNumber") = vItem.ContactNumber
          vList("AddressNumber") = vItem.AddressNumber
        Else
          vList("ContactNumber") = UserContactNumber()
          vList("BookingNumber") = vBookingNumber
          vList("MemberNumber") = GetTextBoxText("MemberNumber")
          vList("AddressNumber") = vItem.AddressNumber
          vList("Title") = GetDropDownValue("Title")
          vList("Forenames") = GetTextBoxText("Forenames")
          vList("Surname") = GetTextBoxText("Surname")
          vList("EMailAddress") = GetTextBoxText("EmailAddress")
          vList("Source") = DefaultParameters("Source")
          vList("UserID") = UserContactNumber.ToString
        End If
        vItem = CType(vDelegateToBeDeleted(vkey), TempDelegate)
        vList("EventDelegateNumber") = vItem.EventDelegateNumber
        DataHelper.UpdateEventDelegate(vList)
      Next
      'deleting remaining delegate
      For Each vkey As Object In vDelegateToBeDeleted.Keys
        If Not vDelegateToBeAdded.ContainsKey(vkey) Then
          vItem = CType(vDelegateToBeDeleted(vkey), TempDelegate)
          vList = New ParameterList(HttpContext.Current)
          vList("ContactNumber") = UserContactNumber()
          vList("BookingNumber") = vBookingNumber
          vList("EventDelegateNumber") = vItem.EventDelegateNumber
          vList("UserID") = UserContactNumber.ToString
          DataHelper.DeleteEventDelegate(vList)
        End If
      Next
      SubmitPage(vBookingNumber, pAdd)
    ElseIf vDelegateToBeDeleted.Count = 0 AndAlso GetTextBoxText("Surname").Length > 0 Then
      'When checkbox is not changed and data is entered in textbox then 
      'delegates will be added
      vList = New ParameterList(HttpContext.Current)
      vList("ContactNumber") = UserContactNumber()
      vList("BookingNumber") = vBookingNumber
      vList("MemberNumber") = GetTextBoxText("MemberNumber")
      vList("AddressNumber") = "1"
      vList("Title") = GetDropDownValue("Title")
      vList("Forenames") = GetTextBoxText("Forenames")
      vList("Surname") = GetTextBoxText("Surname")
      vList("EMailAddress") = GetTextBoxText("EmailAddress")
      vList("Source") = DefaultParameters("Source")
      vList("UserID") = UserContactNumber.ToString
      DataHelper.AddEventDelegate(vList)
      SubmitPage(vBookingNumber, pAdd)
    Else
      For vRow As Integer = 0 To vDataTable.Rows.Count - 1
        If vDataTable.Rows(vRow).Item("Select").ToString = "Y" AndAlso CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = False Then
          vList = New ParameterList(HttpContext.Current)
          vList("ContactNumber") = UserContactNumber()
          vList("BookingNumber") = vBookingNumber
          vList("EventDelegateNumber") = vDataTable.Rows(vRow).Item("EventDelegateNumber").ToString
          vList("UserID") = UserContactNumber.ToString
          DataHelper.DeleteEventDelegate(vList)
        ElseIf vDataTable.Rows(vRow).Item("Select").ToString = "N" AndAlso CType(vDGR.Items(vRow).Cells(0).Controls(0), CheckBox).Checked = True Then
          vList = New ParameterList(HttpContext.Current)
          vList("ContactNumber") = vDataTable.Rows(vRow).Item("ContactNumber").ToString
          vList("BookingNumber") = vBookingNumber
          vList("AddressNumber") = vDataTable.Rows(vRow).Item("AddressNumber").ToString
          vList("MemberNumber") = GetTextBoxText("MemberNumber")
          vList("Title") = GetDropDownValue("Title")
          vList("Forenames") = GetTextBoxText("Forenames")
          vList("Surname") = GetTextBoxText("Surname")
          vList("EmailAddress") = GetTextBoxText("EmailAddress")
          vList("UserID") = UserContactNumber.ToString
          vList("Source") = DefaultParameters("Source")
          DataHelper.AddEventDelegate(vList)
        End If
      Next
      If GetTextBoxText("Surname").Length > 0 Then
        vList = New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserContactNumber()
        vList("BookingNumber") = vBookingNumber
        vList("MemberNumber") = GetTextBoxText("MemberNumber")
        vList("AddressNumber") = "1"
        vList("Title") = GetDropDownValue("Title")
        vList("Forenames") = GetTextBoxText("Forenames")
        vList("Surname") = GetTextBoxText("Surname")
        vList("EMailAddress") = GetTextBoxText("EmailAddress")
        vList("Source") = DefaultParameters("Source")
        vList("UserID") = UserContactNumber.ToString
        DataHelper.AddEventDelegate(vList)
      End If
      SubmitPage(vBookingNumber, pAdd)
    End If
  End Sub
  Private Sub SubmitPage(ByVal pBookingNumber As String, ByVal pAdd As Boolean)
    If pAdd Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      FindEventDelegate(vList, False, False)
      SetTextBoxText("MemberNumber", "")
      SetDropDownText("Title", "")
      SetTextBoxText("Forenames", "")
      SetTextBoxText("Surname", "")
      SetTextBoxText("EmailAddress", "")
    Else
      Dim vSubmitParams As New StringBuilder
      With vSubmitParams
        .Append("&BN=" & pBookingNumber)
      End With
      GoToSubmitPage(vSubmitParams.ToString)
    End If
  End Sub
  Private Class TempDelegate
    Private mvEventDelegateNumber As String
    Private mvContactNumber As String
    Private mvAddressNumber As String
    Private mvMemberNumber As String
    Private mvTitle As String
    Private mvForename As String
    Private mvSurname As String
    Private mvEmail As String

    Public ReadOnly Property EventDelegateNumber() As String
      Get
        Return mvEventDelegateNumber
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As String
      Get
        Return mvContactNumber
      End Get
    End Property
    Public ReadOnly Property AddressNumber() As String
      Get
        Return mvAddressNumber
      End Get
    End Property
    Public ReadOnly Property MemberNumber() As String
      Get
        Return mvMemberNumber
      End Get
    End Property
    Public ReadOnly Property Title() As String
      Get
        Return mvTitle
      End Get
    End Property
    Public ReadOnly Property Forename() As String
      Get
        Return mvForename
      End Get
    End Property
    Public ReadOnly Property Surname() As String
      Get
        Return mvSurname
      End Get
    End Property
    Public ReadOnly Property Email() As String
      Get
        Return mvEmail
      End Get
    End Property
    Public Sub New(ByVal pContactNumber As String, ByVal pEventDelegateNumber As String, ByVal pAddressNumber As String, Optional ByVal pMemberNumber As String = "", Optional ByVal pTitle As String = "", Optional ByVal pForename As String = "", Optional ByVal pSurname As String = "", Optional ByVal pEmail As String = "")
      mvEventDelegateNumber = pEventDelegateNumber
      mvContactNumber = pContactNumber
      mvAddressNumber = pAddressNumber
      mvMemberNumber = pMemberNumber
      mvTitle = pTitle
      mvForename = pForename
      mvSurname = pSurname
      mvEmail = pEmail
    End Sub
  End Class
End Class