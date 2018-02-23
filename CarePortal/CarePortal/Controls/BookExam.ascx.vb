Partial Public Class BookExam
  Inherits CareWebControl
  Implements ICareParentWebControl

  Private mvTopLevelUnitId As Integer
  Private mvExamUnitId As Integer
  Private mvExamSessionId As Integer
  Private mvExamCentreId As Integer
  Private mvExamCentreCode As String
  Private mvUnitsToBook As List(Of Integer)
  Private mvError As Boolean

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctBookExam, tblDataEntry)
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Not InWebPageDesigner() Then
      Dim vValid As Boolean = IsValid()
      If vValid And Not mvError Then
        Try
          SetErrorLabel("")
          Dim vParamList As New ParameterList(HttpContext.Current)
          Dim vList As New ParameterList(HttpContext.Current)
          GetShoppingBasketTransaction(UserContactNumber, vList)               'now check if there is an existing transaction
          vList("ContactNumber") = UserContactNumber()
          vList("AddressNumber") = UserAddressNumber()
          AddUserParameters(vList)
          'Now need to take the payment
          vList("ExamCentreId") = mvExamCentreId
          Dim vExamUnits As New StringBuilder
          For Each vUnit As Integer In mvUnitsToBook
            If vExamUnits.Length > 0 Then vExamUnits.Append(",")
            vExamUnits.Append(vUnit)
          Next
          vList("ExamUnits") = vExamUnits.ToString
          vList("ExamUnitId") = mvTopLevelUnitId
          vList("ExamSessionId") = mvExamSessionId
          vList("Quantity") = "1"

          AddDefaultParameters(vList)
          'vList("Amount") = GetTextBoxText("TotalAmount")
          'vList("Notes") = GetTextBoxText("Notes")
          vList("UserID") = vList("ContactNumber")
          Dim vSkipProcessing As Boolean
          Try
            vParamList = DataHelper.AddExamBooking(vList)
          Catch vEx As ThreadAbortException
            Throw vEx
          Catch vEx As CareException
            SetErrorLabel(vEx.Message)
            vSkipProcessing = True
          End Try
          If vSkipProcessing = False Then
            GoToSubmitPage(String.Format("&XBN={0}", vParamList("ExamBookingNumber")))
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
    If FindControl("WarningMessage1") IsNot Nothing Then FindControl("WarningMessage1").Visible = False
    If FindControl("WarningMessage2") IsNot Nothing Then FindControl("WarningMessage2").Visible = False
    If FindControl("WarningMessage3") IsNot Nothing Then FindControl("WarningMessage3").Visible = False
    If QueryStringHasValue("XN") Then
      mvExamUnitId = IntegerValue(Request.QueryString("XN"))
    Else
      If InitialParameters.ContainsKey("ExamUnitId") Then mvExamUnitId = IntegerValue(InitialParameters("ExamUnitId").ToString)
    End If
    If QueryStringHasValue("XS") Then
      mvExamSessionId = IntegerValue(Request.QueryString("XS"))
    Else
      If InitialParameters.ContainsKey("ExamSessionId") Then mvExamSessionId = IntegerValue(InitialParameters("ExamSessionId").ToString)
    End If
    If QueryStringHasValue("XC") Then
      mvExamCentreId = IntegerValue(Request.QueryString("XC"))
    Else
      If InitialParameters.ContainsKey("ExamCentreId") Then mvExamUnitId = IntegerValue(InitialParameters("ExamCentreId").ToString)
    End If
    mvUnitsToBook = New List(Of Integer)
    mvUnitsToBook.Add(mvExamUnitId)

    If mvExamUnitId > 0 AndAlso mvExamCentreId > 0 AndAlso mvExamSessionId > 0 Then
      Try
        'Now get the heirarchy of units for this session including booking and passed details
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ExamSessionId") = mvExamSessionId
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamSessions, vList))
        If vRow IsNot Nothing Then SetTextBoxText("ExamSessionDescription", vRow.Item("ExamSessionDescription").ToString)
        vList = New ParameterList(HttpContext.Current)
        vList("ExamCentreId") = mvExamCentreId
        vRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamCentres, vList))
        If vRow IsNot Nothing Then
          SetTextBoxText("ExamCentreDescription", vRow.Item("ExamCentreDescription").ToString)
          mvExamCentreCode = vRow.Item("ExamCentreCode").ToString
        End If
        vList = New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserContactNumber()
        If mvExamSessionId > 0 Then vList("ExamSessionId") = mvExamSessionId.ToString
        vList("AllowBookings") = "Y"
        Dim vExamTable As DataTable = GetDataTable(DataHelper.SelectExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamStudentBookingUnits, vList))
        'First find all the units to be booked and the top level unit
        Dim vFindUnit As Integer = mvExamUnitId
        Dim vParentUnit As Integer = 0
        Do
          Dim vUnitRows() As DataRow = vExamTable.Select("ExamUnitId = " & vFindUnit)
          If vUnitRows.Length <> 1 Then
            If FindControl("WarningMessage1") IsNot Nothing Then FindControl("WarningMessage1").Visible = True
            mvError = True
            Exit Sub
          End If
          If vFindUnit = mvExamUnitId Then SetTextBoxText("ExamUnitDescription", vUnitRows(0).Item("ExamUnitDescription").ToString)

          Dim vBooked As Boolean = vUnitRows(0).Item("Booked").ToString.StartsWith("Y")
          'Next check there is not already a booking in the same session
          If vBooked AndAlso vFindUnit = mvExamUnitId Then
            If FindControl("WarningMessage2") IsNot Nothing Then FindControl("WarningMessage2").Visible = True
            mvError = True
            Exit Sub
          End If
          'Now check the unit has not already been passed
          Dim vPassed As Boolean = vUnitRows(0).Item("Passed").ToString.StartsWith("Y")
          If vPassed AndAlso vFindUnit = mvExamUnitId Then
            If FindControl("WarningMessage3") IsNot Nothing Then FindControl("WarningMessage3").Visible = True
            mvError = True
            Exit Sub
          End If
          If vFindUnit <> mvExamUnitId Then
            If Not vBooked And Not vPassed Then mvUnitsToBook.Add(vFindUnit)
          End If
          vParentUnit = IntegerValue(vUnitRows(0).Item("ExamUnitId1").ToString)
          If vParentUnit > 0 Then
            vFindUnit = vParentUnit
          Else
            mvTopLevelUnitId = IntegerValue(vUnitRows(0).Item("ExamUnitId").ToString)
          End If
        Loop While vParentUnit > 0
        'Now we have the top level unit and a full list of the items to book

        vList = New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserContactNumber()
        vList("ExamCentreCode") = mvExamCentreCode
        Dim vExamUnits As New StringBuilder
        For Each vUnit As Integer In mvUnitsToBook
          If vExamUnits.Length > 0 Then vExamUnits.Append(",")
          vExamUnits.Append(vUnit)
        Next
        vList("ExamUnits") = vExamUnits.ToString
        Dim vReturn As ParameterList = DataHelper.CalculateExamBookingPrice(vList)
        SetTextBoxText("Amount", DoubleValue(vReturn("TotalBookingPrice").ToString).ToString("0.00"))
        mvError = False
      Catch ex As Exception
        ProcessError(ex)
        mvError = True
      End Try
    End If
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
End Class