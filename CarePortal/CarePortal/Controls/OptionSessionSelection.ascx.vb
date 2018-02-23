Public Class OptionSessionSelection
  Inherits CareWebControl

  Private mvNumberOfSession As Integer
  Private mvEventNumber As String = ""
  Private mvOptionNumber As String = ""
  Private mvRate As String = ""

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectOptionSessions, tblDataEntry)
      SetControlVisible("WarningMessage", False)
      SetControlVisible("WarningMessage1", False)
      SetControlVisible("WarningMessage2", False)
      Dim vList As New ParameterList(HttpContext.Current)
      If (Request.QueryString("EN") IsNot Nothing AndAlso Request.QueryString("OP") IsNot Nothing) AndAlso (Request.QueryString("EN").Length > 0 AndAlso Request.QueryString("OP").Length > 0) Then
        mvEventNumber = Request.QueryString("EN")
        mvOptionNumber = Request.QueryString("OP")
      Else
        If InitialParameters.ContainsKey("EventNumber") Then mvEventNumber = InitialParameters("EventNumber").ToString
        If InitialParameters.ContainsKey("OptionNumber") Then mvOptionNumber = InitialParameters("OptionNumber").ToString
      End If
      If mvEventNumber.Length > 0 AndAlso mvOptionNumber.Length > 0 Then
        vList("EventNumber") = mvEventNumber
        vList("OptionNumber") = mvOptionNumber
        If UserContactNumber() > 0 Then vList("ContactNumber") = UserContactNumber()
        Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebBookingOptions, vList)
        Dim vDataTable As DataTable
        vDataTable = GetDataTable(vResult)
        If vDataTable IsNot Nothing Then
          If vDataTable.Columns.Contains("PickSessions") AndAlso vDataTable.Rows(0).Item("PickSessions").ToString = "Y" Then
            mvNumberOfSession = IntegerValue(vDataTable.Rows(0).Item("NumberOfSession").ToString)
            mvRate = vDataTable.Rows(0).Item("Rate").ToString
          Else
            If Not InWebPageDesigner() Then
              Dim vSubmitParams As New StringBuilder
              With vSubmitParams
                .Append("&EN=")
                .Append(Request.QueryString("EN"))
                .Append("&OP=")
                .Append(Request.QueryString("OP"))
                .Append("&RA=")
                .Append(Request.QueryString("RA"))
              End With
              GoToSubmitPage(vSubmitParams.ToString)
            End If
          End If
          vList.Remove("EventNumber")
          FillPickSessionGrid(vList)
        Else
          ShowMessageOnlyFromLabel("WarningMessage2", "The selected Booking Option is not available")
        End If
      Else
        vList("FromWPD") = "Y"
        FillPickSessionGrid(vList)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vDataGrid As DataGrid = TryCast(Me.FindControl("PickSessionData"), DataGrid)
        If vDataGrid IsNot Nothing AndAlso vDataGrid.Items.Count > 0 Then
          Dim vSessionPos As Integer
          Dim vSessionCount As Integer = 0
          Dim vSessionList As New StringBuilder
          For vCount As Integer = 1 To vDataGrid.Columns.Count - 1
            Dim vBoundColumn As TemplateColumn = TryCast(vDataGrid.Columns(vCount), TemplateColumn)
            If vBoundColumn IsNot Nothing AndAlso
               DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "SessionNumber" Then
              vSessionPos = vCount
            End If
          Next
          For vRow As Integer = 0 To vDataGrid.Items.Count - 1
            If CType(vDataGrid.Items(vRow).Cells(0).Controls(0), CheckBox).Checked Then
              vSessionCount = vSessionCount + 1
              vSessionList.Append(",")
              vSessionList.Append(DirectCast(vDataGrid.Items(vRow).Cells(vSessionPos).Controls(0), ITextControl).Text)
            End If
          Next
          If vSessionList.Length > 1 Then vSessionList.Remove(0, 1)
          Dim vBookingPage As String = ""
          If vSessionCount = mvNumberOfSession Then
            If Not InWebPageDesigner() Then
              Dim vSubmitParams As New StringBuilder
              With vSubmitParams
                .Append("&EN=")
                .Append(mvEventNumber)
                .Append("&OP=")
                .Append(mvOptionNumber)
                .Append("&RA=")
                .Append(mvRate)
                .Append("&SL=")
                .Append(vSessionList.ToString)
              End With
              GoToSubmitPage(vSubmitParams.ToString)
            End If
          Else
            Dim vMessage As String = GetLabelText("WarningMessage")
            If String.IsNullOrEmpty(vMessage) Then vMessage = "{0} Sessions must be selected"
            SetLabelText("WarningMessage", String.Format(vMessage, mvNumberOfSession))
            SetControlVisible("WarningMessage", True)
          End If
        End If
      Catch vEx As ThreadAbortException
        Throw vEx
      End Try
    End If
  End Sub

  Private Sub FillPickSessionGrid(ByVal pList As ParameterList)
    Dim vDataGrid As DataGrid = TryCast(Me.FindControl("PickSessionData"), DataGrid)
    pList("SystemColumns") = "Y"
    pList("WebPageItemNumber") = Me.WebPageItemNumber
    If vDataGrid IsNot Nothing Then
      Dim vEventData As String = DataHelper.SelectEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingOptionSessions, pList)
      Dim vRowCount As Integer = DataHelper.FillGrid(vEventData, vDataGrid, "PlacesAvailable > 0")
      If vRowCount <= 0 Then
        ShowMessageOnlyFromLabel("WarningMessage1", "No Sessions are available for the selected Booking Option")
      End If
    End If
  End Sub
End Class



