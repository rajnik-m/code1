Partial Public Class ExamSelection
  Inherits CareWebControl

  Private mvHyperLinkText As String = ""
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectExams, tblDataEntry)
      If InitialParameters.ContainsKey("HyperlinkText") Then mvHyperLinkText = InitialParameters("HyperlinkText").ToString
      If Me.FindControl("SearchExam") IsNot Nothing Then
        CType(Me.FindControl("SearchExam"), TextBox).MaxLength = 100
        If Request.QueryString("Exam") IsNot Nothing Then
          CType(Me.FindControl("SearchExam"), TextBox).Text = Request.QueryString("Exam")
        End If
      End If
      If Not IsPostBack Then FindExams()
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

  Private Sub FindExams()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vQueryString As String = ""
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    If GetTextBoxText("SearchExam").Length > 0 Then
      vList("SearchExam") = "*" & GetTextBoxText("SearchExam") & "*"
      vQueryString = "&Exam=" & GetTextBoxText("SearchExam")
    End If
    Dim vBaseList As BaseDataList = TryCast(Me.FindControl("ExamData"), BaseDataList)
    Dim vEditPageNumber As Integer = 0
    Dim vBookingPage As String = ""
    If InitialParameters.ContainsKey("Subject") Then vList("Subject") = InitialParameters("Subject").ToString
    If Request.QueryString("SU") IsNot Nothing AndAlso Request.QueryString("SU").Length > 0 Then vList("Subject") = Request.QueryString("SU")
    If InitialParameters.ContainsKey("SkillLevel") Then vList("SkillLevel") = InitialParameters("SkillLevel").ToString
    If Request.QueryString("SL") IsNot Nothing AndAlso Request.QueryString("SL").Length > 0 Then vList("SkillLevel") = Request.QueryString("SL")
    If InitialParameters.ContainsKey("ExamSessionCode") Then vList("ExamSessionCode") = InitialParameters("ExamSessionCode").ToString
    If Request.QueryString("XSC") IsNot Nothing AndAlso Request.QueryString("XSC").Length > 0 Then vList("ExamSessionCode") = Request.QueryString("XSC")
    If InitialParameters.ContainsKey("ExamCentreCode") Then vList("ExamCentreCode") = InitialParameters("ExamCentreCode").ToString
    If Request.QueryString("XCC") IsNot Nothing AndAlso Request.QueryString("XCC").Length > 0 Then vList("ExamCentreCode") = Request.QueryString("XCC")
    If Request.QueryString("XN") IsNot Nothing AndAlso Request.QueryString("XN").Length > 0 Then vList("ExamUnitId") = Request.QueryString("XN")
    If UserContactNumber() > 0 Then vList("ContactNumber") = UserContactNumber()
    If InitialParameters.ContainsKey("BookingPage") Then vEditPageNumber = IntegerValue(InitialParameters("BookingPage").ToString)
    If vBaseList IsNot Nothing Then
      Dim vRowCount As Integer = DataHelper.GetPagedFinderData(CareNetServices.XMLDataFinderTypes.xdftWebExams, vBaseList, Request, plcHolder, vList, IntegerValue(InitialParameters("ItemsPerPage").ToString), vEditPageNumber, False, vQueryString)
      'Only for display grids. Data list select columns will be handled separately
      If vRowCount > 0 Then
        If (Not InitialParameters.ContainsKey("DisplayFormat")) OrElse InitialParameters("DisplayFormat").ToString = "0" Then
          Dim vDGR As DataGrid = CType(vBaseList, DataGrid)
          If InitialParameters.Contains("BookingPage") Then vBookingPage = InitialParameters("BookingPage").ToString
          Dim vSelectPos As Integer
          Dim vColumn As New BoundColumn()
          'Book Exam column.
          vColumn.HeaderText = ""
          vDGR.Columns.AddAt(0, vColumn)
          vDGR.DataBind()
          For vCount As Integer = 0 To vDGR.Columns.Count - 1
            Dim vBoundColumn As BoundColumn = TryCast(vDGR.Columns(vCount), BoundColumn)
            If vBoundColumn IsNot Nothing AndAlso vBoundColumn.HeaderText = "" Then
              vSelectPos = vCount
            End If
          Next
          Dim vExamUnitIdColumn As Integer = GetDataGridItemIndex(vDGR, "ExamUnitId")
          Dim vExamSessionIdColumn As Integer = GetDataGridItemIndex(vDGR, "ExamSessionId")
          Dim vExamCentreIdColumn As Integer = GetDataGridItemIndex(vDGR, "ExamCentreId")
          If vBookingPage.Length = 0 AndAlso vSelectPos >= 0 Then
            vDGR.Columns(vSelectPos).Visible = False
          Else
            For vRow As Integer = 0 To vDGR.Items.Count - 1
              Dim vExamSessionId As Integer = IntegerValue(vDGR.Items(vRow).Cells(vExamSessionIdColumn).Text)
              Dim vExamCentreId As Integer = IntegerValue(vDGR.Items(vRow).Cells(vExamCentreIdColumn).Text)
              Dim vExamUnitId As Integer = IntegerValue(vDGR.Items(vRow).Cells(vExamUnitIdColumn).Text)
              Dim vParameters As String = String.Format("&XS={0}&XC={1}&XN={2}", vExamSessionId, vExamCentreId, vExamUnitId)
              If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
                vDGR.Items(vRow).Cells(0).Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & """ onclick=""location.href='Default.aspx?pn={0}{1}'"">", vBookingPage, vParameters)
              Else
                vDGR.Items(vRow).Cells(vSelectPos).Text = "<a href='default.aspx?pn=" & vBookingPage & vParameters & "'>" & mvHyperLinkText & "</a>"
              End If
            Next
          End If
          'For displying Exam image
          For vCount As Integer = 1 To vDGR.Columns.Count - 1
            Dim vBoundColumn As BoundColumn = DirectCast(vDGR.Columns(vCount), BoundColumn)
            Dim vPath As String = ""
            If vBoundColumn.DataField = "ExamImage" Then
              For vRow As Integer = 0 To vDGR.Items.Count - 1
                vPath = "Images/Exams/" & vDGR.Items(vRow).Cells(vCount).Text
                'Call the GetImage which checks whether Image is available or not.
                vDGR.Items(vRow).Cells(vCount).Text = GetImage(vPath, vDGR.Items(vRow).Cells(vCount).Text, "Images/Exams/Default.png", "ExamImage")
              Next
              Exit For
            End If
          Next
          vBaseList.Visible = True
        End If
        DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
      Else
        If GetTextBoxText("SearchExam").Length > 0 Then
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        Else
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
          DirectCast(Me.FindControl("SearchExam"), TextBox).Visible = False
          DirectCast(Me.FindControl("SearchExam"), TextBox).Parent.Parent.Visible = False
          DirectCast(Me.FindControl("Search"), Button).Visible = False
        End If
        DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        vBaseList.Visible = False
      End If
    End If
  End Sub

  Public Overrides Sub HandleDataListItemDataBound(ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
    If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
      Dim vBookingPage As String = InitialParameters.OptionalValue("BookingPage").ToString
      Dim vDrv As DataRowView = CType(e.Item.DataItem, DataRowView)

      'Add the book Exam link at the end
      If vBookingPage.Length > 0 Then
        Dim vCount As Integer = e.Item.Controls.Count
        Dim vBookExamLink As New Literal

        Dim vExamSessionId As Integer = IntegerValue(vDrv.Row("ExamSessionId").ToString)
        Dim vExamCentreId As Integer = IntegerValue(vDrv.Row("ExamCentreId").ToString)
        Dim vExamUnitId As Integer = IntegerValue(vDrv.Row("ExamUnitId").ToString)
        Dim vParameters As String = String.Format("&XS={0}&XC={1}&XN={2}", vExamSessionId, vExamCentreId, vExamUnitId)


        If InitialParameters.OptionalValue("HyperlinkFormat") = "B" Then
          vBookExamLink.Text = String.Format("<input type=""button"" class=""Button"" value=""" & mvHyperLinkText & """ onclick=""location.href='Default.aspx?pn={0}{1}'"">", vBookingPage, vParameters)
        Else
          vBookExamLink.Text = String.Format("<a href='Default.aspx?pn={0}{1}'>" & mvHyperLinkText & "</a>", vBookingPage, vParameters)
        End If
        If vCount > 0 Then e.Item.Controls(vCount - 1).Parent.Controls.Add(vBookExamLink)
      End If
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        FindExams()
      Catch vEx As ThreadAbortException
        Throw vEx
      End Try
    End If
  End Sub

End Class
