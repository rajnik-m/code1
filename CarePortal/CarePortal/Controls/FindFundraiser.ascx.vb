Partial Public Class FindFundraiser
  Inherits CareWebControl

  Dim mvCV As CustomValidator

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctFindFundraiser, tblDataEntry, "")
    If InWebPageDesigner() Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      vList("DocumentColumns") = "Y"
      FindFundraisers(vList, False)
    Else
      If Not IsPostBack Then SetResults(False)
    End If
  End Sub

  Protected Overrides Sub AddCustomValidator(ByVal pHTMLTable As HtmlTable)
    Dim vControl As Control = FindControlByName(tblDataEntry, "Search")
    If vControl IsNot Nothing Then
      AddCustomValidator(DirectCast(vControl.Parent, HtmlTableCell), "1", "Please enter some data to search for")
    End If
  End Sub

  Public Overrides Sub ServerValidate(ByVal sender As Object, ByVal args As ServerValidateEventArgs)
    args.IsValid = GetTextBoxText("Surname").Length > 0 OrElse GetTextBoxText("Town").Length > 0 OrElse GetTextBoxText("FundraisingDescription").Length > 0 OrElse GetTextBoxText("EventDesc").Length > 0
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        If DirectCast(sender, Button).ID = "SearchAgain" Then
          SetResults(False)
        Else
          Dim vList As New ParameterList(HttpContext.Current)
          AddOptionalTextBoxValueWithWildCard(vList, "Forenames")
          AddOptionalTextBoxValueWithWildCard(vList, "Surname")
          AddOptionalTextBoxValueWithWildCard(vList, "Town")
          AddOptionalTextBoxValueWithWildCard(vList, "FundraisingDescription")
          AddOptionalTextBoxValueWithWildCard(vList, "EventDesc")
          AddOptionalTextBoxValue(vList, "TargetDate")
          AddOptionalDropDownValue(vList, "Venue")
          AddOptionalDropDownValue(vList, "Organiser")
          AddOptionalDropDownValue(vList, "SkillLevel")
          AddOptionalDropDownValue(vList, "Topic")
          AddOptionalDropDownValue(vList, "EventGroup")
          AddOptionalDropDownValue(vList, "Branch")
          AddOptionalDropDownValue(vList, "DistributionCode")
          vList("SystemColumns") = "Y"
          FindFundraisers(vList, True)
        End If
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub

  Private Sub FindFundraisers(ByVal pList As ParameterList, ByVal pShowCount As Boolean)
    pList("SystemColumns") = "Y"
    Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftContactFundraisingEvents, pList)
    Dim vDGR As DataGrid = CType(FindControlByName(pnlResults, "Results"), DataGrid)
    If vDGR IsNot Nothing Then
      Dim vCount As Integer = DataHelper.FillGrid(vResult, vDGR, "")
      If pShowCount Then
        vDGR.Visible = vCount > 0
        SetResults(True)
        Dim vLabel As Label = CType(FindControlByName(pnlResults, "ResultMessage"), Label)
        If vLabel IsNot Nothing Then vLabel.Text = String.Format(vLabel.Text, vCount)
      End If
    End If
  End Sub

  Private Sub SetResults(ByVal pVisible As Boolean)
    pnlResults.Visible = pVisible
    pnlFinder.Visible = Not pVisible
  End Sub

End Class