Partial Public Class MaintainActivityGroup
  Inherits CareWebControl

  Private mvGroupTable As DataTable
  Private mvActivityTable As DataTable

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctMaintainActivityGroup, tblDataEntry)

      Dim vList As New ParameterList(HttpContext.Current)
      vList("UsageCode") = "B"
      vList("ContactGroup") = "CON"
      vList("ActivityGroup") = InitialParameters("ActivityGroup")
      mvGroupTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtActivityDataSheet, vList)
      If mvGroupTable Is Nothing Then Throw New CareException("Activity Data sheet has no items to be entered")
      If Not InWebPageDesigner() Then
        mvActivityTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCategories, UserOrNewContactNumber)
      End If
      Dim vHTMLRow As HtmlTableRow
      Dim vHTMLCell As HtmlTableCell
      Dim vDDL As DropDownList
      Dim vActivityList As New ParameterList(HttpContext.Current)
      Dim vLastActivity As String = ""

      For Each vRow As DataRow In mvGroupTable.Rows
        If vLastActivity <> vRow("ActivityCode").ToString Then
          vLastActivity = vRow("ActivityCode").ToString
          vHTMLRow = New HtmlTableRow
          vHTMLCell = New HtmlTableCell
          vHTMLCell.InnerHtml = vRow("ActivityDesc").ToString
          vHTMLCell.Attributes("Class") = "DataEntryLabel"
          vHTMLRow.Cells.Add(vHTMLCell)
          vHTMLCell = New HtmlTableCell
          vDDL = New DropDownList
          vDDL.CssClass = "DataEntryItem"
          vDDL.Width = New Unit(200)
          vDDL.DataTextField = "ActivityValueDesc"
          vDDL.DataValueField = "ActivityValue"
          vActivityList("Activity") = vRow("ActivityCode")
          DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtActivityValues, vDDL, True, vActivityList)
          If Not IsPostBack And Not mvActivityTable Is Nothing Then
            For Each vActivityRow As DataRow In mvActivityTable.Rows
              If vActivityRow("ActivityCode").ToString = vRow("ActivityCode").ToString Then
                SelectListItem(vDDL, vActivityRow("ActivityValueCode").ToString)
                Exit For
              End If
            Next
          End If
          vHTMLCell.Controls.Add(vDDL)
          vHTMLRow.Cells.Add(vHTMLCell)
          tblDataEntry.Rows.Insert(tblDataEntry.Rows.Count - 1, vHTMLRow)
        End If
      Next
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vList As New ParameterList(HttpContext.Current)
        Dim vUpdateList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserOrNewContactNumber()
        vUpdateList("OldContactNumber") = UserOrNewContactNumber()
        vList("Source") = DefaultParameters("Source")
        Dim vLastActivity As String = ""
        Dim vFound As Boolean
        Dim vIndex As Integer
        Dim vDDL As DropDownList

        For Each vRow As DataRow In mvGroupTable.Rows
          If vLastActivity <> vRow("ActivityCode").ToString Then
            vLastActivity = vRow("ActivityCode").ToString
            vList("Activity") = vRow("ActivityCode")
            vDDL = TryCast(tblDataEntry.Rows(vIndex).Cells(1).Controls(0), DropDownList)
            If vDDL IsNot Nothing AndAlso vDDL.SelectedValue.Length > 0 Then
              vFound = False
              If Not mvActivityTable Is Nothing Then
                For Each vActivityRow As DataRow In mvActivityTable.Rows
                  If vActivityRow("ActivityCode").ToString = vRow("ActivityCode").ToString Then
                    vUpdateList("OldActivity") = vActivityRow("ActivityCode")
                    vUpdateList("OldActivityValue") = vActivityRow("ActivityValueCode")
                    vUpdateList("OldSource") = vActivityRow("SourceCode")
                    vUpdateList("OldValidFrom") = vActivityRow("ValidFrom")
                    vUpdateList("OldValidTo") = vActivityRow("ValidTo")
                    vUpdateList("Activity") = vActivityRow("ActivityCode")
                    vUpdateList("ActivityValue") = vDDL.SelectedValue
                    DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vUpdateList)
                    vFound = True
                    Exit For
                  End If
                Next
              End If
              If Not vFound Then
                vList("ActivityValue") = vDDL.SelectedValue
                DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctActivities, vList)
              End If
            End If
            vIndex += 1
          End If
        Next
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
End Class