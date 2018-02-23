Public Class ShowEventBookings
  Inherits CareWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctShowEventBookings, tblDataEntry)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      vList("ContactNumber") = UserContactNumber().ToString
      Dim vDataGrid As DataGrid = TryCast(Me.FindControl("EventBookingData"), DataGrid)
      If vDataGrid IsNot Nothing Then
        Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebEventBookings, vList)
        DataHelper.FillGrid(vResult, vDataGrid)
        If vDataGrid.Items.Count > 0 Then
          Dim vDelegateSelectionPage As String = ""
          If InitialParameters.Contains("DelegateSelectionPage") Then vDelegateSelectionPage = InitialParameters("DelegateSelectionPage").ToString
          Dim vColumn As New TemplateColumn()
          vColumn.HeaderText = ""
          vDataGrid.Columns.AddAt(0, vColumn)
          vDataGrid.DataBind()
          Dim vBookingPos As Integer = -1
          Dim vSelectPos As Integer = -1
          Dim vContactPos As Integer = -1
          Dim vBookingsClosePos As Integer
          For vCount As Integer = 0 To vDataGrid.Columns.Count - 1
            Dim vBoundColumn As TemplateColumn = DirectCast(vDataGrid.Columns(vCount), TemplateColumn)
            If vBoundColumn.HeaderText = "" Then
              vSelectPos = vCount
            ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "BookingNumber" Then
              vBookingPos = vCount
            ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "BookerContactNumber" Then
              vContactPos = vCount
            ElseIf DirectCast(vBoundColumn.ItemTemplate, DisplayTemplate).DataItem = "BookingsClose" Then
              vBookingsClosePos = vCount
            End If
          Next
          If vDelegateSelectionPage.Length = 0 AndAlso vSelectPos >= 0 Then
            vDataGrid.Columns(vSelectPos).Visible = False
          ElseIf vBookingPos >= 0 AndAlso vContactPos >= 0 Then
            For vRow As Integer = 0 To vDataGrid.Items.Count - 1
              Dim vAllowEdit As Boolean = True
              If vBookingsClosePos >= 0 Then
                Dim vBookingsClose As String = DirectCast(vDataGrid.Items(vRow).Cells(vBookingsClosePos).Controls(0), ITextControl).Text
                If DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.portal_bookings_close_lock) = True AndAlso IsDate(vBookingsClose) AndAlso CDate(vBookingsClose) < Date.Today Then
                  vAllowEdit = False
                End If
              End If
              If vAllowEdit AndAlso UserContactNumber() = IntegerValue(DirectCast(vDataGrid.Items(vRow).Cells(vContactPos).Controls(0), ITextControl).Text) Then
                vDataGrid.Items(vRow).Cells(vSelectPos).Text = "<a href='default.aspx?pn=" & vDelegateSelectionPage & "&BN=" & DirectCast(vDataGrid.Items(vRow).Cells(vBookingPos).Controls(0), ITextControl).Text & "'>Update Delegates</a>"
              End If
            Next
          End If
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
        Else
          vDataGrid.Visible = False
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = True
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

End Class