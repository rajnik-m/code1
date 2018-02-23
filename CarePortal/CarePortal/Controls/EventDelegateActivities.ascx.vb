Public Class EventDelegateActivities
  Inherits CareWebControl

  Dim mvEventNumber As Integer
  Dim mvContactNumber As String
  Dim mvBookingNumber As String
  Dim mvDelegateDataTable As DataTable

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctEventDelegateActivities, tblDataEntry)
    SetControlVisible("WarningMessage3", False)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vDelegateName As Label = CType(FindControlByName(Me, "WarningMessage1"), Label)
    Dim vDelegateLabel As Label = CType(FindControlByName(Me, "WarningMessage2"), Label)
    Try
      If Request.QueryString("BN") IsNot Nothing AndAlso Not String.IsNullOrEmpty(Request.QueryString("BN").ToString) Then
        mvBookingNumber = Request.QueryString("BN").ToString
      ElseIf InitialParameters.Contains("BookingNumber") Then
        mvBookingNumber = InitialParameters("BookingNumber").ToString
      End If
      If mvBookingNumber IsNot Nothing AndAlso mvBookingNumber.Trim.Length > 0 Then

        mvEventNumber = CInt(CInt(mvBookingNumber) / 10000)

        vList("EventNumber") = mvEventNumber
        mvDelegateDataTable = DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventInformation, vList)

        'If no activity group then just submit
        If String.IsNullOrEmpty(mvDelegateDataTable.Rows(0).Item("ActivityGroup").ToString) Then
          GoToSubmitPage()
        Else
          'Go get the activities for this group
          Dim vResultData As DataTable
          vList = New ParameterList(HttpContext.Current)
          vList("ActivityGroup") = mvDelegateDataTable.Rows(0).Item("ActivityGroup").ToString
          vResultData = DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtActivityFromActivityGroup, vList)

          'Check which items should be visible
          For Each vCareWebControl As CareWebControl In mvPageCareControls
            If vCareWebControl.HandlesActivities Then
              Dim vVisible As Boolean = False
              For Each vdr As DataRow In vResultData.Rows
                Dim vActivityDetail As String = vdr.Item("Activity").ToString
                If String.Equals(vCareWebControl.InitialParameters("Activity"), vActivityDetail) Or String.Equals(vCareWebControl.DefaultParameters("Activity"), vActivityDetail) Then
                  vVisible = True
                End If
              Next
              vCareWebControl.Visible = vVisible
            End If
            If Object.Equals(vCareWebControl, Me) Then
              vCareWebControl.GroupName = "DelegateActivities"
            Else
              vCareWebControl.ParentGroup = "DelegateActivities"
            End If
          Next

          'Get the delegates for this booking
          vList = New ParameterList(HttpContext.Current)
          vList("BookingNumber") = mvBookingNumber
          mvDelegateDataTable = DataHelper.GetEventDataTable(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates, vList)
          If mvDelegateDataTable Is Nothing Then
            vDelegateName.Text = String.Format("Invalid Booking number {0} or No delegate associated with this Booking number", mvBookingNumber)
            For Each vCareWebControl As CareWebControl In mvPageCareControls
              If Not Object.Equals(vCareWebControl, Me) Then vCareWebControl.Visible = False
            Next
          Else
            'Get the delegate number
            Dim vTable As New DataTable
            If Request.QueryString("DN") IsNot Nothing AndAlso Not String.IsNullOrEmpty(Request.QueryString("DN")) Then
              EventDelegateNumber = Request.QueryString("DN").ToString
              mvContactNumber = mvDelegateDataTable.Select(String.Format("EventDelegateNumber = {0}", EventDelegateNumber))(0).Item("ContactNumber").ToString()
            Else
              EventDelegateNumber = mvDelegateDataTable.Rows(0).Item("EventDelegateNumber").ToString
              mvContactNumber = mvDelegateDataTable.Rows(0).Item("ContactNumber").ToString
            End If
            Dim vCurrentDelegateNumber As Integer

            For vCounter As Integer = 0 To mvDelegateDataTable.Rows.Count - 1
              If String.Equals(mvDelegateDataTable.Rows(vCounter).Item("EventDelegateNumber").ToString, EventDelegateNumber) Then
                vCurrentDelegateNumber = vCounter + 1
              End If
            Next vCounter
            'Set the delegate
            If vDelegateLabel IsNot Nothing Then vDelegateLabel.Text = String.Format(vDelegateLabel.Text, vCurrentDelegateNumber, mvDelegateDataTable.Rows.Count)
            If Not IsPostBack Then
              vList = New ParameterList(HttpContext.Current)
              vList("ContactNumber") = mvContactNumber
              vTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList)
              If vDelegateName IsNot Nothing Then vDelegateName.Text = vTable.Rows(0).Item("ContactName").ToString
              For Each vCareWebControl As CareWebControl In mvPageCareControls
                If vCareWebControl.Visible = True Then
                  'I assume this will set the default
                  vCareWebControl.ProcessContactSelection(vTable)
                  vCareWebControl.DontClearChild = True
                End If
              Next
            End If
          End If
        End If
      Else 
        If Not InWebPageDesigner() Then
          SetControlVisible("WarningMessage3", True)
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vNextDelegate As Boolean = False

    vList("EventDelegateNumber") = EventDelegateNumber
    If DependantControls IsNot Nothing Then
      For Each vCareWebControl As ICareChildWebControl In DependantControls
        If DirectCast(vCareWebControl, CareWebControl).Visible Then vCareWebControl.SubmitChild(vList)
      Next
    End If
    For Each vDr As DataRow In mvDelegateDataTable.Rows
      If vNextDelegate Then
        ProcessRedirect("Default.aspx?pn=" & Request.QueryString("pn").ToString & "&DN=" & vDr.Item("EventDelegateNumber").ToString & "&BN=" & mvBookingNumber)
      End If
      If String.Equals(vDr.Item("EventDelegateNumber").ToString, EventDelegateNumber) Then
        If Not Object.Equals(vDr, mvDelegateDataTable.Rows(mvDelegateDataTable.Rows.Count - 1)) Then
          vNextDelegate = True
        End If
      End If
    Next
    If Not vNextDelegate Then
      GoToSubmitPage()
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        For Each vCareWebControl As CareWebControl In mvPageCareControls
          vCareWebControl.DependantControls = New List(Of ICareChildWebControl)
        Next
        For Each vCareWebControl As CareWebControl In mvPageCareControls
          If TypeOf vCareWebControl Is ICareChildWebControl Then                'If the control needs a parent 
            If vCareWebControl.ParentGroup.Length > 0 Then                      'Check that it has one
              Dim vParentControl As CareWebControl = FindGroupControl(vCareWebControl.ParentGroup)
              If vParentControl Is Nothing Then                                 'Also make sure the parent exists
                If vCareWebControl.ParentGroup = "registereduser" Then
                  Me.DependantControls.Add(CType(vCareWebControl, ICareChildWebControl))
                Else
                  Throw New CareException(String.Format("Cannot find Web Module with Group Name {0} defined for Web Module {1}", vCareWebControl.ParentGroup, WebPageItemName))
                End If
              Else
                vParentControl.DependantControls.Add(CType(vCareWebControl, ICareChildWebControl))
              End If
            Else
              Throw New CareException(String.Format("No Parent Group Name is defined for Web Module {0}", vCareWebControl.WebPageItemName))
            End If
          End If
        Next
        'OK we are all valid so we need to submit all the items - do the parents first and then children of those parents
        For Each vCareWebControl As CareWebControl In mvPageCareControls
          If Not vCareWebControl.NeedsParent Then                               'If the control does not needs a parent 
            vCareWebControl.ProcessSubmit()
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