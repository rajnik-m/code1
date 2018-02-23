Partial Public Class SubmitAll
  Inherits CareWebControl

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvCenterControl = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSubmitAll, tblDataEntry)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
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
                If vCareWebControl.ParentGroup = "RegisteredUser" OrElse vCareWebControl.ParentGroup = "SelectedContact" OrElse vCareWebControl.ParentGroup = "SelectedOrganisation" Then
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

  Public Overrides Sub ProcessSubmit()
    If Me.DependantControls IsNot Nothing AndAlso Me.DependantControls.Count > 0 AndAlso _
       Session("RegisteredUserName") IsNot Nothing AndAlso Session("RegisteredUserName").ToString.Length > 0 AndAlso _
       HttpContext.Current.User.Identity.IsAuthenticated Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = UserContactNumber()
      vList("AddressNumber") = UserAddressNumber()
      SubmitChildControls(vList)
    End If
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If Not IsPostBack Then
      If Session("RegisteredUserName") IsNot Nothing AndAlso Session("RegisteredUserName").ToString.Length > 0 AndAlso HttpContext.Current.User.Identity.IsAuthenticated Then
        Dim vTable As DataTable = Nothing
        Dim vGroupName As String = "RegisteredUser"
        For Each vCareWebControl As CareWebControl In mvPageCareControls
          If TypeOf vCareWebControl Is ICareChildWebControl Then                'If the control needs a parent 
            If Not DontClearChild Then vCareWebControl.ClearControls()
            If vCareWebControl.ParentGroup = "RegisteredUser" OrElse vCareWebControl.ParentGroup = "SelectedContact" OrElse vCareWebControl.ParentGroup = "SelectedOrganisation" Then
              vGroupName = vCareWebControl.ParentGroup
              If vTable Is Nothing Then vTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, GetContactNumberFromParentGroup(vCareWebControl.ParentGroup))
            End If
          End If
        Next
        If vTable IsNot Nothing Then
          GroupName = vGroupName
          ProcessContactSelection(vTable)
        End If
      End If
    End If
  End Sub
End Class