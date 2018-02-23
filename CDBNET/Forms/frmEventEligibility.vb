Public Class frmEventEligibility

  Private mvContactInfo As ContactInfo
  Private mvEventInfo As CareEventInfo
  Private mvSource As String

  Public Sub New(ByVal pContactInfo As ContactInfo, ByVal pEventInfo As CareEventInfo, ByVal pSource As String)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pContactInfo, pEventInfo, pSource)
  End Sub

  Private Sub InitialiseControls(ByVal pContactInfo As ContactInfo, ByVal pEventInfo As CareEventInfo, ByVal pSource As String)
    SetControlTheme()
    Me.Text = ControlText.FrmEventEligibility
    Me.Icon = My.Resources.Events
    mvContactInfo = pContactInfo
    mvEventInfo = pEventInfo
    mvSource = pSource
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optEventEligibility))
    epl.SetValue("Delegate", mvContactInfo.ContactName.ToString)
    epl.SetValue("Eligibility", MultiLine(mvEventInfo.EligibilityText))
  End Sub

  Private Sub cmdApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApprove.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub cmdDefer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDefer.Click
    AddActivity(False)
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub cmdReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReject.Click
    AddActivity(True)
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
    Me.Close()
    FormHelper.ShowCardIndex(CareServices.XMLContactDataSelectionTypes.xcdtContactInformation, mvContactInfo.ContactNumber, False)
  End Sub

  Private Sub AddActivity(ByVal pReject As Boolean)
    Try
      Dim vList As New ParameterList(True)
      vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
      vList("Source") = mvSource
      vList("ValidFrom") = AppValues.TodaysDate
      vList("ValidTo") = mvEventInfo.EndDate.ToString(AppValues.DateFormat)
      If pReject Then
        vList("Activity") = mvEventInfo.RejectedBookingActivity
        vList("ActivityValue") = mvEventInfo.RejectedBookingActivityValue
      Else
        vList("Activity") = mvEventInfo.DeferredBookingActivity
        vList("ActivityValue") = mvEventInfo.DeferredBookingActivityValue
      End If
      DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActivities, vList)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
End Class