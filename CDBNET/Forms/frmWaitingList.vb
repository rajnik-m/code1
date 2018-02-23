Public Class frmWaitingList

  Private mvEventInfo As CareEventInfo

  Public Sub New(ByVal pEventInfo As CareEventInfo)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pEventInfo)
  End Sub

  Private Sub InitialiseControls(ByVal pEventInfo As CareEventInfo)
    SetControlTheme()
    mvEventInfo = pEventInfo
    Me.Text = GetInformationMessage(ControlText.frmWaitingList, mvEventInfo.EventDescription)
    Me.lblBookings.Text = ControlText.lblBookings
    Me.lblDelegates.Text = ControlText.lblDelegates
    dgrBookings.MultipleSelect = True
    dgrDelegates.MultipleSelect = True
    PopulateLists()
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      ProcessSelections()
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub ProcessSelections()
    Try
      Dim vList As ArrayListEx = dgrBookings.GetSelectedRowIntegers("BookingNumber")
      Dim vNoTransfer As New ArrayListEx
      Dim vResultList As ParameterList
      For Each vBooking As Integer In vList
        vResultList = DataHelper.TransferWaitingListBooking(vBooking)
        If Not vResultList.ContainsKey("BookingNumber") Then
          vNoTransfer.Add(vBooking)
        End If
      Next
      If vNoTransfer.Count > 0 Then ShowInformationMessage(InformationMessages.ImCannotTransfer, vNoTransfer.CSList)
    Catch vCareEx As CareException
      If vCareEx.ErrorNumber = CareException.ErrorNumbers.enCCAuthorisationFailed OrElse vCareEx.ErrorNumber = CareException.ErrorNumbers.enCardAuthorisationUnexpectedTimeout Then
        ShowInformationMessage(vCareEx.Message)
      Else
        Throw vCareEx
      End If
    End Try
  End Sub

  Private Sub PopulateLists()
    dgrBookings.Populate(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEVWaitingBookings, mvEventInfo.EventNumber))
    dgrDelegates.Populate(DataHelper.GetEventData(CareServices.XMLEventDataSelectionTypes.xedtEVWaitingDelegates, mvEventInfo.EventNumber))
    dgrBookings.SynchSelections(dgrDelegates, "BookingNumber")
  End Sub

  Private Sub cmdApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdApply.Click
    Try
      ProcessSelections()
      PopulateLists()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub dgrBookings_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgrBookings.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub dgrBookings_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgrBookings.RowSelected
    dgrBookings.SynchSelections(dgrDelegates, "BookingNumber")
  End Sub

  Private Sub dgrDelegates_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgrDelegates.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub dgrDelegates_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgrDelegates.RowSelected
    dgrDelegates.SynchSelections(dgrBookings, "BookingNumber")
    dgrBookings.SynchSelections(dgrDelegates, "BookingNumber")
  End Sub

End Class