Partial Public Class AddEvent
  Inherits CareWebControl

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    InitialiseControls(CareNetServices.WebControlTypes.wctAddEvent, tblDataEntry)
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vList As New ParameterList(HttpContext.Current)
    AddOptionalTextBoxValue(vList, "EventDesc")
    AddOptionalTextBoxValue(vList, "StartDate")
    AddOptionalTextBoxValue(vList, "StartTime")
    AddOptionalTextBoxValue(vList, "EndDate")
    AddOptionalTextBoxValue(vList, "EndTime")

    'Topic and Sub Topic

    vList("External") = "N"
    vList("Booking") = "N"
    vList("MultiSession") = "N"
    vList("FreeOfCharge") = "Y"
    vList("CandidateNumberingMethod") = "I"
    vList("FirstCandidateNumber") = 1
    vList("CandidateNumberBlockSize") = "1000"
    vList("EligibilityCheckRequired") = "N"
    vList("ChargeForWaiting") = "N"
    vList("WaitingListControlMethod") = "A"
    vList("EventGroup") = "EVE"
    vList("BookingsClose") = vList("StartDate")

    vList("Venue") = "LON"
    vList("Subject") = "FR"
    vList("SkillLevel") = "ALL"
    vList("MinimumAttendees") = "1"
    vList("MaximumOnWaitingList") = "99"
    vList("MaximumAttendees") = "100"
    vList("TargetAttendees") = "100"
    Dim vReturnList As ParameterList = DataHelper.AddEvent(vList)

    Dim vAddOptionList As New ParameterList(HttpContext.Current)
    vAddOptionList("EventNumber") = vReturnList("EventNumber")
    vAddOptionList("OptionDesc") = "Standard"
    vAddOptionList("PickSessions") = "N"
    vAddOptionList("NumberOfSessions") = "1"
    vAddOptionList("IssueEventResources") = "Y"
    vAddOptionList("DeductFromEvent") = "Y"
    vAddOptionList("MaximumBookings") = "1"
    vAddOptionList("Product") = "1DS"             '
    vAddOptionList("Rate") = "STD"
    Dim vReturnOptionList As ParameterList = DataHelper.AddEventBookingOption(vAddOptionList)

    Dim vUpdateList As New ParameterList(HttpContext.Current)
    vUpdateList("EventNumber") = vReturnList("EventNumber")
    vUpdateList("Booking") = "Y"
    vUpdateList("WebPublish") = "Y"
    Dim vUpdateReturnList As ParameterList = DataHelper.UpdateEvent(vUpdateList)

  End Sub

End Class