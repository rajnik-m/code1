Namespace Access

  Partial Public Class EventPersonnel

    Private mvAppointmentDescription As String = ""

    Public Property AppointmentDescription() As String
      Get
        AppointmentDescription = mvAppointmentDescription
      End Get
      Set(ByVal Value As String)
        mvAppointmentDescription = Value
      End Set
    End Property

    Public Sub Reschedule(ByRef pNewStartDate As String, ByRef pNewStartTime As String, ByRef pNewEndDate As String, ByRef pNewEndTime As String)
      mvClassFields.Item(EventPersonnelFields.StartDate).Value = pNewStartDate
      mvClassFields.Item(EventPersonnelFields.StartTime).Value = pNewStartTime
      mvClassFields.Item(EventPersonnelFields.EndDate).Value = pNewEndDate
      mvClassFields.Item(EventPersonnelFields.EndTime).Value = pNewEndTime
    End Sub

  End Class

End Namespace

