Public Class frmActionSchedule
  Inherits MaintenanceParentForm

  Private mvSelectionSetNumber As Integer
  Private mvDisplayType As DisplayTypes = DisplayTypes.Quarter
  Private mvFirstDate As Date = Today
  Private mvPreviousTab As TabPage
  Private mvBrowserMenu As New BrowserMenu(Nothing)

  Private Enum DisplayTypes
    StepPevious
    Day
    Week
    Month
    Quarter
    Year
    StepNext
  End Enum

  Private Const COL_CONTACT_NUMBER As Integer = 0
  Private Const COL_CONTACT_NAME As Integer = 1
  Private Const COL_REQUEST_STAGE As Integer = 2
  Private Const COL_FIRST_DATE As Integer = 3

  Public Sub New(ByVal pSelectionSetNo As Integer, ByVal pDesc As String)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pSelectionSetNo, pDesc)
  End Sub

  Private Structure Appointment
    Dim StartDate As Date
    Dim EndDate As Date
    Dim RecordType As String
    Dim UniqueID As Integer
  End Structure

  Private Sub InitialiseControls(ByVal pSelectionSetNo As Integer, ByVal pDesc As String)
    SetControlTheme()
    SettingsName = "ActionSchedule"
    MainHelper.SetMDIParent(Me)
    tab.SelectedIndex = DisplayTypes.Quarter
    mvSelectionSetNumber = pSelectionSetNo
    Me.Text = String.Format("Action Schedule - {0}", pDesc)

    vas.TextTipPolicy = FarPoint.Win.Spread.TextTipPolicy.Floating
    vas.TextTipAppearance = New FarPoint.Win.Spread.TipAppearance(Color.Yellow, Color.Black, vas.Font)
    vas.Sheets(0).RowHeaderAutoText = FarPoint.Win.Spread.HeaderAutoText.Blank
    vas.Sheets(0).ColumnHeader.Rows(0).Height = vas.Sheets(0).ColumnHeader.Rows(0).Height * 2
    Populate()
  End Sub

  Public Overrides Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes)
    Populate()
  End Sub

  Public Overrides Sub RefreshData()
    Populate()
  End Sub

  Private Sub Populate()
    Dim vList As New ParameterList(True)
    Dim vStartDate As Date
    Dim vEndDate As Date
    Dim vRequiredColumns As Integer

    vStartDate = GetStartDate()
    Select Case mvDisplayType
      Case DisplayTypes.Day
        vEndDate = DateAdd(DateInterval.Day, 1, mvFirstDate)
        vRequiredColumns = 1
      Case DisplayTypes.Week
        vEndDate = DateAdd(DateInterval.Day, 7, vStartDate)
        vRequiredColumns = 7
      Case DisplayTypes.Month
        vEndDate = DateAdd(DateInterval.Month, 1, vStartDate)
        vRequiredColumns = CInt(DateDiff(DateInterval.Day, vStartDate, vEndDate))
      Case DisplayTypes.Quarter
        vEndDate = DateAdd(DateInterval.Quarter, 1, vStartDate)
        vRequiredColumns = CInt(DateDiff(DateInterval.Day, vStartDate, vEndDate)) \ 7
      Case DisplayTypes.Year
        vEndDate = DateAdd(DateInterval.Year, 1, vStartDate)
        vRequiredColumns = 12
    End Select
    vList.Add("StartDate", vStartDate)
    vList.Add("EndDate", vEndDate)
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(DataHelper.GetSelectionSetAppointments(mvSelectionSetNumber, vList))
    With vas_Sheet1
      .RowCount = 0
      .ColumnCount = COL_FIRST_DATE + vRequiredColumns
      .Columns(COL_CONTACT_NUMBER).Visible = False

      .ColumnHeader.Cells(0, COL_CONTACT_NAME).Text = "Contact Name"
      .ColumnHeader.Cells(0, COL_REQUEST_STAGE).Text = "Stage"
      .FrozenColumnCount = COL_FIRST_DATE

      Dim vDate As Date
      Dim vLastDate As Date
      Dim vColumnWidth As Integer = 100
      For vIndex As Integer = 0 To vRequiredColumns - 1
        vDate = GetColumnDate(vStartDate, vIndex)
        Select Case mvDisplayType
          Case DisplayTypes.Day
            vLastDate = vDate
            .ColumnHeader.Cells(0, COL_FIRST_DATE + vIndex).Text = vDate.ToString("ddd dd MMM yyyy")
            vColumnWidth = 300
          Case DisplayTypes.Week
            vLastDate = vDate
            .ColumnHeader.Cells(0, COL_FIRST_DATE + vIndex).Text = vDate.ToString("ddd dd MMM yyyy")
            vColumnWidth = 150
          Case DisplayTypes.Month
            vLastDate = vDate
            .ColumnHeader.Cells(0, COL_FIRST_DATE + vIndex).Text = vDate.ToString("ddd dd/MM")
            vColumnWidth = 40
          Case DisplayTypes.Quarter
            vLastDate = vDate.AddDays(6)
            .ColumnHeader.Cells(0, COL_FIRST_DATE + vIndex).Text = vDate.ToString("dd MMM yy")
          Case DisplayTypes.Year
            vLastDate = vDate.AddMonths(1).AddDays(-1)
            .ColumnHeader.Cells(0, COL_FIRST_DATE + vIndex).Text = vDate.ToString("MMM yy")
        End Select
        .Columns(COL_FIRST_DATE + vIndex).Width = vColumnWidth
        If vLastDate < Today Then
          .Columns(COL_FIRST_DATE + vIndex).BackColor = Color.Gainsboro
        Else
          .Columns(COL_FIRST_DATE + vIndex).BackColor = Color.White
        End If
      Next
      If vTable IsNot Nothing Then
        Dim vRowNumber As Integer
        Dim vLastContactNumber As Integer
        Dim vColNo As Integer
        For Each vRow As DataRow In vTable.Rows
          If vLastContactNumber <> IntegerValue(vRow("ContactNumber").ToString) Then
            .RowCount = .RowCount + 1
            vRowNumber = .RowCount - 1
            .SetValue(vRowNumber, COL_CONTACT_NUMBER, vRow("ContactNumber"))
            .SetValue(vRowNumber, COL_CONTACT_NAME, vRow("ContactName"))
            .SetValue(vRowNumber, COL_REQUEST_STAGE, vRow("RequestStageDesc"))
            vLastContactNumber = IntegerValue(vRow("ContactNumber").ToString)
          End If
          Dim vDateString As String = vRow("StartDate").ToString
          If vDateString.Length > 0 Then
            Dim vAppDate As Date = CDate(vDateString)
            Dim vRowNo As Integer = vRowNumber
            Select Case mvDisplayType
              Case DisplayTypes.Day
                vColNo = COL_FIRST_DATE
              Case DisplayTypes.Week
                Dim vDays As Integer = CInt(DateDiff(DateInterval.Day, vStartDate, vAppDate))
                vColNo = COL_FIRST_DATE + vDays
              Case DisplayTypes.Month
                Dim vDays As Integer = CInt(DateDiff(DateInterval.Day, vStartDate, vAppDate))
                vColNo = COL_FIRST_DATE + vDays
              Case DisplayTypes.Quarter
                Dim vDays As Integer = CInt(DateDiff(DateInterval.Day, vStartDate, vAppDate))
                vColNo = COL_FIRST_DATE + (vDays \ 7)
              Case DisplayTypes.Year
                Dim vMonths As Integer = CInt(DateDiff(DateInterval.Month, vStartDate, vAppDate))
                vColNo = COL_FIRST_DATE + vMonths
            End Select
            While .GetTag(vRowNo, vColNo) IsNot Nothing
              vRowNo += 1
              If vRowNo = .RowCount Then .RowCount = .RowCount + 1
            End While
            .SetValue(vRowNo, vColNo, vRow("Description"))
            If .Columns(vColNo).BackColor = Color.White Then
              .Cells(vRowNo, vColNo).ForeColor = SystemColors.HighlightText
              .Cells(vRowNo, vColNo).BackColor = SystemColors.Highlight
            End If
            Dim vAppointment As New Appointment
            vAppointment.StartDate = vAppDate
            vAppointment.EndDate = CDate(vRow("EndDate"))
            vAppointment.RecordType = vRow("RecordType").ToString
            vAppointment.UniqueID = IntegerValue(vRow("UniqueID").ToString)
            .SetTag(vRowNo, vColNo, vAppointment)
          End If
        Next
      End If
      .Columns(COL_CONTACT_NAME).Width = .Columns(COL_CONTACT_NAME).GetPreferredWidth
      .Columns(COL_REQUEST_STAGE).Width = .Columns(COL_REQUEST_STAGE).GetPreferredWidth
      .Cells(0, 0, .RowCount - 1, .ColumnCount - 1).Locked = True
    End With
  End Sub

  Private Function GetColumnDate(ByVal pStartDate As Date, ByVal pIndex As Integer) As Date
    Select Case mvDisplayType
      Case DisplayTypes.Day
        Return pStartDate
      Case DisplayTypes.Week
        Return pStartDate.AddDays(pIndex)
      Case DisplayTypes.Month
        Return pStartDate.AddDays(pIndex)
      Case DisplayTypes.Quarter
        Return pStartDate.AddDays(pIndex * 7)
      Case DisplayTypes.Year
        Return pStartDate.AddMonths(pIndex)
    End Select
  End Function

  Private Function GetStartDate() As Date
    Select Case mvDisplayType
      Case DisplayTypes.Day
        Return mvFirstDate
      Case DisplayTypes.Week
        Dim vOffset As Integer = mvFirstDate.DayOfWeek - 1
        If vOffset < 0 Then
          vOffset = -6
        ElseIf vOffset > 0 Then
          vOffset = -vOffset
        End If
        Return mvFirstDate.AddDays(vOffset)
      Case DisplayTypes.Month
        Return New Date(mvFirstDate.Year, mvFirstDate.Month, 1)
      Case DisplayTypes.Quarter
        Dim vStartMonth As Integer
        Select Case mvFirstDate.Month
          Case 1 To 3
            vStartMonth = 1
          Case 4 To 6
            vStartMonth = 4
          Case 7 To 9
            vStartMonth = 7
          Case Else
            vStartMonth = 10
        End Select
        Return New Date(mvFirstDate.Year, vStartMonth, 1)
      Case DisplayTypes.Year
        Return New Date(mvFirstDate.Year, 1, 1)
    End Select
  End Function

  Private Sub tab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab.SelectedIndexChanged
    If mvSelectionSetNumber > 0 AndAlso tab.SelectedIndex >= 0 Then
      mvDisplayType = CType(tab.SelectedIndex, DisplayTypes)
      Populate()
    End If
  End Sub

  Private Sub tab_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles tab.Selecting
    If e.Action = TabControlAction.Selecting Then
      If e.TabPage Is tbpRight Then
        e.Cancel = True
        Select Case mvDisplayType
          Case DisplayTypes.Day
            mvFirstDate = mvFirstDate.AddDays(1)
          Case DisplayTypes.Week
            mvFirstDate = mvFirstDate.AddDays(7)
          Case DisplayTypes.Month
            mvFirstDate = mvFirstDate.AddMonths(1)
          Case DisplayTypes.Quarter
            mvFirstDate = mvFirstDate.AddMonths(3)
          Case DisplayTypes.Year
            mvFirstDate = mvFirstDate.AddYears(1)
        End Select
        Populate()
      ElseIf e.TabPage Is tbpLeft Then
        e.Cancel = True
        If mvFirstDate > Today Then
          Select Case mvDisplayType
            Case DisplayTypes.Day
              mvFirstDate = mvFirstDate.AddDays(-1)
            Case DisplayTypes.Week
              mvFirstDate = mvFirstDate.AddDays(-7)
            Case DisplayTypes.Month
              mvFirstDate = mvFirstDate.AddMonths(-1)
            Case DisplayTypes.Quarter
              mvFirstDate = mvFirstDate.AddMonths(-3)
            Case DisplayTypes.Year
              mvFirstDate = mvFirstDate.AddYears(-1)
          End Select
          Populate()
        End If
      End If
    End If
  End Sub

  Private Sub vas_CellClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles vas.CellClick
    If e.RowHeader Then
      If e.Column >= 0 Then
        vas.ContextMenuStrip = mvBrowserMenu
      End If
      mvBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
      mvBrowserMenu.ItemNumber = CInt(vas.Sheets(0).GetValue(e.Row, COL_CONTACT_NUMBER))
    End If
  End Sub

  Private Sub vas_CellDoubleClick(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.CellClickEventArgs) Handles vas.CellDoubleClick
    If e.Column = COL_CONTACT_NAME Then
      FormHelper.ShowContactCardIndex(CInt(vas.Sheets(0).GetValue(e.Row, COL_CONTACT_NUMBER)))
    ElseIf e.Column >= COL_FIRST_DATE Then
      Dim vTag As Object = vas.Sheets(0).GetTag(e.Row, e.Column)
      If vTag IsNot Nothing AndAlso DirectCast(vTag, Appointment).RecordType = "A" Then
        FormHelper.EditAction(DirectCast(vTag, Appointment).UniqueID, Me)
      Else
        Dim vDate As Date = GetColumnDate(GetStartDate, e.Column - COL_FIRST_DATE)
        Dim vContactNumber As Integer
        Dim vRow As Integer = e.Row
        Do
          vContactNumber = IntegerValue(vas.Sheets(0).GetText(vRow, COL_CONTACT_NUMBER).ToString)
          vRow -= 1
        Loop While vContactNumber = 0
        Dim vContactInfo As New ContactInfo(vContactNumber)
        Dim vList As New ParameterList
        vList("ScheduledOn") = String.Format("{0} {1}", vDate.ToString(AppValues.DateFormat), AppValues.DefaultActionStartOfDay)
        FormHelper.EditAction(0, Me, vList, vContactInfo)
      End If
    End If
  End Sub

  Private Sub vas_TextTipFetch(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.TextTipFetchEventArgs) Handles vas.TextTipFetch
    Dim vSS As FarPoint.Win.Spread.FpSpread = DirectCast(sender, FarPoint.Win.Spread.FpSpread)
    With vSS.Sheets(0)
      If e.Column > COL_CONTACT_NAME Then
        Dim vDesc As String = .Cells(e.Row, e.Column).Text
        If vDesc.Length > 0 Then
          e.ShowTip = True
          Dim vAppointment As Appointment = CType(.Cells(e.Row, e.Column).Tag, Appointment)
          e.TipText = String.Format("{0} {1} - {2}", vDesc, vAppointment.StartDate, vAppointment.EndDate)
        End If
      End If
    End With
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub

End Class