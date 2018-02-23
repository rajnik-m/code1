Public Class frmNotify
  Inherits PersistentForm

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    InitialiseControls()

  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents cmdRefresh As System.Windows.Forms.Button
  Friend WithEvents cmdClearAll As System.Windows.Forms.Button
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents pnl As System.Windows.Forms.Panel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNotify))
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdClear = New System.Windows.Forms.Button
    Me.cmdClearAll = New System.Windows.Forms.Button
    Me.cmdRefresh = New System.Windows.Forms.Button
    Me.cmdClose = New System.Windows.Forms.Button
    Me.pnl = New System.Windows.Forms.Panel
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl.SuspendLayout()
    Me.pnl.SuspendLayout()
    Me.SuspendLayout()
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdClear)
    Me.bpl.Controls.Add(Me.cmdClearAll)
    Me.bpl.Controls.Add(Me.cmdRefresh)
    Me.bpl.Controls.Add(Me.cmdClose)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 237)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(576, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdClear
    '
    Me.cmdClear.Location = New System.Drawing.Point(77, 6)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(94, 27)
    Me.cmdClear.TabIndex = 0
    Me.cmdClear.Text = "&Clear"
    '
    'cmdClearAll
    '
    Me.cmdClearAll.Location = New System.Drawing.Point(186, 6)
    Me.cmdClearAll.Name = "cmdClearAll"
    Me.cmdClearAll.Size = New System.Drawing.Size(94, 27)
    Me.cmdClearAll.TabIndex = 1
    Me.cmdClearAll.Text = "Clear &All"
    '
    'cmdRefresh
    '
    Me.cmdRefresh.Location = New System.Drawing.Point(295, 6)
    Me.cmdRefresh.Name = "cmdRefresh"
    Me.cmdRefresh.Size = New System.Drawing.Size(94, 27)
    Me.cmdRefresh.TabIndex = 2
    Me.cmdRefresh.Text = "Refresh"
    '
    'cmdClose
    '
    Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdClose.Location = New System.Drawing.Point(404, 6)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.Size = New System.Drawing.Size(94, 27)
    Me.cmdClose.TabIndex = 3
    Me.cmdClose.Text = "Close"
    '
    'pnl
    '
    Me.pnl.Controls.Add(Me.dgr)
    Me.pnl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnl.Location = New System.Drawing.Point(0, 0)
    Me.pnl.Name = "pnl"
    Me.pnl.Size = New System.Drawing.Size(576, 237)
    Me.pnl.TabIndex = 2
    '
    'dgr
    '
    Me.dgr.AccessibleDescription = "Display List"
    Me.dgr.AccessibleName = "Display List"
    Me.dgr.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(576, 237)
    Me.dgr.TabIndex = 1
    '
    'frmNotify
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.CancelButton = Me.cmdClose
    Me.ClientSize = New System.Drawing.Size(576, 276)
    Me.Controls.Add(Me.pnl)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmNotify"
    Me.bpl.ResumeLayout(False)
    Me.pnl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub InitialiseControls()
    SetControlTheme()
    Me.cmdClear.Text = ControlText.CmdClear
    Me.cmdClearAll.Text = ControlText.CmdClearnAll
    Me.cmdRefresh.Text = ControlText.CmdNRefresh
    Me.cmdClose.Text = ControlText.CmdClose
    Me.Text = ControlText.frmNotify
    SettingsName = "Notify"
    MainHelper.SetMDIParent(Me)
    dgr.AutoSetHeight = False
    DoRefresh()
  End Sub
  
  Private Sub UpdateNotifications(ByVal pDataSet As DataSet)
    Try
      Dim vCount As Integer = dgr.RowCount
      Dim vLastRow As Integer = dgr.CurrentRow
      If vLastRow < 0 Then vLastRow = 0
      dgr.Clear()
      dgr.Populate(pDataSet)
      If dgr.RowCount = 0 Then
        Me.Close()
      Else
        If vLastRow >= dgr.RowCount Then vLastRow = dgr.RowCount - 1
        dgr.SelectRow(vLastRow)
        Me.Text = String.Format(InformationMessages.frmNotifyCaption, dgr.RowCount)
        If vCount < dgr.RowCount And Me.WindowState = FormWindowState.Minimized Then
          Me.WindowState = FormWindowState.Normal
          DoBeep()
        End If
        Dim vCol As Integer = dgr.GetColumn("ItemCode")
        Dim vActions As Integer
        Dim vDocuments As Integer
        Dim vMeeting As Integer
        Dim vAlerts As Integer
        For vRow As Integer = 0 To dgr.RowCount - 1
          Select Case dgr.GetValue(vRow, vCol)
            Case "A"
              vActions += 1
            Case "D"
              vDocuments += 1
            Case "M"
              vMeeting += 1
            Case "I"
              vAlerts += 1
          End Select
        Next
        vCount = vActions + vDocuments + vMeeting
        cmdClear.Enabled = (vCount > 0 OrElse vAlerts > 0)
        cmdClearAll.Enabled = (vCount > 0 OrElse vAlerts > 0)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Public Sub DoRefresh()
    If Me.InvokeRequired Then
      Me.Invoke(New MethodInvoker(AddressOf DoRefresh))
    Else
      Dim vSelect As Boolean = (Settings.NotifyActions Or Settings.NotifyDeadlines Or Settings.NotifyDocuments Or Settings.NotifyMeetings) AndAlso DataHelper.UserInfo.ContactNumber > 0
      Dim vDataSet As DataSet = Nothing
      If vSelect Then
        Dim vList As New ParameterList(True)
        If Settings.NotifyActions Then vList("NotifyActions") = "Y"
        If Settings.NotifyDocuments Then vList("NotifyDocuments") = "Y"
        If Settings.NotifyDeadlines Then vList("NotifyDeadlines") = "Y"
        If Settings.NotifyMeetings Then vList("NotifyMeetings") = "Y"
        vDataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactNotifications, DataHelper.UserInfo.ContactNumber, vList)
      End If
      MainHelper.UpdateNotificationIcon(vDataSet)
      UpdateNotifications(vDataSet)   'May close
    End If
  End Sub

  Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Me.Close()
  End Sub
  Private Sub frmNotify_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    FormHelper.NotifyForm = Nothing
  End Sub
  Private Sub dgr_RowDoubleClicked(ByVal pSender As Object, ByVal pRow As Integer) Handles dgr.RowDoubleClicked
    Try
      Dim vItemCode As String = dgr.GetValue(pRow, "ItemCode")
      Dim vNumberCol As Integer = dgr.GetColumn("ItemNumber")
      Dim vItemNumber As Integer = CInt(dgr.GetValue(pRow, vNumberCol))
      Dim vItemNumbers As New ArrayListEx
      Dim vList As New ParameterList
      For vRow As Integer = 0 To dgr.RowCount - 1
        vItemNumbers.Add(dgr.GetValue(vRow, vNumberCol))
      Next
      Select Case vItemCode
        Case "A", "O"
          vList("ActionNumbers") = vItemNumbers.CSList
          vList.IntegerValue("SelectActionNumber") = vItemNumber
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftActions, vList)
        Case "D"
          vList("DocumentNumbers") = vItemNumbers.CSList
          vList.IntegerValue("SelectDocumentNumber") = vItemNumber
          FormHelper.ShowFinder(CareServices.XMLDataFinderTypes.xdftDocuments, vList)
        Case "M"
          vList("MeetingNumbers") = vItemNumbers.CSList
          vList.IntegerValue("SelectMeetingNumber") = vItemNumber
          FormHelper.ShowFinder(CType(CareNetServices.XMLDataFinderTypes.xdftMeetings, CareServices.XMLDataFinderTypes), vList)
      End Select
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub cmdRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
    Dim vCursor As New BusyCursor
    Try
      DoRefresh()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Try
      Dim vItemCode As String = dgr.GetValue(dgr.CurrentRow, "ItemCode")
      Dim vItemNumber As Integer = CInt(dgr.GetValue(dgr.CurrentRow, "ItemNumber"))
      Dim vList As New ParameterList(True)
      vList("Notified") = "Y"
      Select Case vItemCode
        Case "A"
          vList.IntegerValue("ActionNumber") = vItemNumber
          vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
        Case "D"
          vList.IntegerValue("DocumentNumber") = vItemNumber
          vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, vList)
        Case "I"
          vList.IntegerValue("ItemNumber") = vItemNumber
          DataHelper.UpdateEntityAlertItem(vList)
        Case "M"
          vList.IntegerValue("MeetingNumber") = vItemNumber
          vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
          vList("LinkType") = dgr.GetValue(dgr.CurrentRow, "LinkType")
          DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctMeetingLinks, vList)

      End Select
      DoRefresh()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
    Try
      Dim vList As New ParameterList(True)
      vList("Notified") = "Y"
      vList.IntegerValue("ContactNumber") = DataHelper.UserInfo.ContactNumber
      DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctActionLink, vList)
      DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctDocumentLink, vList)
      DataHelper.UpdateEntityAlertItem(vList)
      DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctMeetingLinks, vList)
      DoRefresh()
      MainHelper.SetNotificationTime()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

End Class
