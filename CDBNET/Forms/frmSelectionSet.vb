Public Class frmSelectionSet
  Inherits ThemedForm

#Region " Windows Form Designer generated code "

  Public Sub New(ByVal pSSNo As Integer, ByVal pDesc As String)
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    InitialiseControls(pSSNo, pDesc)
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
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents bpl As CDBNETCL.ButtonPanel
  Friend WithEvents dgr As CDBNETCL.DisplayGrid
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectionSet))
    Me.dgr = New CDBNETCL.DisplayGrid
    Me.bpl = New CDBNETCL.ButtonPanel
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.bpl.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgr
    '
    Me.dgr.AccessibleName = "Display Grid"
    Me.dgr.AllowSorting = True
    Me.dgr.AutoSetHeight = False
    Me.dgr.AutoSetRowHeight = False
    Me.dgr.DisplayTitle = Nothing
    Me.dgr.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgr.Location = New System.Drawing.Point(0, 0)
    Me.dgr.MaxGridRows = 8
    Me.dgr.MultipleSelect = False
    Me.dgr.Name = "dgr"
    Me.dgr.RowCount = 10
    Me.dgr.ShowIfEmpty = False
    Me.dgr.Size = New System.Drawing.Size(680, 257)
    Me.dgr.TabIndex = 0
    '
    'bpl
    '
    Me.bpl.Controls.Add(Me.cmdCancel)
    Me.bpl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.bpl.DockingStyle = CDBNETCL.ButtonPanel.ButtonPanelDockingStyles.bpdsBottom
    Me.bpl.Location = New System.Drawing.Point(0, 257)
    Me.bpl.Name = "bpl"
    Me.bpl.Size = New System.Drawing.Size(680, 39)
    Me.bpl.TabIndex = 1
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(292, 6)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(96, 27)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "Close"
    '
    'frmSelectionSet
    '
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(680, 296)
    Me.Controls.Add(Me.dgr)
    Me.Controls.Add(Me.bpl)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmSelectionSet"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.bpl.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Dim WithEvents mvBrowserMenu As BrowserMenu
  Dim mvSSNo As Integer
  Dim mvContactNumber As Integer

  Private Sub InitialiseControls(ByVal pSSNo As Integer, ByVal pDesc As String)
    SetControlTheme()
    Me.cmdCancel.Text = ControlText.CmdClose
    MainHelper.SetMDIParent(Me)
    Me.Text = pDesc
    mvSSNo = pSSNo
    mvBrowserMenu = New BrowserMenu(Nothing)
    mvBrowserMenu.RemoveSupported = True
    Dim vDataSet As DataSet = DataHelper.GetSelectionSetData(mvSSNo)
    If Not vDataSet Is Nothing Then
      dgr.Populate(vDataSet)
      dgr.ContextMenuStrip = mvBrowserMenu
      dgr.AllowDrop = True
      If dgr.RowCount > 0 Then dgr.SelectRow(0)
    End If
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub dgr_ContactSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pContactNumber As Integer) Handles dgr.ContactSelected
    FormHelper.ShowContactCardIndex(pContactNumber)
  End Sub

  Private Sub dgr_RowSelected(ByVal pSender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgr.RowSelected
    mvBrowserMenu.EntityType = HistoryEntityTypes.hetContacts
    mvBrowserMenu.ItemNumber = CInt(dgr.GetValue(pRow, "ContactNumber"))
    mvContactNumber = mvBrowserMenu.ItemNumber
  End Sub

  Private Sub dgr_ContactDropped(ByVal pSender As Object, ByVal pContactInfo As ContactInfo) Handles dgr.ContactDropped
    Try
      Dim vDoBeep As Boolean
      If String.IsNullOrEmpty(pContactInfo.SelectedContactNumbers) Then
        vDoBeep = AddContact(pContactInfo)
      Else
        'BR20416 - This looks like it is trying to deal with multiple selects, Contact Finder results will only allow single select.
        '          The first drag populates SelectedContactNumbers, subsequent drags from Contact Finder leave it as an empty string.
        For Each vContactNumber As String In pContactInfo.SelectedContactNumbers.Split(","c)
          pContactInfo.ContactNumber = IntegerValue(vContactNumber)
          vDoBeep = AddContact(pContactInfo)
        Next
      End If
      If vDoBeep Then Beep()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function AddContact(ByVal pContactInfo As ContactInfo) As Boolean
    Try
      Dim vContactCol As Integer = dgr.GetColumn("ContactNumber")
      If dgr.FindRow(vContactCol, IntegerValue(pContactInfo.ContactNumber)) = False Then
        DataHelper.AddSelectionSetContact(mvSSNo, pContactInfo)
        Dim vDataSet As DataSet = DataHelper.GetSelectionSetData(mvSSNo)
        dgr.Populate(vDataSet)
        UserHistory.UpdateSelectionSetData(vDataSet, mvSSNo)
        dgr.FindRow(vContactCol, pContactInfo.ContactNumber)
      Else
        Return True
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Function

  Private Sub mvBrowserMenu_Remove(ByVal pEntityType As CDBNETCL.HistoryEntityTypes, ByVal pNumber As Integer) Handles mvBrowserMenu.Remove
    Try
      Dim vContactCol As Integer = dgr.GetColumn("ContactNumber")
      If dgr.FindRow(vContactCol, pNumber) = True Then
        DataHelper.DeleteSelectionSetContact(mvSSNo, pNumber, 0)
        Dim vDataSet As DataSet = DataHelper.GetSelectionSetData(mvSSNo)
        dgr.Populate(vDataSet)
        If dgr.RowCount > 0 Then dgr.SelectRow(0)
        UserHistory.UpdateSelectionSetData(vDataSet, mvSSNo)
      Else
        Beep()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Public Sub RefreshData(ByVal pDataSet As DataSet)
    Me.dgr.Populate(pDataSet)
  End Sub

  Public ReadOnly Property SelectionSetNumber() As Integer
    Get
      Return mvSSNo
    End Get
  End Property

  Public ReadOnly Property SelectedContactNumber() As Integer
    Get
      Return mvContactNumber
    End Get
  End Property

End Class
