Public Class frmReportDataSelection

  
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

  End Sub

  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

  End Sub

  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

  End Sub

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub

  Private Sub InitialiseControls()
#If DEBUG Then
    Static mvDebugTest As New DebugTest(Me)       'Include this statement so we can keep track of memory leakage of forms
#End If
    SetControlColors(Me)
    lblAvailableItems.Text = "Available Items"
    lblAvailableItems.Visible = True
    lblOrder.Text = "Order:"
    lblOrder.Visible = True
    lblReportDataSet.Text = "Report Data Set:"
    lblReportDataSet.Visible = True
    lblReportType.Text = "Report Type:"
    lblReportType.Visible = True
    lblSelectedItems.Text = "Selected Items"
    lblSelectedItems.Visible = True
    lblSourceReport.Text = "Source Report:"
    lblSourceReport.Visible = True

    cmdOK.Enabled = False
    cmdSave.Enabled = False
    cboOutputType.Items.Add("Landscape Report")
    cboOutputType.Items.Add("Portrait Report")
    cboOutputType.Items.Add("Mail Merge Report")
    cboOutputType.SelectedIndex = 0

    cboOrder.Items.Add("Contact Number")
    cboOrder.Items.Add("Surname")
    cboOrder.Items.Add("Country/Town")
    cboOrder.Items.Add("Address Branch")
    cboOrder.SelectedIndex = 0

  End Sub

  Private Sub frmReportDataSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    bpl.RepositionButtons()
    Dim vList As New ParameterList(True)
    vList("ReportCode") = "SSDSRP"
    Dim vTable As DataTable = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtReports, vList)
    cboReport.DisplayMember = "ReportName"
    cboReport.ValueMember = "ReportNumber"
    cboReport.DataSource = vTable

  End Sub
End Class