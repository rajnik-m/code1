Public Class frmModuleContent

  Private mvWebPageItemNumber As Integer
  Private mvDataSet As DataSet

  Public Sub New(ByVal pWebPageItemNumber As Integer)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pWebPageItemNumber)
  End Sub

  Private Sub InitialiseControls(ByVal pWebPageItemNumber As Integer)
    mvWebPageItemNumber = pWebPageItemNumber
    Dim vParams As New ParameterList(True)
    vParams.IntegerValue("WebPageItemNumber") = mvWebPageItemNumber
    mvDataSet = DataHelper.GetWebData(CareWebAccess.XMLWebDataSelectionTypes.wstPageItemControls, vParams)
    dgr.Populate(mvDataSet)
    dgr.SetCellsEditable()
    dgr.SetCellsReadOnly(-1, -1, True, True)
    dgr.SetColumnVisible("OldSequenceNumber", False)
    dgr.SetCheckBoxColumn("Visible")
    dgr.SetBackgroundColour("Visible", Color.White)
    dgr.SetCheckBoxColumn("MandatoryItem")
    dgr.SetBackgroundColour("MandatoryItem", Color.White)
    dgr.SetColumnWritable("ControlCaption")
    dgr.SetBackgroundColour("ControlCaption", Color.White)
    dgr.SetColumnWritable("ControlWidth")
    dgr.SetBackgroundColour("ControlWidth", Color.White)
    dgr.SetColumnWritable("ControlHeight")
    dgr.SetBackgroundColour("ControlHeight", Color.White)
    dgr.AllowRowMove()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vParams As New ParameterList(True)
    vParams.IntegerValue("WebPageItemNumber") = mvWebPageItemNumber
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(mvDataSet)
    If vTable IsNot Nothing Then
      For Each vRow As DataRow In vTable.Rows
        vParams.ObjectValue("WebPageItemControl" & vRow("SequenceNumber").ToString) = vRow
      Next
      DataHelper.UpdateWebPageItemControls(vParams)
    End If
  End Sub

  Private Sub cmdRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRevert.Click
    If ShowQuestion("Confirm revert the content of this item to default values?", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
      Dim vParams As New ParameterList(True)
      vParams.IntegerValue("WebPageItemNumber") = mvWebPageItemNumber
      DataHelper.UpdateWebPageItemControls(vParams)
      Me.DialogResult = Windows.Forms.DialogResult.OK
      Me.Close()
    End If
  End Sub

  Private Sub frmModuleContent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    bpl.RepositionButtons()
  End Sub

  Private Sub dgr_RowMoved(ByVal sender As Object, ByVal pOldRowNumber As Integer, ByVal pNewRowNumber As Integer) Handles dgr.RowMoved
    Dim vSequenceNumber As Integer
    For vRow As Integer = 0 To dgr.RowCount - 1
      vSequenceNumber += 1
      dgr.SetValue(vRow, "SequenceNumber", vSequenceNumber.ToString)
    Next
  End Sub
End Class