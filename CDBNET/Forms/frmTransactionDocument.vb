Public Class frmTransactionDocument

  Public Enum TransactionDocumentTypes
    tdtTransaction
    tdtPaymentPlan
  End Enum

  Private mvDataSet As DataSet
  Private mvList As ParameterList

  Public Sub New(ByVal pTransDocumentType As TransactionDocumentTypes, ByRef pDataSet As DataSet, ByRef pList As ParameterList)
    mvDataSet = pDataSet
    mvList = pList
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pTransDocumentType)
  End Sub
  Private Sub InitialiseControls(ByVal pTransDocumentType As TransactionDocumentTypes)
    SetControlTheme()
    If pTransDocumentType = TransactionDocumentTypes.tdtPaymentPlan Then
      Me.Text = ControlText.FrmTransactionDocumentPP
    Else
      Me.Text = ControlText.FrmTransactionDocument
    End If
    With epl
      .Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optTransactionDocument))
      .SetValue("Mailing", mvList("Mailing"))
      .SetValue("EarliestFulfilmentDate", AppValues.TodaysDate)
      .FindDateTimePicker("EarliestFulfilmentDate").ShowCheckBox = False
    End With
    PopulateDataGrid()
    cmdEdit.Enabled = BooleanValue(mvDataSet.Tables(1).Rows(0).Item("EditDocument").ToString)
  End Sub
  Private Sub PopulateDataGrid()
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(mvDataSet)
    If vTable IsNot Nothing Then
      vTable.Columns.Add(New DataColumn("IncludeNew", GetType(System.Boolean)))
      For Each vRow As DataRow In vTable.Rows
        vRow("IncludeNew") = BooleanValue(vRow("Include").ToString)
      Next
      vTable.Columns.Item("Include").ColumnName = "IncludeOld"
      vTable.Columns.Item("IncludeNew").ColumnName = "Include"
    End If
    With dgr
      If .Populate(mvDataSet, "IncludeOld = 'Y'") > 0 Then
        .SetCellsEditable()
        .SetCellsReadOnly()
        .SetCheckBoxColumn("Include")
        .AllowSorting = False
        Dim vCol As Integer = dgr.GetColumn("Include")
        For vRow As Integer = 0 To dgr.RowCount - 1
          If dgr.GetValue(vRow, "Mandatory") = "Y" Then dgr.SetCellsReadOnly(vRow, vCol, , True)
        Next
      End If
    End With
  End Sub

  Private Sub frmTransactionDocument_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    If DocumentApplication IsNot Nothing Then DocumentApplication.ProcessAppActive()
  End Sub
  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    If epl.AddValuesToList(mvList, True) Then
      Me.Close()
    End If
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    If pParameterName = "Mailing" Then
      pList("TransactionDocument") = "Y"
    End If
  End Sub
  Private Sub epl_ValueChanged(ByVal pSender As Object, ByVal pParameterName As String, ByVal pValue As String) Handles epl.ValueChanged
    Select Case pParameterName
      Case "Mailing"
        mvList("Mailing") = epl.GetValue("Mailing")
        mvDataSet = DataHelper.GetMailingDocumentParagraphs(mvList)
        PopulateDataGrid()
    End Select
  End Sub
  Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
    Dim vTable As DataTable = DataHelper.GetTableFromDataSet(mvDataSet)
    If vTable IsNot Nothing Then
      If vTable.Columns.Contains("MailingTemplate") AndAlso vTable.Rows(0).Item("MailingTemplate").ToString.Length > 0 AndAlso vTable.Columns.Contains("StandardDocument") AndAlso vTable.Rows(0).Item("StandardDocument").ToString.Length > 0 Then
        Dim vFileName As String = DataHelper.GetTempFile("")
        Dim vList As New ParameterList(True)
        vList("MailingTemplate") = vTable.Rows(0).Item("MailingTemplate").ToString
        If DataHelper.GetMailingDocumentMergeFile(vList, vFileName) Then
          Dim vLookupList As New ParameterList(True)
          vLookupList("StandardDocument") = vTable.Rows(0).Item("StandardDocument").ToString
          Dim vRow As DataRow = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtStandardDocuments, vLookupList).Rows(0)
          If TypeOf DocumentApplication Is WordApplication Then DirectCast(DocumentApplication, WordApplication).BuildMailingDocument(vRow.Item("StandardDocument").ToString, vRow.Item("DocfileExtension").ToString, vFileName, vTable)
        End If
      End If
    End If
  End Sub
  Public ReadOnly Property DataSet() As DataSet
    Get
      Return mvDataSet
    End Get
  End Property
End Class