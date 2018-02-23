Public Class frmGenMAddress
  Private mvContactNumber As Integer
  Private mvAddressNumber As Integer
  'Private mvRevision As Integer
  'Private mvSelectionSet As Integer
  Private mvMailingInfo As MailingInfo

  Dim mvDataTable As DataTable = Nothing
  Dim mvFormText As String = ""

  Public ReadOnly Property AddressNumber() As Integer
    Get
      Return mvAddressNumber
    End Get
  End Property

  Public Sub New(ByVal pContactNumber As Integer, ByVal pAddressNumber As Integer, ByVal pMailingInfo As MailingInfo)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    mvContactNumber = pContactNumber
    mvAddressNumber = pAddressNumber
    mvFormText = pMailingInfo.Caption
    mvMailingInfo = pMailingInfo
    'mvSelectionSet = pSelectionSet
    'mvRevision = pRevision

    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls()
  End Sub
  Private Sub InitialiseControls()
    Me.Text = mvFormText + " - Contact Address Selection"
    SetControlTheme()
    drg.MaxGridRows = DisplayTheme.DefaultMaxGridRows
    Dim vParams As New ParameterList(True)
    'vParams.IntegerValue("AddressNumber") = mvAddressNumber
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddresses, mvContactNumber, vParams)
    If vDataSet IsNot Nothing Then 'And vDataSet.Tables.Count > 1 Then mvDataTable = vDataSet.Tables(1)
      If vDataSet.Tables.Contains("Column") Then
        vDataSet.Tables("Column").Clear()

        DataHelper.AddDataColumn(vDataSet.Tables("Column"), "AddressNumber", "Number")
        DataHelper.AddDataColumn(vDataSet.Tables("Column"), "AddressLine", "Address")
        DataHelper.AddDataColumn(vDataSet.Tables("Column"), "Historical", "Hist")
        DataHelper.AddDataColumn(vDataSet.Tables("Column"), "NewColumn", "Usage")
      End If

      drg.Populate(vDataSet)
      GetAddressUsages()
    End If
  End Sub

  Private Function GetAddressUsages() As String
    Dim vParams As New ParameterList(True)
    Dim vResult As String = ""
    Dim vAddressNumber As Integer
    vParams("ContactNumber") = mvContactNumber.ToString
    vParams("GetAllAddressUsages") = "Y"
    Dim vRows() As DataRow = Nothing
    Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactAddressUsages, mvContactNumber, vParams)
    If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
      For vRowIndex As Integer = 0 To drg.RowCount - 1
        vAddressNumber = IntegerValue(drg.GetValue(vRowIndex, "AddressNumber"))
        vRows = vDataSet.Tables("DataRow").Select(String.Format("AddressNumber = '{0}'", vAddressNumber))
        If vRows.Length > 0 Then
          drg.SetValue(vRowIndex, "NewColumn", vRows(0).Item("AddressUsageDesc").ToString())
        End If
      Next
    End If
    Return vResult
  End Function

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Try
      Dim vParams As New ParameterList(True)
      vParams.IntegerValue("ContactNumber") = mvContactNumber
      vParams.IntegerValue("AddressNumber") = mvAddressNumber
      vParams.IntegerValue("SelectionSetNumber") = mvMailingInfo.SelectionSet
      vParams.IntegerValue("Revision") = mvMailingInfo.Revision
      vParams("ApplicationName") = mvMailingInfo.MailingTypeCode
      DataHelper.UpdateMailingContactAddress(vParams)

      If drg.CurrentRow > -1 Then
        mvAddressNumber = CInt(drg.GetValue(drg.CurrentRow, "AddressNumber"))
      Else
        mvAddressNumber = 0
      End If

      Me.Close()

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub drg_RowSelected(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pDataRow As System.Int32) Handles drg.RowSelected
    Try
      If drg.RowCount > 0 Then
        cmdOk.Enabled = (drg.GetValue(drg.CurrentDataRow, "AddressNumber").Length > 0)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub
End Class