Public Class frmDesignTraderApp
  Private mvApplicationNumber As Integer
  Private mvPages As New CollectionList(Of TraderPage)
  Private mvCurPageIndex As Integer
  Private mvMaxPageType As Integer
  Private mvActivePages() As Integer
  Private mvCurActivePageIndex As Integer
  Private mvCurPage As TraderPage
  Private WithEvents mvCustomiseMenu As CustomiseMenu
  Private mvControlsTable As DataTable



  Public Sub New(ByVal pApplicationNumber As Integer)

    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    mvApplicationNumber = pApplicationNumber
    InitializeControl()
    
  End Sub
  Private Sub InitializeControl()
    mvCustomiseMenu = New CustomiseMenu
    epl.ContextMenuStrip = mvCustomiseMenu
    Dim vTD As New TraderApplication(mvApplicationNumber, 0, True)
    mvPages = vTD.Pages
    mvMaxPageType = GetMaxPageType()
    AssignPages()
    mvCurActivePageIndex = 0
    mvCurPageIndex = mvActivePages(mvCurActivePageIndex)
    SetControlsTable()
    SetPage(getPage())
    If mvActivePages.Length <= 1 Then
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
    Else
      cmdPrevious.Enabled = False
    End If

  End Sub

  Private Function IsPageDesignable(ByVal pIndex As Integer) As Boolean
    Select Case mvPages(pIndex).PageType
      Case CareServices.TraderPageType.tpTransactionAnalysisSummary, CareServices.TraderPageType.tpPaymentPlanSummary, CareServices.TraderPageType.tpPurchaseInvoiceSummary, CareServices.TraderPageType.tpPurchaseOrderSummary, _
      CareServices.TraderPageType.tpInvoicePayments, CareServices.TraderPageType.tpMembershipMembersSummary, CareServices.TraderPageType.tpStatementList, CareServices.TraderPageType.tpDummyPage, CareServices.TraderPageType.tpBatchInvoiceSummary
        IsPageDesignable = False
      Case Else
        IsPageDesignable = True
    End Select
  End Function

  Private Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click
    Try
      If mvCurActivePageIndex <= mvActivePages.Length - 1 Then mvCurActivePageIndex = mvCurActivePageIndex + 1
      mvCurPageIndex = mvActivePages(mvCurActivePageIndex)
      SetPage(getPage())
      epl.Refresh()
      If mvCurActivePageIndex = mvActivePages.Length - 1 Then
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
      Else
        cmdPrevious.Enabled = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function GetMaxPageType() As Integer
    Dim vMaxPageType As Integer = 0
    For vCtr As Integer = 0 To mvPages.Count - 1
      If vMaxPageType < mvPages(vCtr).PageType Then
        vMaxPageType = mvPages(vCtr).PageType
      End If
    Next
    Return vMaxPageType
  End Function


  Private Sub SetPage(ByVal pTraderPage As TraderPage)

    Dim vEditPanelInfo As EditPanelInfo
    mvControlsTable.DefaultView.RowFilter = "PageType = '" & pTraderPage.PageType & "'"
    vEditPanelInfo = New EditPanelInfo(pTraderPage, mvControlsTable)

    epl.Init(vEditPanelInfo)
    epl.FormatButtons()
    Me.Text = pTraderPage.PageCode
    mvCurPage = pTraderPage
    mvCustomiseMenu.SetContext(mvApplicationNumber, mvCurPage.PageCode)

    For Each Vcon As Control In epl.Controls
      If TypeName(Vcon) = "TextLookupBox" Then
        Dim vTxt As TextLookupBox
        vTxt = TryCast(Vcon, TextLookupBox)
        vTxt.IsDesign = True
      End If
    Next
    epl.Refresh()
  End Sub

  Private Sub cmdPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrevious.Click
    Try
      If mvCurActivePageIndex >= 0 Then mvCurActivePageIndex = mvCurActivePageIndex - 1
      mvCurPageIndex = mvActivePages(mvCurActivePageIndex)
      epl.Clear()
      SetPage(getPage())
      epl.Refresh()
      If mvCurActivePageIndex = 0 Then
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
      ElseIf mvCurActivePageIndex < mvActivePages.Length - 1 Then
        cmdNext.Enabled = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function getPage() As TraderPage
    For vCtr As Integer = 0 To mvPages.Count - 1
      If mvPages(vCtr).PageType = mvCurPageIndex Then
        Return mvPages(vCtr)
      End If
    Next
    Return Nothing
  End Function
  Private Sub AssignPages()
    For vCtr As Integer = 1 To mvMaxPageType
      For vIncr As Integer = 0 To mvPages.Count - 1
        If vCtr = mvPages(vIncr).PageType Then
          If IsPageDesignable(vIncr) Then
            If Not (mvPages(vIncr).Menu And mvPages(vIncr).First = mvPages(vIncr).Last) Then
              If mvActivePages Is Nothing Then
                ReDim Preserve mvActivePages(0)
              Else
                ReDim Preserve mvActivePages(mvActivePages.Length)
              End If
              mvActivePages(mvActivePages.Length - 1) = vCtr
            End If
          End If
        End If
      Next
    Next
  End Sub

  Private Sub cmdRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRevert.Click
    Try
      If ShowQuestion(QuestionMessages.QmRevertApplication, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        Dim vList As New ParameterList(True)
        vList.IntegerValue("ReportNumber") = mvApplicationNumber
        vList("ReportCode") = mvCurPage.PageCode
        DataHelper.DeleteReportControl(vList)
        SetControlsTable()
        SetPage(mvCurPage)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub UpdatePanel(ByVal pRevert As Boolean) Handles mvCustomiseMenu.UpdatePanel
    Try
      SetControlsTable()
      SetPage(mvCurPage)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
  Private Sub SetControlsTable()
    Dim vDataSet As DataSet = DataHelper.GetTraderApplication(mvApplicationNumber, 0, True)
    If vDataSet.Tables("TraderControls") IsNot Nothing Then
      mvControlsTable = vDataSet.Tables("TraderControls")
    End If
  End Sub
End Class


