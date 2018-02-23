Public Class ContactExamCertificatesMenu
  Inherits ContextMenuStrip

  Public Event ReprintSelected(ByVal sender As Object, e As ReprintSelectedEventArgs)
  Public Event RecallRequested(ByVal sender As Object, e As EventArgs)

  Private mvMenuItems As New CollectionList(Of MenuToolbarCommand)

  Public Sub New()
    MyBase.New()
    Dim vRecallItem As New MenuToolbarCommand("_RECALL", CStr("Toggle Recall"), 0)
    vRecallItem.OnClick = Sub(sender As Object, e As EventArgs)
                            RaiseEvent RecallRequested(sender, New EventArgs())
                          End Sub
    Me.Items.Add(vRecallItem.MenuStripItem)

    Dim vMenuItems As DataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamCertReprintTypes, New ParameterList(True, True))
    If vMenuItems IsNot Nothing AndAlso vMenuItems.Tables.Contains("DataRow") Then
      Dim vCommandId As Integer = CInt(Integer.MaxValue / 2)
      For Each vRow As DataRow In vMenuItems.Tables("DataRow").Rows
        Dim vItem As New MenuToolbarCommand(CStr(vRow("ExamCertReprintType")), CStr(vRow("ExamCertReprintTypeDesc")), vCommandId)
        vItem.OnClick = Sub(sender As Object, e As EventArgs)
                          RaiseEvent ReprintSelected(sender, New ReprintSelectedEventArgs(DirectCast(sender, ToolStripMenuItem).Name.Substring(3)))
                        End Sub
        Me.Items.Add(vItem.MenuStripItem)
        mvMenuItems.Add(CStr(vRow("ExamCertReprintType")), vItem)
        vCommandId += 1
      Next
      'MenuToolbarCommand.SetAccessControl(mvMenuItems)
    End If
  End Sub

  Private Sub Menu_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Opening
    SetVisibleItems(e)
  End Sub

  Public Property ItemsEnabled As Boolean = False

  Public Sub SetVisibleItems(ByVal e As System.ComponentModel.CancelEventArgs)
    Dim vCursor As New BusyCursor
    Try
      For Each vItem As MenuToolbarCommand In mvMenuItems
        vItem.SetContextItemVisible(Me, True)
        vItem.SetContextItemEnabled(Me, Me.ItemsEnabled)
      Next vItem
    Catch vException As Exception
      DataHelper.HandleException(vException)
    Finally
      vCursor.Dispose()
    End Try
  End Sub
End Class
