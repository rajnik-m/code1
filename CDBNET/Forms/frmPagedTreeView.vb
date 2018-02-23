Imports CDBNETBiz

Public Class frmPagedTreeView
  Inherits CDBNETCL.PersistentForm


  Public ReadOnly Property DataContext As PagedTreeViewDataContext
    Get
      Return PagedTreeView.DataContext
    End Get
  End Property

  Public Sub Init(context As PagedTreeViewDataContext)
    PagedTreeView.Init(context)
    If context IsNot Nothing Then
      Me.SettingsName = String.Format("f{0}_c{1}", Me.GetType().Name, context.ToString())
      Me.SetSize()
    End If
  End Sub

  Private Sub frmPagedTreeView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    If PagedTreeView.Closing() = False Then
      e.Cancel = True
    End If
  End Sub

  Private Sub PagedTreeView_EntitySelected(sender As Object, entityID As String, entityType As HistoryEntityTypes) Handles PagedTreeView.EntitySelected
    Dim vEntityNumber As Integer
    If Integer.TryParse(entityID, vEntityNumber) Then
      MainHelper.NavigateHistoryItem(entityType, vEntityNumber, False)
    End If
  End Sub
End Class