Imports CDBNETXAML

<ExplorerViewOption(ExplorerViewOptions.SuppressInTabStrip)>
Public Class frmModernMenuViewer
  Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

  End Sub

  Public Property DataContext As Object
    Get
      Return HomeScreen1.DataContext
    End Get
    Set(value As Object)
      HomeScreen1.DataContext = value
    End Set
  End Property

End Class