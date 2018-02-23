

Namespace Access
  Public Class TraderPages
    Implements System.Collections.IEnumerable

    Private mvCol As Collection

    Public Function Add() As TraderPage
      Dim vPage As New TraderPage

      mvCol.Add(vPage)
      Add = vPage
    End Function

    Default Public ReadOnly Property Item(ByVal pIndexKey As Integer) As TraderPage
      Get
        Item = CType(mvCol.Item(pIndexKey), TraderPage)
      End Get
    End Property

    'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
    'Public ReadOnly Property NewEnum() As stdole.IUnknown
    'Get
    'NewEnum = mvCol._NewEnum
    'End Get
    'End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Sub Remove(ByRef pIndexKey As Integer)
      mvCol.Remove(pIndexKey)
    End Sub

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
      mvCol = New Collection
    End Sub
    Public Sub New()
      MyBase.New()
      Class_Initialize_Renamed()
    End Sub

  End Class
End Namespace
