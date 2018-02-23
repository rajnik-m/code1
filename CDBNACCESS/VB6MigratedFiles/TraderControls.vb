

Namespace Access
  Public Class TraderControls
    Implements System.Collections.IEnumerable

    Private mvCol As New Collection

    Public Function AddFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRS As CDBRecordSet, ByVal pTraderPageType As TraderPage.TraderPageType, ByRef pTabIndex As Integer, ByVal pNextTop As Integer, ByVal pPageFirstIndex As Integer, ByVal pTraderApplication As TraderApplication) As TraderControl
      'pPageFirstIndex = Index number of first control on Trader Page
      Dim vTraderControl As New TraderControl
      Dim vBaseName As String
      Dim vControlCount As Integer
      Dim vCount As Integer
      Dim vKey As String
      Dim vName As String
      Dim vPageCode As String

      vControlCount = mvCol.Count()
      vTraderControl.InitFromRecordSet(pEnv, pRS, pTraderPageType, pTabIndex, pNextTop, vControlCount, (vControlCount - pPageFirstIndex), pTraderApplication)

      pTabIndex = pTabIndex + 1
      If pRS.Fields("control_type").Value = "dtt" Then pTabIndex = pTabIndex + 1
      vPageCode = pRS.Fields("fp_page_type").Value

      If Len(vTraderControl.ParameterName) = 0 Then
        vBaseName = Replace(StrConv(Replace(vTraderControl.AttributeName, "_", " "), VbStrConv.ProperCase), " ", "")
        vKey = vPageCode & vBaseName
        If Exists(vKey) Then
          vCount = 1
          Do
            vCount = vCount + 1
            vName = vBaseName & vCount
            vKey = vPageCode & vName
          Loop While Exists(vKey)
        Else
          vName = vBaseName
        End If
        vTraderControl.ParameterName = vName
      End If

      vKey = vPageCode & vTraderControl.ParameterName
      '  Debug.Print "Key: " & vKey
      mvCol.Add(vTraderControl, CStr(vKey))
      AddFromRecordSet = vTraderControl

    End Function


    Default Public ReadOnly Property Item(ByVal pIndexKey As Integer) As TraderControl
      Get
        Return CType(mvCol.Item(pIndexKey + 1), TraderControl)  '??
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function Exists(ByVal pKey As String) As Boolean
      Return mvCol.Contains(pKey)
    End Function

  End Class
End Namespace
