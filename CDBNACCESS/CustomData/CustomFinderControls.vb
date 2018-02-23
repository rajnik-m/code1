Public Class CustomFinderControls
  Inherits CollectionList(Of CustomFinderControl)

  Protected mvMaxHeight As Long

  Public Overloads Function AddFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pStartNumberingAt As Integer) As CustomFinderControl
    Dim vCustomFinderControl As New CustomFinderControl(pEnv)
    vCustomFinderControl.InitFromRecordSet(pRecordSet)

    If vCustomFinderControl.ParameterName.Length = 0 Then
      Dim vBaseName As String = vCustomFinderControl.AttributeName
      If pStartNumberingAt > 0 Then vBaseName = vBaseName & pStartNumberingAt
      Dim vName As String
      If MyBase.ContainsKey(vBaseName) Then
        Dim vCount As Integer
        If pStartNumberingAt > 0 Then
          vBaseName = Substring(vBaseName, 0, vBaseName.Length - pStartNumberingAt.ToString.Length)
          vCount = pStartNumberingAt
        Else
          vCount = 1                              'Start with 2
        End If
        Do
          vCount = vCount + 1
          vName = vBaseName & vCount
        Loop While MyBase.ContainsKey(vName)
      Else
        vName = vBaseName
      End If
      vCustomFinderControl.ParameterName = vName
    End If
    If vCustomFinderControl.ControlTop + vCustomFinderControl.ControlHeight > mvMaxHeight Then mvMaxHeight = vCustomFinderControl.ControlTop + vCustomFinderControl.ControlHeight
    MyBase.Add(vCustomFinderControl.ParameterName, vCustomFinderControl)
    Return vCustomFinderControl
  End Function

End Class
