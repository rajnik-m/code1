Friend Class FDEDisplayLabel
  Inherits CareFDEControl

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pEditing)
  End Sub

  Friend Sub New(ByVal pType As CareNetServices.FDEControlTypes, ByVal pRow As DataRow, ByVal pInitialSettings As String, ByVal pDefaultSettings As String, ByVal pFDEPageNumber As Integer, ByVal pSequenceNumber As Integer, ByVal pEditing As Boolean)
    MyBase.New(pType, pRow, pInitialSettings, pDefaultSettings, pFDEPageNumber, pSequenceNumber, pEditing)
  End Sub

  Friend Overrides Sub SetDefaults()
    MyBase.SetDefaults()

    Dim vControl As Control = FindControl(epl, "Header")
    If vControl IsNot Nothing Then
      Dim vLabel As TransparentLabel = DirectCast(vControl, TransparentLabel)
      If mvDefaultSettings.Length > 0 Then
        Dim vList As ParameterList = GetParameterListFromSettings(mvDefaultSettings)
        If vList.ContainsKey("FontName") Then
          Dim vFontStyle As FontStyle = CType(vList("FontStyle"), FontStyle)
          Dim vFont As New Font(vList("FontName"), CSng(vList("FontSize")), vFontStyle)
          vLabel.Font = vFont
        End If
      End If
      If mvInitialSettings.Length > 0 Then
        Dim vList As ParameterList = GetParameterListFromSettings(mvInitialSettings)
        If vList.ContainsKey("DisplayText") Then
          Dim vText As String = vList("DisplayText")
          vText = vText.Replace("+", ",")
          vText = vText.Replace("^", "=")
          vLabel.Text = vText
        End If
      End If
    End If
  End Sub
End Class
