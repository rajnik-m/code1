Namespace Access

  Public Class CDBDataRow
    Private mvColumns As CollectionList(Of CDBDataColumn)
    Private mvValues() As String
    Private mvIndex As Integer

    Friend Sub New(ByVal pColumns As CollectionList(Of CDBDataColumn), ByVal pIndex As Integer)
      mvColumns = pColumns
      ReDim mvValues(pColumns.Count)
      mvIndex = pIndex
    End Sub

    Public ReadOnly Property Index As Integer
      Get
        Return mvIndex
      End Get
    End Property

    Friend Sub ResetColumnsCount()
      ReDim Preserve mvValues(mvColumns.Count)
    End Sub

    Public ReadOnly Property DoubleItem(ByVal pKey As String) As Double
      Get
        Return DoubleValue(Item(pKey))
      End Get
    End Property

    Public ReadOnly Property LongItem(ByVal pKey As String) As Integer
      Get
        Return IntegerValue(Item(pKey))
      End Get
    End Property

    Public ReadOnly Property IntegerItem(ByVal pKey As String) As Integer
      Get
        Return IntegerValue(Item(pKey))
      End Get
    End Property

    Public Property Item(ByVal pKey As String) As String
      Get
        If Not mvColumns.ContainsKey(pKey) Then RaiseError(DataAccessErrors.daeFieldNotFound, pKey)
        Return mvValues(mvColumns(pKey).Index)
      End Get
      Set(ByVal pValue As String)
        If Not mvColumns.ContainsKey(pKey) Then RaiseError(DataAccessErrors.daeFieldNotFound, pKey)
        mvValues(mvColumns(pKey).Index) = pValue
      End Set
    End Property

    Public Property Item(ByVal pIndex As Integer) As String
      Get
        Return mvValues(pIndex)
      End Get
      Set(ByVal pValue As String)
        mvValues(pIndex) = pValue
      End Set
    End Property

    Public Property BoolItem(ByVal pKey As String) As Boolean
      Get
        If Not mvColumns.ContainsKey(pKey) Then RaiseError(DataAccessErrors.daeFieldNotFound, pKey)
        Return mvValues(mvColumns(pKey).Index) = "Y"
      End Get
      Set(ByVal pValue As Boolean)
        If Not mvColumns.ContainsKey(pKey) Then RaiseError(DataAccessErrors.daeFieldNotFound, pKey)
        If pValue Then
          mvValues(mvColumns(pKey).Index) = "Y"
        Else
          mvValues(mvColumns(pKey).Index) = "N"
        End If
      End Set
    End Property

    Public Sub SetAttended(ByVal pIndex As String)
      Try
        Dim vIndex As Integer = mvColumns(pIndex).Index
        Select Case mvValues(vIndex)
          Case "Y"
            mvValues(vIndex) = ProjectText.String22054 'Yes
          Case "N"
            mvValues(vIndex) = ProjectText.String22055 'No
          Case "A"
            mvValues(vIndex) = ProjectText.String22056 'Apologies
          Case Else
            'Leave it alone may already have been set
        End Select
      Catch ex As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, pIndex)
      End Try
    End Sub

    Public Sub SetChairPerson(ByVal pIndex As String)
      Try
        Dim vIndex As Integer = mvColumns(pIndex).Index
        Select Case mvValues(vIndex)
          Case "C"
            mvValues(vIndex) = ProjectText.String22054 'Yes
          Case Else
            'Leave it alone may already have been set
        End Select
      Catch ex As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, pIndex)
      End Try
    End Sub

    Public Sub SetYNValue(ByVal pIndex As String)
      SetYNValue(pIndex, False, False)
    End Sub

    Public Sub SetYNValue(ByVal pIndex As String, ByVal pNullIsYes As Boolean)
      SetYNValue(pIndex, pNullIsYes, False)
    End Sub

    Public Sub SetYNValue(ByVal pIndex As String, ByVal pNullIsYes As Boolean, ByVal pSetNo As Boolean)
      Dim vYes As String = ProjectText.String15904
      Dim vNo As String = ProjectText.String22055
      Try
        Dim vIndex As Integer = mvColumns(pIndex).Index
        Select Case mvValues(vIndex)
          Case "Y"
            mvValues(vIndex) = vYes
          Case vYes
            'Leave it alone
          Case "N", vNo
            If pSetNo Then
              mvValues(vIndex) = vNo
            Else
              mvValues(vIndex) = ""
            End If
          Case Else
            If pNullIsYes AndAlso mvValues(vIndex) = "" Then
              mvValues(vIndex) = vYes
            Else
              mvValues(vIndex) = ""
            End If
        End Select
      Catch ex As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, pIndex)
      End Try
    End Sub

    Public Sub ChangeSign(ByVal pIndex As String)
      Try
        Dim vIndex As Integer = mvColumns(pIndex).Index
        mvValues(vIndex) = (DoubleValue(mvValues(vIndex)) * -1).ToString("F")
      Catch ex As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, CStr(pIndex))
      End Try
    End Sub

    Public Sub SetCurrentFutureHistoric(ByVal pStatusColumn As String, ByVal pStatusOrderColumn As String)
      Try
        Dim vStatusIndex As Integer = mvColumns(pStatusColumn).Index
        Dim vStatusOrderIndex As Integer = mvColumns(pStatusOrderColumn).Index
        If Item("ValidFrom").Length > 0 AndAlso CDate(Item("ValidFrom")) > Date.Today Then
          mvValues(vStatusIndex) = ProjectText.String23501    'Future
          mvValues(vStatusOrderIndex) = "1"
        ElseIf Item("ValidTo").Length > 0 AndAlso CDate(Item("ValidTo")) < Date.Today Then
          mvValues(vStatusIndex) = ProjectText.String23503    'Historic
          mvValues(vStatusOrderIndex) = "3"
        Else
          mvValues(vStatusIndex) = ProjectText.String23502    'Current
          mvValues(vStatusOrderIndex) = "2"
        End If
        Exit Sub
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, pStatusColumn)
      End Try
    End Sub

    Friend Sub SetDescriptionFromCode(ByVal pCodeColumn As String, ByVal pDescColumn As String, ByVal pData As CollectionList(Of LookupItem))
      Try
        Dim vCodeCol As Integer = mvColumns(pCodeColumn).Index
        Dim vDescCol As Integer = mvColumns(pDescColumn).Index
        If mvValues(vCodeCol).Length > 0 AndAlso pData.ContainsKey(mvValues(vCodeCol)) = True Then
          mvValues(vDescCol) = pData.Item(mvValues(vCodeCol)).LookupDesc
        End If
      Catch vEX As Exception
        RaiseError(DataAccessErrors.daeFieldNotFound, pDescColumn)
      End Try
    End Sub
  End Class
End Namespace