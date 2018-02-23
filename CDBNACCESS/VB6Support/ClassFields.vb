Namespace Access

  Partial Public Class ClassFields

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum AmendmentHistoryCreation
      ahcDefault
      ahcYes
      ahcNo
    End Enum
    Public Enum OrderByDirection
      Ascending
      Descending
    End Enum

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CreateAmendmentHistory(ByVal pEnv As CDBEnvironment, ByVal pAHC As AmendmentHistoryCreation) As Boolean
      Select Case pAHC
        Case AmendmentHistoryCreation.ahcYes
          CreateAmendmentHistory = True
        Case AmendmentHistoryCreation.ahcNo
          CreateAmendmentHistory = False
        Case Else
          Select Case mvTableName
            Case "events", "event_booking_options", "event_contacts", "event_organisers",
                 "event_submissions", "event_venue_bookings", "event_owners", "event_topics",
                 "sessions", "event_personnel", "external_resources", "event_resources",
                 "session_tests", "session_test_results", "session_activities", "sundry_costs",
                 "event_sources", "event_mailings", "event_personnel_tasks"
              CreateAmendmentHistory = pEnv.GetConfigOption("event_amendment_history", False)
            Case "orders"
              CreateAmendmentHistory = pEnv.GetConfigOption("fp_pay_plan_amendment_history")
          End Select
      End Select
    End Function

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearSetValues()
      For Each vClassField As ClassField In Me
        vClassField.SetValueOnly = ""
      Next
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub SetUniqueFieldsFromPrimaryKeys()
      Dim vClassField As ClassField
      For vIndex As Integer = 1 To Me.Count
        vClassField = Item(vIndex)
        If vClassField.PrimaryKey Then SetUniqueField(vIndex)
      Next
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Delete(ByVal pConn As CDBConnection)
      Delete(pConn, Nothing, "", False, 0)
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Delete(ByVal pConn As CDBConnection, ByVal pEnv As CDBEnvironment, ByVal pAmendedBy As String, ByVal pAudit As Boolean)
      Delete(pConn, pEnv, pAmendedBy, pAudit, 0)
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save(ByVal pEnv As CDBEnvironment, ByRef pExisting As Boolean)
      Save(pEnv, pExisting, "", False, 0)
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save(ByVal pEnv As CDBEnvironment, ByRef pExisting As Boolean, ByVal pAmendedBy As String, ByVal pAudit As Boolean)
      Save(pEnv, pExisting, pAmendedBy, pAudit, 0)
    End Sub

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property ItemDataType(ByVal pIndexKey As String) As CDBField.FieldTypes
      Get
        ItemDataType = Item(AttributeName(pIndexKey)).FieldType
      End Get
    End Property

    ''' <summary>
    ''' SHOULD ONLY USED BY VB6 MIGRATEDCODE
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property ItemValue(ByVal pIndexKey As String) As String
      Set(ByVal Value As String)
        Item(AttributeName(pIndexKey)).Value = Value
      End Set
    End Property

    Public Sub VerifyUnique(ByVal pConn As CDBConnection)
      Dim vWhereFields As New CDBFields
      Dim vKeyChanged As Boolean

      For Each vClassField As ClassField In Me
        With vClassField
          If .PrimaryKey Then
            If .ValueChanged Then vKeyChanged = True
            vWhereFields.Add(.Name, .FieldType, .Value)
            If .SpecialColumn Then vWhereFields((vWhereFields.Count)).SpecialColumn = True
          End If
        End With
      Next vClassField
      If vKeyChanged Then
        If pConn.GetCount(mvTableName, vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeDuplicateRecord)
      End If
    End Sub

    Public ReadOnly Property Caption As String
      Get
        Dim vResult As String = Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(Me.DatabaseTableName.Replace("_", " "))
        If vResult.EndsWith("uses") Then
          vResult = vResult.Remove(vResult.Length - 2)
        ElseIf vResult.EndsWith("ies") Then
          vResult = vResult.Remove(vResult.Length - 3)
          vResult += "y"
        ElseIf vResult.EndsWith("s") Then
          vResult = vResult.Remove(vResult.Length - 1)
        End If
        Dim vSpecialWords As New List(Of Tuple(Of String, String)) From
          {
            New Tuple(Of String, String)("matrices", "matrix")
          }
        vSpecialWords.ForEach(Function(vWordPair) vResult = vResult.Replace(vWordPair.Item1, vWordPair.Item2))
        Return vResult
      End Get
    End Property
  End Class

End Namespace
