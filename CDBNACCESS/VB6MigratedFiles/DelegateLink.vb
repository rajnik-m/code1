

Namespace Access
  Partial Public Class DelegateLink

    ''' <summary>DO NOT USE - USED BY VB6 MIGRATED CODE ONLY</summary>
    ''' <param name="pEventDelegateNumber"></param>
    ''' <param name="pContactNumber"></param>
    ''' <param name="pRelationship"></param>
    ''' <param name="pValidFrom"></param>
    ''' <param name="pValidTo"></param>
    ''' <param name="pNotes"></param>
    ''' <remarks></remarks>
    Public Overloads Sub Create(ByVal pEventDelegateNumber As Integer, ByVal pContactNumber As Integer, ByVal pRelationship As String, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pNotes As String)
      With mvClassFields
        .Item(DelegateLinkFields.EventDelegateNumber).Value = CStr(pEventDelegateNumber)
        .Item(DelegateLinkFields.ContactNumber).Value = CStr(pContactNumber)
        .Item(DelegateLinkFields.Relationship).Value = pRelationship
        .Item(DelegateLinkFields.ValidFrom).Value = pValidFrom
        .Item(DelegateLinkFields.ValidTo).Value = pValidTo
        If IsDate(pValidTo) Then
          If CDate(pValidTo) < Today Then .Item(DelegateLinkFields.Historical).Bool = True
        End If
        .Item(DelegateLinkFields.Notes).Value = pNotes
      End With
    End Sub

    ''' <summary>DO NOT USE - USED BY VB6 MIGRATED CODE ONLY</summary>
    ''' <param name="pContactNumber"></param>
    ''' <param name="pValidFrom"></param>
    ''' <param name="pValidTo"></param>
    ''' <param name="pNotes"></param>
    ''' <remarks></remarks>
    Public Overloads Sub Update(ByVal pContactNumber As Integer, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pNotes As String)
      With mvClassFields
        .Item(DelegateLinkFields.ContactNumber).Value = CStr(pContactNumber)
        .Item(DelegateLinkFields.ValidFrom).Value = pValidFrom
        .Item(DelegateLinkFields.ValidTo).Value = pValidTo
        If IsDate(pValidTo) Then
          If CDate(pValidTo) < Today Then .Item(DelegateLinkFields.Historical).Bool = True
        End If
        .Item(DelegateLinkFields.Notes).Value = pNotes
      End With
    End Sub
  End Class
End Namespace
