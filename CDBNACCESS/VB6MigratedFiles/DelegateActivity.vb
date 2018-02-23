

Namespace Access
  Partial Public Class DelegateActivity

    ''' <summary>DO NOT USE - USED BY VB6 MIGRATED CODE ONLY</summary>
    ''' <param name="pDelegateNumber"></param>
    ''' <param name="pActivity"></param>
    ''' <param name="pActivityValue"></param>
    ''' <param name="pQuantity"></param>
    ''' <param name="pSource"></param>
    ''' <param name="pValidFrom"></param>
    ''' <param name="pValidTo"></param>
    ''' <param name="pActivityDate"></param>
    ''' <param name="pNotes"></param>
    ''' <remarks></remarks>
    Public Overloads Sub Create(ByVal pDelegateNumber As Integer, ByVal pActivity As String, ByVal pActivityValue As String, ByVal pQuantity As String, ByVal pSource As String, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pActivityDate As String, ByVal pNotes As String)
      With mvClassFields
        .Item(DelegateActivityFields.EventDelegateNumber).Value = CStr(pDelegateNumber)
        .Item(DelegateActivityFields.Activity).Value = pActivity
        .Item(DelegateActivityFields.ActivityValue).Value = pActivityValue
        .Item(DelegateActivityFields.Quantity).Value = pQuantity
        .Item(DelegateActivityFields.Source).Value = pSource
        .Item(DelegateActivityFields.ValidFrom).Value = pValidFrom
        .Item(DelegateActivityFields.ValidTo).Value = pValidTo
        .Item(DelegateActivityFields.ActivityDate).Value = pActivityDate
        .Item(DelegateActivityFields.Notes).Value = pNotes
      End With
    End Sub

    ''' <summary>DO NOT USE - USED BY VB6 MIGRATED CODE ONLY</summary>
    ''' <param name="pActivityValue"></param>
    ''' <param name="pQuantity"></param>
    ''' <param name="pSource"></param>
    ''' <param name="pValidFrom"></param>
    ''' <param name="pValidTo"></param>
    ''' <param name="pActivityDate"></param>
    ''' <param name="pNotes"></param>
    ''' <remarks></remarks>
    Public Overloads Sub Update(ByVal pActivityValue As String, ByVal pQuantity As String, ByVal pSource As String, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pActivityDate As String, ByVal pNotes As String)
      With mvClassFields
        .Item(DelegateActivityFields.ActivityValue).Value = pActivityValue
        .Item(DelegateActivityFields.Quantity).Value = pQuantity
        .Item(DelegateActivityFields.Source).Value = pSource
        .Item(DelegateActivityFields.ValidFrom).Value = pValidFrom
        .Item(DelegateActivityFields.ValidTo).Value = pValidTo
        .Item(DelegateActivityFields.ActivityDate).Value = pActivityDate
        .Item(DelegateActivityFields.Notes).Value = pNotes
      End With
    End Sub
  End Class
End Namespace
