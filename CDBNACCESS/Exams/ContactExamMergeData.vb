Imports System.Linq
Imports Advanced.Data.Merge

Public Class ContactExamMergeData
  Implements IMergeOperation

  Private ReadOnly mvMasterContact As Contact
  Private ReadOnly mvDuplicateContact As Contact
  Private ReadOnly mvEnvironment As CDBEnvironment

  Public ReadOnly Property MasterContact As Contact
    Get
      Return mvMasterContact
    End Get
  End Property
  Public ReadOnly Property DuplicateContact As Contact
    Get
      Return mvDuplicateContact
    End Get
  End Property
  Public ReadOnly Property Environment As CDBEnvironment
    Get
      Return mvEnvironment
    End Get
  End Property

  Public Sub New(pEnvironment As CDBEnvironment, pMasterContact As Contact, pDuplicateContact As Contact)
    mvMasterContact = pMasterContact
    mvDuplicateContact = pDuplicateContact
    mvEnvironment = pEnvironment
  End Sub
  Public Sub ExecuteOperation() Implements IMergeOperation.ExecuteOperation

    Dim vOperations As IList(Of IMergeOperation) = GetExamMerges()

    vOperations.ToList().ForEach(Sub(vOperation) vOperation.ExecuteOperation())
  End Sub

  Private Function GetExamMerges() As IList(Of IMergeOperation)
    Dim vResult As New List(Of IMergeOperation)

    Dim vSummaryMerges As MergeListOperation(Of Contact, ExamStudentHeader) = LoadMerges(Of ExamStudentHeader)({Contact.ContactFields.ContactNumber})
    vResult.Add(vSummaryMerges)

    Dim vBookingMerges As MergeListOperation(Of Contact, ExamBooking) = LoadMerges(Of ExamBooking)({Contact.ContactFields.ContactNumber})
    vResult.Add(vBookingMerges)

    'Dim vContactCerts As MergeListOperation(Of Contact, ContactExamCert) = LoadMerges(Of ContactExamCert)({Contact.ContactFields.ContactNumber})
    'vResult.Add(vContactCerts)

    'Dim vContactExemptions As MergeListOperation(Of Contact, ExamStudentExemption) = LoadMerges(Of ExamStudentExemption)({Contact.ContactFields.ContactNumber})
    'vResult.Add(vContactExemptions)

    Return vResult
  End Function

  Private Function LoadMerges(Of T As CARERecord)(pFieldIndexes As IEnumerable(Of Integer)) As MergeListOperation(Of Contact, T)
    Dim vOperation As MergeListOperation(Of Contact, T) = Nothing

    Dim vMasterItems As IList(Of T) = Me.MasterContact.GetRelatedList(Of T)(pFieldIndexes)
    Dim vDuplicateItems As IList(Of T) = Me.DuplicateContact.GetRelatedList(Of T)(pFieldIndexes)
    vOperation = New MergeListOperation(Of Contact, T)(Me.MasterContact, vMasterItems, vDuplicateItems)

    Return vOperation
  End Function
End Class
