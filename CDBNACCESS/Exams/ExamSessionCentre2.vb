
Namespace Access

  Partial Public Class ExamSessionCentre
    Inherits CARERecord

    Public Overloads Sub Init(ByVal pExamSessionId As Integer, ByVal pExamCentreId As Integer)
      CheckClassFields()
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ExamSessionCentreFields.ExamSessionId).Name, pExamSessionId)
      vWhereFields.Add(mvClassFields(ExamSessionCentreFields.ExamCentreId).Name, pExamCentreId)
      MyBase.InitWithPrimaryKey(vWhereFields)
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(ExamSessionCentreFields.ExamSessionId).Name, ExamSessionId)
      vWhereFields.Add(mvClassFields(ExamSessionCentreFields.ExamCentreId).Name, ExamCentreId)
      If mvEnv.Connection.GetCount("exam_schedule", vWhereFields) > 0 Then RaiseError(DataAccessErrors.daeReferencedInOtherTable, "an Exam Schedule")
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
    End Sub

  End Class

End Namespace