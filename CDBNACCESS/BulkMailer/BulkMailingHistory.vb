Public Class BulkMailingHistory
  Inherits MailingHistory

  Public Sub New(pMailingId As Integer, pEnv As CDBEnvironment)
    MyBase.New(pEnv)
    Dim vWhereClause As New CDBFields
    vWhereClause.Add("bulk_mailer_mailing", pMailingId, CDBField.FieldWhereOperators.fwoEqual)
    Dim vMailing As String = (New SQLStatement(pEnv.Connection, "mailing", "mailings", vWhereClause)).GetDataTable.Rows(0)("mailing").ToString
    vWhereClause.Clear()
    vWhereClause.Add("mailing", vMailing, CDBField.FieldWhereOperators.fwoEqual)
    InitFromDataRow((New SQLStatement(pEnv.Connection, "mailing,mailing_date,mailing_by,number_in_mailing,mailing_number,mailing_filename,notes,issue_id,number_emails_bounced,number_emails_opened,number_emails_clicked", "mailing_history", vWhereClause)).GetDataTable.Rows(0), False)
    If Not Existing Then
      Dim vParams As New CDBParameters
      vParams.Add("Mailing", vMailing)
      vParams.Add("MailingDate", TodaysDate)
      vParams.Add("MailingBy", mvEnv.User.UserID)
      vParams.Add("NumberInMailing", 0)
      vParams.Add("IssueId", "")
      vParams.Add("NumberEmailsBounced", 0)
      vParams.Add("NumberEmailsOpened", 0)
      vParams.Add("NumberEmailsClicked", 0)
      Create(vParams)
      Save()
    End If
  End Sub

  Public Overloads Property NumberInMailing As Integer
    Get
      Return MyBase.NumberInMailing
    End Get
    Set(value As Integer)
      Dim vParams As New CDBParameters
      vParams.Add("NumberInMailing", value)
      Update(vParams)
      Save()
    End Set
  End Property

  Public Overloads Property NumberEmailsBounced As Integer
    Get
      Return MyBase.NumberEmailsBounced
    End Get
    Set(value As Integer)
      Dim vParams As New CDBParameters
      vParams.Add("NumberEmailsBounced", value)
      Update(vParams)
      Save()
    End Set
  End Property

  Public Overloads Property NumberEmailsOpened As Integer
    Get
      Return MyBase.NumberEmailsOpened
    End Get
    Set(value As Integer)
      Dim vParams As New CDBParameters
      vParams.Add("NumberEmailsOpened", value)
      Update(vParams)
      Save()
    End Set
  End Property

  Public Overloads Property NumberEmailsClicked As Integer
    Get
      Return MyBase.NumberEmailsClicked
    End Get
    Set(value As Integer)
      Dim vParams As New CDBParameters
      vParams.Add("NumberEmailsClicked", value)
      Update(vParams)
      Save()
    End Set
  End Property

End Class
