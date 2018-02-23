Namespace Access.BulkMailer

  Public Class BulkMailingProperties

    Private mvEnvironment As CDBEnvironment = Nothing
    Private mvMailingId As Integer = 0
    Private mvMailing As String = String.Empty
    Private mvFromName As String = String.Empty
    Private mvFromAddress As String = String.Empty
    Private mvReplyToAddress As String = String.Empty
    Private mvMailingCode As String = String.Empty
    Private mvMailingIdValid As Boolean = False
    Private mvMailingCodeValid As Boolean = False

    Public Sub New(pEnvironment As CDBEnvironment, pMailingId As Integer, pMailing As String, pFromName As String, pFromAddress As String, pReplyToAddress As String)
      mvEnvironment = pEnvironment
      mvMailingId = pMailingId
      mvMailing = pMailing
      mvFromName = pFromName
      mvFromAddress = pFromAddress
      mvReplyToAddress = pReplyToAddress
      mvMailingIdValid = True
    End Sub

    Public Sub New(pEnvironment As CDBEnvironment, pMailingCode As String, pMailing As String, pFromName As String, pFromAddress As String, pReplyToAddress As String)
      mvEnvironment = pEnvironment
      mvMailingCode = pMailingCode
      mvMailing = pMailing
      mvFromName = pFromName
      mvFromAddress = pFromAddress
      mvReplyToAddress = pReplyToAddress
      mvMailingIdValid = True
    End Sub

    Public ReadOnly Property MailingId As Integer
      Get
        If Not mvMailingIdValid Then
          Dim vWhereClause As New CDBFields
          vWhereClause.Add("mailing", MailingCode, CDBField.FieldWhereOperators.fwoEqual)
          Dim vSql As New SQLStatement(mvEnvironment.Connection, "bulk_mailer_mailing", "mailings", vWhereClause)
          mvMailingId = vSql.GetIntegerValue
          mvMailingIdValid = True
        End If
        Return mvMailingId
      End Get
    End Property

    Public ReadOnly Property Mailing As String
      Get
        Return mvMailing
      End Get
    End Property

    Public ReadOnly Property FromName As String
      Get
        Return mvFromName
      End Get
    End Property

    Public ReadOnly Property FromAddress As String
      Get
        Return mvFromAddress
      End Get
    End Property

    Public ReadOnly Property ReplyToAddress As String
      Get
        Return mvReplyToAddress
      End Get
    End Property

    Public ReadOnly Property MailingCode As String
      Get
        If Not mvMailingCodeValid Then
          Dim vWhereClause As New CDBFields
          vWhereClause.Add("bulk_mailer_mailing", MailingId, CDBField.FieldWhereOperators.fwoEqual)
          Dim vSql As New SQLStatement(mvEnvironment.Connection, "mailing", "mailings", vWhereClause)
          mvMailingCode = vSql.GetValue
          mvMailingCodeValid = True
        End If
        Return mvMailingCode
      End Get
    End Property
  End Class
End Namespace

