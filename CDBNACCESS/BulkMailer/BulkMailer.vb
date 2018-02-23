Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Linq

Namespace Access.BulkMailer

  ''' <summary>
  ''' The generic interface implemented by all bulk mailers
  ''' </summary>
  Public MustInherit Class BulkMailer
    Implements IDisposable

    Private mvEnvironment As CDBEnvironment = Nothing
    Private mvLogger As Logger = Nothing

    Protected Sub New(pEnvironment As CDBEnvironment)
      mvEnvironment = pEnvironment
    End Sub
    ''' <summary>
    ''' The mailings available from this bulk mailer.
    ''' </summary>
    ''' <value>A list of available mailings as <see cref="BulkMailing" /> instances.</value>
    ''' <remarks>A bulk mailer mailing is a template that has been set up in the bulk mailer 
    ''' to send to a mailing list.</remarks>
    Public ReadOnly Property Mailings As IList(Of BulkMailing)
      Get
        Dim vMailings As List(Of BulkMailing) = AvailableMailings
        Dim vResult As Dictionary(Of Integer, BulkMailing) = vMailings.ToDictionary(Function(x) x.MailingId)
        If vMailings.Count > 0 Then
          Dim vWhereClause As New CDBFields
          vWhereClause.Add("bulk_mailer_mailing", "", CDBField.FieldWhereOperators.fwoNOT)
          Dim vSql As New SQLStatement(mvEnvironment.Connection, "bulk_mailer_mailing", "mailings", vWhereClause)
          For Each vRow As DataRow In vSql.GetDataTable.Rows
            If vResult.Keys.Contains(CInt(vRow(0))) Then
              vResult.Remove(CInt(vRow(0)))
            End If
          Next vRow
        End If
        Return vResult.Values.ToList.AsReadOnly
      End Get
    End Property

    Protected MustOverride ReadOnly Property AvailableMailings As List(Of BulkMailing)

    ''' <summary>
    ''' The properties of a mailings from the bulk mailer.
    ''' </summary>
    ''' <param name="pMailingId">The id of the mailing to get the properties of.</param>
    ''' <value>A <see cref="BulkMailingProperties" /> item for the mailing.</value>
    Public MustOverride ReadOnly Property MailingProperties(pMailingId As Integer) As BulkMailingProperties

    Public ReadOnly Property MailingProperties(pMailingCode As String) As BulkMailingProperties
      Get
        Dim vResult As BulkMailingProperties = Nothing
        Dim vMailingId As Integer = 0
        If mvEnvironment IsNot Nothing Then
          Dim vWhereClause As New CDBFields
          vWhereClause.Add("mailing", pMailingCode, CDBField.FieldWhereOperators.fwoEqual)
          Dim vSql As New SQLStatement(mvEnvironment.Connection, "bulk_mailer_mailing", "mailings", vWhereClause)
          vMailingId = vSql.GetIntegerValue
        End If
        If vMailingId > 0 Then
          vResult = MailingProperties(vMailingId)
        End If
        Return vResult
      End Get
    End Property

    ''' <summary>
    ''' Initiate sending a mailing to a list of contacts.
    ''' </summary>
    ''' <param name="pContactsFilename">The name of a CSV file containing the contact data.</param>
    ''' <param name="pMailingId">The ID of the mailing to send.</param>
    Public Function SendMailingToList(ByVal pContactsFilename As String, ByVal pMailingId As Integer, ByVal pMailingCode As String, ByVal pSendDate As Date) As Integer
      Dim vTransactionStarted As Boolean = False
      If mvEnvironment IsNot Nothing Then
        vTransactionStarted = mvEnvironment.Connection.StartTransaction()
        Dim vMailing As New Mailing(mvEnvironment)
        vMailing.Init(pMailingCode)
        If vMailing.BulkMailerMailing <> pMailingId Then
          vMailing.BulkMailerMailing = pMailingId
          vMailing.BulkMailerStatisticsDate = TodaysDateAndTime()
          vMailing.Save()
        End If
      End If
      If vTransactionStarted Then
        mvEnvironment.Connection.CommitTransaction()
      End If
      Return MailToList(pContactsFilename, pMailingId, pSendDate, pMailingCode)
    End Function

    ''' <summary>
    ''' Send a mailing to a list of contacts.
    ''' </summary>
    ''' <param name="pContactsFilename">The name of a CSV file containing the contact data.</param>
    ''' <param name="pMailingId">The ID of the mailing to send.</param>
    ''' <remarks>This is the function that actually performs the mailing.</remarks>
    Protected MustOverride Function MailToList(ByVal pContactsFilename As String, ByVal pMailingId As Integer, ByVal pSendDate As Date, ByVal pMailingCode As String) As Integer

    ''' <summary>
    ''' The statistics for a mailing.
    ''' </summary>
    ''' <param name="pMailingId">The ID of the mailing.</param>
    ''' <value>The mailing statistics</value>
    Public MustOverride ReadOnly Property Statistics(ByVal pMailingId As Integer) As BulkMailingStats

    ''' <summary>
    ''' The activity detail for a mailing.
    ''' </summary>
    ''' <param name="pMailingId">The ID of the mailing.</param>
    ''' <param name="pSince">The Earliest dtae and time to get activities for.</param>
    ''' <value>A list of <see cref="BulkMailerActvity" /> for the mailing</value>
    Public MustOverride ReadOnly Property ActivityDetail(ByVal pMailingId As Integer, ByVal pSince As Date) As List(Of BulkMailerActvity)

    ''' <summary>
    ''' The environment.
    ''' </summary>
    ''' <value>The <see cref="CDBEnvironment" /> instance related to this <see cref="BulkMailer"/>.</value>
    Protected ReadOnly Property Environment As CDBEnvironment
      Get
        Return mvEnvironment
      End Get
    End Property

    Protected Function GetContactEmail(ByVal pContactNumber As Integer) As String
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("c.contact_number", pContactNumber, CDBField.FieldWhereOperators.fwoEqual)
      vWhereFields.Add(New CDBField("co.valid_from", CDBField.FieldTypes.cftDate, TodaysDateAndTime, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual))
      vWhereFields.Add(New CDBField("co.valid_to", CDBField.FieldTypes.cftDate, TodaysDateAndTime, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual))
      vWhereFields.Add("d.email", "Y", CDBField.FieldWhereOperators.fwoEqual)
      Dim vJoins As New AnsiJoins
      Dim vJoin As New AnsiJoin("communications co", "co.contact_number", "c.contact_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vJoins.Add(vJoin)
      vJoin = New AnsiJoin("devices d", "d.device", "co.device", AnsiJoin.AnsiJoinTypes.InnerJoin)
      vJoins.Add(vJoin)
      Dim vSql As New SQLStatement(mvEnvironment.Connection, "co.""number""", "contacts c", vWhereFields, "CASE WHEN co.preferred_method = 'Y' THEN 0 WHEN co.device_default = 'Y' THEN 1  WHEN co.device_default = 'N' THEN 2 ELSE 3 END", vJoins)
      vSql.MaxRows = 1
      Return vSql.GetValue
    End Function

    Protected Sub InsertContactEmailing(ByVal pContactNumber As Integer, ByVal pEmailAddress As String, ByVal pMailingNumber As String)
      Dim vContactEmailing As ContactEmailing = ContactEmailing.CreateInstance(mvEnvironment, pContactNumber, pEmailAddress, pMailingNumber)
      vContactEmailing.SetProcessed()
      vContactEmailing.Save()
    End Sub

    Protected ReadOnly Property Logger As Logger
      Get
        If mvLogger Is Nothing Then
          mvLogger = New Logger(Environment, "Bulk_Mailer")
        End If
        Return mvLogger
      End Get
    End Property

#Region "IDisposable Support"
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          Logger.Dispose()
        End If
      End If
      Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

  End Class

End Namespace
