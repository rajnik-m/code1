Imports System.IO
Imports dotMailer.Sdk
Imports dotMailer.Sdk.AddressBook
Imports dotMailer.Sdk.Campaigns
Imports dotMailer.Sdk.Collections
Imports dotMailer.Sdk.Contacts
Imports dotMailer.Sdk.Objects
Imports dotMailer.Sdk.Objects.Collections
Imports dotMailer.Sdk.Objects.Campaigns

Namespace Access.BulkMailer

  ''' <summary>
  ''' A <see cref="BulkMailer" /> implementation specific to DotMailer.
  ''' </summary>
  Public NotInheritable Class DotMailer
    Inherits BulkMailer

    ''' <summary>
    ''' Initializes a new instance of the <see cref="DotMailer" /> class.
    ''' </summary>
    ''' <param name="pLoginId">The ID used to log in to DotMailer.</param>
    ''' <param name="pPassword">The password used to log in to DotMailer.</param>
    ''' <remarks>Instances of this class must be obtained using <see cref="BulkMailerFactory.GetBulkMailerInstance" />
    ''' and not and not by direct instanciation.  This constructor is declared as 'friend' to </remarks>
    Friend Sub New(ByVal pLoginId As String, ByVal pPassword As String, ByVal pEnvironment As CDBEnvironment)
      MyBase.New(pEnvironment)
      Me.LoginId = pLoginId
      Me.Password = pPassword
    End Sub

    Private Property LoginId As String
    Private Property Password As String

    Private mvMailer As MailerInternals = Nothing
    Private ReadOnly Property Mailer As MailerInternals
      Get
        If mvMailer Is Nothing Then
          mvMailer = New MailerInternals(Me.LoginId, Me.Password)
        End If
        Return mvMailer
      End Get
    End Property

    ''' <summary>
    ''' The available DotMailer campaigns.
    ''' </summary>
    ''' <value>A list of the names and IDs of available DotMailer campaigns</value>
    Protected Overrides ReadOnly Property AvailableMailings As System.Collections.Generic.List(Of BulkMailing)
      Get
        Dim vCampaigns As New List(Of BulkMailing)
        For Each vCampaign As DmCampaign In Me.Mailer.CampaignFactory.ListCampaigns
          vCampaigns.Add(New BulkMailing(vCampaign.Id, vCampaign.Name))
        Next vCampaign
        Return vCampaigns
      End Get
    End Property

    ''' <summary>
    ''' The properties of a bulk mailer mailing.
    ''' </summary>
    ''' <param name="pMailingId">The id of the mailing to get the properties of.</param>
    ''' <value>A <see cref="BulkMailingProperties" /> item for the mailing.</value>
    Public Overrides ReadOnly Property MailingProperties(ByVal pMailingId As Integer) As BulkMailingProperties
      Get
        Dim vResult As BulkMailingProperties = Nothing
        Try
          Dim vCampaign As DmCampaign = Me.Mailer.CampaignFactory.GetCampaign(pMailingId)
          Dim vAccountInfo As DmAccountInfo = Me.Mailer.AccountInfo
          vResult = New BulkMailingProperties(Environment, vCampaign.Id, vCampaign.Name, vCampaign.FromName, "DotMailer", vCampaign.ReplyToEmailAddress)
        Catch ex As DmException
        End Try
        Return vResult
      End Get
    End Property

    ''' <summary>
    ''' Send a campaign to a list of contacts.
    ''' </summary>
    ''' <param name="pContactsFilename">The name of a CSV file containing the contact data.</param>
    ''' <param name="pMailingId">The ID of the DotMailer campaign to send.</param>
    ''' <remarks>The CSV file must contain a column called Email.  Other data is stored by column name, with
    ''' the a custom field for the column being useded on DotMailer if it exsits.  Column names are
    ''' case insensitive.  Columns which do not have a corresponding custom data field are ignored.</remarks>
    Protected Overrides Function MailToList(ByVal pContactsFilename As String, ByVal pMailingId As Integer, ByVal pSendDate As Date, ByVal pMailingCode As String) As Integer
      Dim vAddressBook As DmAddressBook = GetAddressBook(pMailingId)
      Dim vInput As CsvReader = New CsvReader(pContactsFilename)
      Dim vContacts As List(Of DmContact) = New List(Of DmContact)
      Dim vInputSchema As DataTable = vInput.GetReferenceTable
      Dim vLogger As New StreamWriter(Environment.GetLogFileName("logfilename"), False, New UTF8Encoding)

      While vInput.Read()
        Dim vValues(vInput.FieldCount - 1) As String
        vInput.GetValues(vValues)
        Dim vContact As DmContact = CreateContactFromRow(vValues, vInputSchema, pMailingCode)
        If vContact IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(vContact.Email) AndAlso vContact.Email.Contains("@") Then
          vContacts.Add(vContact)
        End If
      End While
      vAddressBook.RemoveAllContacts(False, False)
      If vContacts.Count > 0 Then
        vAddressBook.Contacts.AddRange(vContacts)
        vAddressBook.Contacts.AddToAddressBook(vAddressBook)
        Dim vSendDate As Date = pSendDate
        'Dot Mailer will accept but not send email(s) with a send date in the past 
        'so where DateTime is in past (including 2 minute network transmission buffer) increment to future DateTime
        Dim vSafeSendDateInFuture As Date = DateAndTime.DateAdd(DateInterval.Minute, 2, Now)
        If Date.Compare(vSendDate, vSafeSendDateInFuture) < 0 Then
          vSendDate = vSafeSendDateInFuture
        End If
        Me.Mailer.CampaignFactory.SendCampaignToAddressBooks(Nothing, pMailingId, New List(Of DmAddressBook)({vAddressBook}), vSendDate.ToUniversalTime)
      End If
      Return vAddressBook.Contacts.Count
    End Function

    ''' <summary>
    ''' Gets or creates an address book for the mailing.
    ''' </summary>
    ''' <param name="pMailingId">The ID of the mailing.</param><returns></returns>
    Private Function GetAddressBook(ByVal pMailingId As Integer) As DmAddressBook
      Dim vAddressBooks As IList(Of DmAddressBook) = Me.Mailer.AddressBookFactory.ListAddressBooks
      Dim vAddressBook As DmAddressBook = Nothing
      Dim vBookEnum As IEnumerator(Of DmAddressBook) = vAddressBooks.GetEnumerator
      Dim vCampaignName As String = Me.Mailer.CampaignFactory.GetCampaign(pMailingId).Name
      While vBookEnum.MoveNext AndAlso vAddressBook Is Nothing
        If vBookEnum.Current.Name = vCampaignName Then
          vAddressBook = vBookEnum.Current
        End If
      End While
      If vAddressBook Is Nothing Then
        vAddressBook = Me.Mailer.AddressBookFactory.CreateAddressBook(vCampaignName)
      Else
        vAddressBook.RemoveAllContacts(False, False)
      End If
      Return vAddressBook
    End Function

    ''' <summary>
    ''' Creates the contact from record in the CSV file.
    ''' </summary>
    ''' <param name="pValues">A string array containing the values.</param>
    ''' <param name="pSchema">The schema.</param>
    ''' <returns>The created <see cref="DmContact"/> object.</returns>
    Private Function CreateContactFromRow(ByVal pValues As String(), ByVal pSchema As DataTable, ByVal pMailingCode As String) As DmContact
      Dim vEmail As String = String.Empty
      If pSchema.Columns.Contains("Email") Then
        vEmail = pValues(pSchema.Columns("Email").Ordinal)
      Else
        If pSchema.Columns.Contains("Contact Number") Then
          vEmail = GetContactEmail(Integer.Parse(pValues(pSchema.Columns("Contact Number").Ordinal)))
        End If
      End If
      Dim vContact As DmContact = Nothing
      Try
        vEmail = (New System.Net.Mail.MailAddress(vEmail)).Address
        If Not String.IsNullOrWhiteSpace(vEmail) Then
          vContact = Me.Mailer.ContactFactory.CreateNewDmContact(vEmail)
          Dim vDataFieldDefinitions As DmDataFieldDefinitionCollection = Me.Mailer.ContactFactory.DataFieldDefinitions
          For vFieldOrdinal As Integer = 0 To pValues.Length - 1
            Dim vColumnName As String = pSchema.Columns(vFieldOrdinal).ColumnName.Replace(" ", "_")
            If String.Compare(vColumnName, "FirstName", StringComparison.CurrentCultureIgnoreCase) = 0 Then
              vContact.FirstName = pValues(vFieldOrdinal)
            ElseIf String.Compare(vColumnName, "Label_Name", StringComparison.CurrentCultureIgnoreCase) = 0 Or String.Compare(vColumnName, "FullName", StringComparison.CurrentCultureIgnoreCase) = 0 Then
              vContact.Fullname = pValues(vFieldOrdinal)
            ElseIf String.Compare(vColumnName, "Gender", StringComparison.CurrentCultureIgnoreCase) = 0 Then
              vContact.Gender = pValues(vFieldOrdinal)
            ElseIf String.Compare(vColumnName, "Surname", StringComparison.CurrentCultureIgnoreCase) = 0 Or String.Compare(vColumnName, "Surname", StringComparison.CurrentCultureIgnoreCase) = 0 Then
              vContact.LastName = pValues(vFieldOrdinal)
            ElseIf String.Compare(vColumnName, "Postcode", StringComparison.CurrentCultureIgnoreCase) = 0 Then
              vContact.Postcode = pValues(vFieldOrdinal)
            Else
              If vDataFieldDefinitions.Contains(vColumnName) Then
                If Not vContact.DataFields.Contains(vColumnName) Then
                  vContact.DataFields.Add(vDataFieldDefinitions(vColumnName))
                End If
                vContact(vColumnName) = pValues(vFieldOrdinal)
              End If
            End If
          Next vFieldOrdinal
          InsertContactEmailing(If(pSchema.Columns.Contains("Contact Number"), Integer.Parse(pValues(pSchema.Columns("Contact Number").Ordinal)), 0), vEmail, pMailingCode)
        End If
      Catch ex As Exception
        Logger.LogMessage(Resources.ErrorText.DaeInvalidEmailAddress, New String() {vEmail, If(pSchema.Columns.Contains("Contact Number"), pValues(pSchema.Columns("Contact Number").Ordinal), String.Empty)})
      End Try
      Return vContact
    End Function

    ''' <summary>
    ''' The statistics for a DotMailer campaign.
    ''' </summary>
    ''' <param name="MailingId">The ID of the DotMailer campaign</param>
    ''' <value>A list of campaign statistics</value>
    Public Overrides ReadOnly Property Statistics(ByVal MailingId As Integer) As BulkMailingStats
      Get
        Dim vSummary As DmCampaignSummary = Me.Mailer.CampaignFactory.GetCampaign(MailingId).Summary
        Return New BulkMailingStats(vSummary.DateSent.ToLocalTime, vSummary.NumTotalSent, vSummary.NumTotalHardBounces + vSummary.NumTotalSoftBounces, vSummary.NumTotalOpens, vSummary.NumTotalClicks)
      End Get
    End Property

    ''' <summary>
    ''' The activity detail for a mailing.
    ''' </summary>
    ''' <value>
    ''' A list of <see cref="BulkMailerActvity" /> for the mailing
    ''' </value>
    ''' <param name="pMailingId">The ID of the mailing.</param>
    '''   <param name="pSince">The Earliest dtae and time to get activities for.</param>
    ''' <exception cref="System.NotSupportedException"></exception>
    Public Overrides ReadOnly Property ActivityDetail(ByVal pMailingId As Integer, ByVal pSince As Date) As List(Of BulkMailerActvity)
      Get
        Dim vResult As New List(Of BulkMailerActvity)
        Dim vActivities As DmCampaignContactActivityCollection = Me.Mailer.CampaignFactory.ListCampaignActivitiesSinceDate(pMailingId, pSince.ToUniversalTime)
        For Each vActivity As DmCampaignContactActivity In vActivities
          vResult.Add(New BulkMailerActvity(vActivity.ContactEmail, (vActivity.HardBounced Or vActivity.SoftBounced), vActivity.DateFirstOpened.ToLocalTime, If(vActivity.NumClicks > 0, Today, Nothing), vActivity.Unsubscribed))
        Next vActivity
        Return vResult
      End Get
    End Property

    ''' <summary>
    ''' An internal classs to manage the DotMailer interface.  It implements the required interfaces into the
    ''' DotMailer SDK as lazy-initialised constant properties to avoid unnecessary overheads making repeated 
    ''' web service calls to get these items.
    ''' </summary>
    Private Class MailerInternals

      Private mvDotMailer As DmService = Nothing
      Private mvAccountInfo As DmAccountInfo = Nothing
      Private mvAddressBookFactory As AddressBookFactory = Nothing
      Private mvCampaignFactory As CampaignFactory = Nothing
      Private mvContactFactory As ContactFactory = Nothing

      ''' <summary>
      ''' Initializes a new instance of the <see cref="MailerInternals" /> class.
      ''' </summary>
      ''' <param name="pLoginId">The ID to use when logging in to DotMailer.</param>
      ''' <param name="pPassword">The password to use when logging in to DotMailer.</param>
      ''' <remarks>As standard, if the DotMailer DLL is not found, IIS will intercept the error and just return a
      ''' status 500.  We need a little more information to go back to the user than this, so we intercept the error 
      ''' and throw a <see cref="CareException"/> with the original error as the <see cref="CareException.InnerException"/>.</remarks>
      Friend Sub New(ByVal pLoginId As String, ByVal pPassword As String)
        Try
          mvDotMailer = DmServiceFactory.Create(pLoginId, pPassword)
        Catch vEx As FileNotFoundException
          RaiseError(DataAccessErrors.daeDotMailerSdkNotFound)
        Catch vException As Exceptions.DmServiceNotValidException
          RaiseError(DataAccessErrors.daeDotMailerNotValid, vException.Message)
        End Try
      End Sub

      ''' <summary>
      ''' Gets account information for the DotMailer account.
      ''' </summary>
      ''' <remarks>As standard, if the DotMailer DLL is not found, IIS will intercept the error and just return a
      ''' status 500.  We need a little more information to go back to the user than this, so we intercept the error 
      ''' and throw a <see cref="CareException"/> with the original error as the <see cref="CareException.InnerException"/>.</remarks>
      Friend ReadOnly Property AccountInfo As DmAccountInfo
        Get
          If mvAccountInfo Is Nothing Then
            mvAccountInfo = mvDotMailer.GetAccountInfo
          End If
          Return mvAccountInfo
        End Get
      End Property

      ''' <summary>
      ''' Gets the address book factory for the DotMailer account.
      ''' </summary>
      ''' <remarks>As standard, if the DotMailer DLL is not found, IIS will intercept the error and just return a
      ''' status 500.  We need a little more information to go back to the user than this, so we intercept the error 
      ''' and throw a <see cref="CareException"/> with the original error as the <see cref="CareException.InnerException"/>.</remarks>
      Friend ReadOnly Property AddressBookFactory As AddressBookFactory
        Get
          If mvAddressBookFactory Is Nothing Then
            mvAddressBookFactory = mvDotMailer.AddressBooks
          End If
          Return mvAddressBookFactory
        End Get
      End Property

      ''' <summary>
      ''' Gets the campaign factory for the DotMailer account.
      ''' </summary>
      ''' <remarks>As standard, if the DotMailer DLL is not found, IIS will intercept the error and just return a
      ''' status 500.  We need a little more information to go back to the user than this, so we intercept the error 
      ''' and throw a <see cref="CareException"/> with the original error as the <see cref="CareException.InnerException"/>.</remarks>
      Friend ReadOnly Property CampaignFactory As CampaignFactory
        Get
          If mvCampaignFactory Is Nothing Then
            mvCampaignFactory = mvDotMailer.Campaigns
          End If
          Return mvCampaignFactory
        End Get
      End Property

      ''' <summary>
      ''' Gets the contact factory for the DotMailer account.
      ''' </summary>
      ''' <remarks>As standard, if the DotMailer DLL is not found, IIS will intercept the error and just return a
      ''' status 500.  We need a little more information to go back to the user than this, so we intercept the error 
      ''' and throw a <see cref="CareException"/> with the original error as the <see cref="CareException.InnerException"/>.</remarks>
      Friend ReadOnly Property ContactFactory As ContactFactory
        Get
          If mvContactFactory Is Nothing Then
            mvContactFactory = mvDotMailer.Contacts
          End If
          Return mvContactFactory
        End Get
      End Property

    End Class

  End Class

End Namespace
