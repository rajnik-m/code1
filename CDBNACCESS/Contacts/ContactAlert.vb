Namespace Access

  Public Class ContactAlert
    Inherits CARERecord


#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ContactAlertFields
      AllFields = 0
      ContactAlert
      ContactAlertDesc
      ContactAlertSql
      ContactAlertMessage
      ContactGroup
      SequenceNumber
      ShowAsDialog
      RgbValue
      ContactAlertType
      ContactAlertMessageType
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_alert").PrefixRequired = True
        .Add("contact_alert_desc")
        .Add("contact_alert_sql", CDBField.FieldTypes.cftMemo)
        .Add("contact_alert_message")
        .Add("contact_group").PrefixRequired = True
        .Add("sequence_number", CDBField.FieldTypes.cftLong)
        .Add("show_as_dialog")
        .Add("rgb_value", CDBField.FieldTypes.cftLong)
        .Add("contact_alert_type")
        .Add("contact_alert_message_type")

        .Item(ContactAlertFields.ContactAlert).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ca"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_alerts"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ContactAlertCode() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlert).Value
      End Get
    End Property
    Public ReadOnly Property ContactAlertDesc() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlertDesc).Value
      End Get
    End Property
    Public ReadOnly Property ContactAlertSql() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlertSql).Value
      End Get
    End Property
    Public ReadOnly Property ContactAlertMessage() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlertMessage).Value
      End Get
    End Property
    Public ReadOnly Property ContactGroup() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactGroup).Value
      End Get
    End Property
    Public ReadOnly Property SequenceNumber() As Integer
      Get
        Return mvClassFields(ContactAlertFields.SequenceNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ShowAsDialog() As String
      Get
        Return mvClassFields(ContactAlertFields.ShowAsDialog).Value
      End Get
    End Property
    Public ReadOnly Property RgbValue() As Nullable(Of Integer)
      Get
        Dim vValue As Nullable(Of Integer) = Nothing
        If String.IsNullOrWhiteSpace(mvClassFields.Item(ContactAlertFields.RgbValue).Value) = False Then
          vValue = mvClassFields(ContactAlertFields.RgbValue).IntegerValue
        End If
        Return vValue
      End Get
    End Property
    Public ReadOnly Property ContactAlertType() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlertType).Value
      End Get
    End Property
    Public ReadOnly Property ContactAlertMessageType() As String
      Get
        Return mvClassFields(ContactAlertFields.ContactAlertMessageType).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactAlertFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactAlertFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Public Methods"
    ''' <summary>
    ''' Returns a row for each alert where the AlertSQL returns a record
    ''' </summary>
    ''' <param name="pDataTable"></param>
    ''' <param name="pContactNumber"></param>
    ''' <remarks></remarks>
    Public Sub GetContactAlerts(ByVal pDataTable As CDBDataTable, ByVal pContactNumber As Integer)
      'Get the list of configured alerts for the contact
      Dim vContactAlerts As List(Of ContactAlert) = GetContactAlerts(pContactNumber)
      'Check if the contact matches any of the criteria for the alerts
      BuildAlertsDataTable(pDataTable, vContactAlerts, pContactNumber)

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbEntityAlerts) Then
        'Get configured entity alerts
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("entity_item_number", CDBField.FieldTypes.cftInteger, pContactNumber)
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "entity_alert_number,entity_alert_message,rgb_value,show_as_dialog,entity_alert_desc,entity_item_number,email_address", "entity_alerts", vWhereFields, "sequence_number")
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
        While vRS.Fetch
          If vRS.Fields.Item("show_as_dialog").Bool Then
            'if ShowAsDialog then display alert along with contact alerts
            Dim vRow As CDBDataRow = pDataTable.AddRow
            vRow.Item("ContactNumber") = pContactNumber.ToString()
            vRow.Item("ContactAlert") = vRS.Fields.Item("entity_alert_number").Value
            vRow.Item("AlertMessageDesc") = vRS.Fields.Item("entity_alert_message").Value
            vRow.Item("RgbAlertMessage") = vRS.Fields.Item("rgb_value").Value
            vRow.Item("ShowAsDialog") = vRS.Fields.Item("show_as_dialog").Value
            vRow.Item("AlertMessageType") = "W"    'Warning message
          Else
            'add an entry in the entity alert items table which can be viewed in the notification form
            Dim vParams As New CDBParameters()
            vParams.Add("EntityAlertNumber", vRS.Fields.Item("entity_alert_number").Value)
            vParams.Add("EntityAlertDesc", vRS.Fields.Item("entity_alert_desc").Value)
            vParams.Add("EntityItemNumber", vRS.Fields.Item("entity_item_number").Value)
            vParams.Add("EntityAlertMessage", vRS.Fields.Item("entity_alert_message").Value)
            vParams.Add("AlertNotified", "N")
            AddEntityAlertItem(vParams)
          End If

          'Check if an email needs to be sent
          If vRS.Fields("email_address").Value.Length > 0 Then
            Dim vEmailJob As New EmailJob(mvEnv)
            vEmailJob.Init()
            Dim vContact As New Contact(mvEnv)
            vContact.Init(vRS.Fields.Item("entity_item_number").IntegerValue)
            vEmailJob.SendEmail(String.Format("{0} ({1})", vRS.Fields.Item("entity_alert_desc").Value, vContact.Name), vRS.Fields.Item("entity_alert_message").Value, vRS.Fields("email_address").Value, vContact.ContactNumber.ToString)
          End If
        End While
        vRS.CloseRecordSet()
      End If
    End Sub

    Public Sub GetContactFinanceAlerts(ByVal pDataTable As CDBDataTable, ByVal pContactNumber As Integer, ByVal pTraderApplicationNumber As Integer)
      'Get the list of configured alerts for the Contact and Trader Application
      Dim vParams As New CDBParameters()
      vParams.Add("ContactNumber", pContactNumber)
      vParams.Add("TraderApplicationNumber", pTraderApplicationNumber)
      Dim vContactAlerts As List(Of ContactAlert) = GetContactFinanceAlerts(vParams)
      BuildAlertsDataTable(pDataTable, vContactAlerts, pContactNumber)
    End Sub

    Private Sub BuildAlertsDataTable(ByVal pDataTable As CDBDataTable, ByVal pContactAlerts As List(Of ContactAlert), ByVal pContactNumber As Integer)
      For Each vAlert As ContactAlert In pContactAlerts
        'Replace the placeholder in the AlertSql with the contact number
        Dim vSQL As String = ReplaceString(vAlert.ContactAlertSql, "#", pContactNumber.ToString())
        Dim vResult As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        With vResult
          'if the AlertSQL returns one or more records then add the alert to the output table
          If .Fetch() Then
            Dim vRow As CDBDataRow = pDataTable.AddRow
            vRow.Item("ContactNumber") = pContactNumber.ToString()
            vRow.Item("ContactAlert") = vAlert.ContactAlertCode
            vRow.Item("AlertMessageDesc") = vAlert.ContactAlertMessage
            Dim vRGBValue As Nullable(Of Integer) = vAlert.RgbValue
            vRow.Item("RgbAlertMessage") = If(vRGBValue.HasValue, vRGBValue.Value.ToString(), String.Empty)
            vRow.Item("ShowAsDialog") = vAlert.ShowAsDialog
            vRow.Item("AlertMessageType") = vAlert.ContactAlertMessageType
          End If
          .CloseRecordSet()
        End With
      Next
    End Sub
#End Region

#Region "Private Methods"

    Private Sub AddEntityAlertItem(ByVal pParams As CDBParameters)
      Dim vAlertItem As New EntityAlertItem(mvEnv)
      vAlertItem.Create(pParams)
      vAlertItem.Save()
    End Sub

    ''' <summary>
    ''' Gets the alerts that are configured for the contact group to which the contact belongs
    ''' </summary>
    ''' <param name="pContactNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetContactAlerts(ByVal pContactNumber As Integer) As List(Of ContactAlert)
      Dim vContactAlerts As New List(Of ContactAlert)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "ca.contact_group", "c.contact_group")
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("c.contact_number", pContactNumber)
      vWhereFields.Add("c.contact_type", CDBField.FieldTypes.cftCharacter, "O", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("ca.contact_alert_type", "C")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "contact_alert,sequence_number", "contact_alerts ca", vWhereFields, "sequence_number", vAnsiJoins)

      Dim vAnsiJoins2 As New AnsiJoins()
      vAnsiJoins2.Add("organisations o", "ca.contact_group", "o.organisation_group")
      Dim vWhereFields2 As New CDBFields()
      vWhereFields2.Add("o.organisation_number", pContactNumber)
      vWhereFields2.Add("ca.contact_alert_type", "C")
      Dim vUnion As New SQLStatement(mvEnv.Connection, "contact_alert,sequence_number", "contact_alerts ca", vWhereFields2, "", vAnsiJoins2)
      vSQLStatement.AddUnion(vUnion)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
      While vRS.Fetch
        Dim vRecord As New ContactAlert(mvEnv)
        vRecord.Init(vRS.Fields("contact_alert").Value)
        vContactAlerts.Add(vRecord)
      End While
      vRS.CloseRecordSet()
      Return vContactAlerts
    End Function

    Private Function GetContactFinanceAlerts(ByVal pParams As CDBParameters) As List(Of ContactAlert)
      Dim vAlert As New ContactAlert(Me.Environment)
      Dim vContactAlerts As New List(Of ContactAlert)

      Dim vTraderApplicationNumber As Integer = pParams.ParameterExists("TraderApplicationNumber").IntegerValue

      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "ca.contact_group", "c.contact_group")
      vAnsiJoins.Add("contact_alert_links cal", "ca.contact_alert", "cal.contact_alert")

      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("c.contact_number", pParams("ContactNumber").IntegerValue)
      If vTraderApplicationNumber > 0 Then
        vWhereFields.Add("cal.trader_application_number", vTraderApplicationNumber)
      End If
      vWhereFields.Add("c.contact_type", CDBField.FieldTypes.cftCharacter, "O", CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("ca.contact_alert_type", "F")
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vAlert.GetRecordSetFields(), "contact_alerts ca", vWhereFields, "sequence_number", vAnsiJoins)

      Dim vAnsiJoins2 As New AnsiJoins()
      vAnsiJoins2.Add("organisations o", "ca.contact_group", "o.organisation_group")
      vAnsiJoins2.Add("contact_alert_links cal", "ca.contact_alert", "cal.contact_alert")
      Dim vWhereFields2 As New CDBFields()
      vWhereFields2.Add("o.organisation_number", pParams("ContactNumber").IntegerValue)
      If vTraderApplicationNumber > 0 Then
        vWhereFields2.Add("cal.trader_application_number", vTraderApplicationNumber)
      End If
      vWhereFields2.Add("ca.contact_alert_type", "F")
      Dim vUnion As New SQLStatement(mvEnv.Connection, vAlert.GetRecordSetFields(), "contact_alerts ca", vWhereFields2, "", vAnsiJoins2)

      vSQLStatement.AddUnionAll(vUnion)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
      While vRS.Fetch
        Dim vRecord As New ContactAlert(mvEnv)
        vRecord.InitFromRecordSet(vRS)
        vContactAlerts.Add(vRecord)
      End While
      vRS.CloseRecordSet()
      Return vContactAlerts
    End Function

#End Region

  End Class
End Namespace
