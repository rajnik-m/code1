Namespace Access

  Public Class EntityAlertItem
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum EntityAlertItemFields
      AllFields = 0
      EntityAlertItemNumber
      EntityAlertNumber
      EntityAlertDesc
      EntityItemNumber
      EntityAlertMessage
      AlertNotified
      CreatedBy
      CreatedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("entity_alert_item_number", CDBField.FieldTypes.cftLong)
        .Add("entity_alert_number", CDBField.FieldTypes.cftLong)
        .Add("entity_alert_desc")
        .Add("entity_item_number", CDBField.FieldTypes.cftLong)
        .Add("entity_alert_message")
        .Add("alert_notified")
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(EntityAlertItemFields.EntityAlertItemNumber).PrimaryKey = True
        .SetControlNumberField(EntityAlertItemFields.EntityAlertItemNumber, "AI")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "eai"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "entity_alert_items"
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
    Public ReadOnly Property EntityAlertItemNumber() As Integer
      Get
        Return mvClassFields(EntityAlertItemFields.EntityAlertItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property EntityAlertNumber() As Integer
      Get
        Return mvClassFields(EntityAlertItemFields.EntityAlertNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property EntityAlertDesc() As String
      Get
        Return mvClassFields(EntityAlertItemFields.EntityAlertDesc).Value
      End Get
    End Property
    Public ReadOnly Property EntityItemNumber() As Integer
      Get
        Return mvClassFields(EntityAlertItemFields.EntityItemNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property EntityAlertMessage() As String
      Get
        Return mvClassFields(EntityAlertItemFields.EntityAlertMessage).Value
      End Get
    End Property
    Public ReadOnly Property AlertNotified() As String
      Get
        Return mvClassFields(EntityAlertItemFields.AlertNotified).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(EntityAlertItemFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(EntityAlertItemFields.CreatedOn).Value
      End Get
    End Property
#End Region

  End Class
End Namespace