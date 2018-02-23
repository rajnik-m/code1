Namespace Access

  Public Class FundraisingRequest
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum FundraisingRequestFields
      AllFields = 0
      FundraisingRequestNumber
      ContactNumber
      RequestDate
      RequestDescription
      FundraisingRequestStage
      FundraisingStatus
      FundraisingRequestType
      Source
      TargetAmount
      PledgedAmount
      PledgedDate
      ReceivedAmount
      ReceivedDate
      Notes
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("fundraising_request_number", CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("request_date", CDBField.FieldTypes.cftDate)
        .Add("request_description")
        .Add("fundraising_request_stage")
        .Add("fundraising_status")
        .Add("fundraising_request_type")
        .Add("source")
        .Add("target_amount", CDBField.FieldTypes.cftNumeric)
        .Add("pledged_amount", CDBField.FieldTypes.cftNumeric)
        .Add("pledged_date", CDBField.FieldTypes.cftDate)
        .Add("received_amount", CDBField.FieldTypes.cftNumeric)
        .Add("received_date", CDBField.FieldTypes.cftDate)
        .Add("notes", CDBField.FieldTypes.cftMemo)

        .Item(FundraisingRequestFields.FundraisingRequestNumber).PrimaryKey = True
        .SetControlNumberField(FundraisingRequestFields.FundraisingRequestNumber, "FR")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "fr"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "fundraising_requests"
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
    Public ReadOnly Property FundraisingRequestNumber() As Integer
      Get
        Return mvClassFields(FundraisingRequestFields.FundraisingRequestNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(FundraisingRequestFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property RequestDate() As String
      Get
        Return mvClassFields(FundraisingRequestFields.RequestDate).Value
      End Get
    End Property
    Public ReadOnly Property RequestDescription() As String
      Get
        Return mvClassFields(FundraisingRequestFields.RequestDescription).Value
      End Get
    End Property
    Public ReadOnly Property FundraisingRequestStage() As String
      Get
        Return mvClassFields(FundraisingRequestFields.FundraisingRequestStage).Value
      End Get
    End Property
    Public ReadOnly Property FundraisingStatus() As String
      Get
        Return mvClassFields(FundraisingRequestFields.FundraisingStatus).Value
      End Get
    End Property
    Public ReadOnly Property FundraisingRequestType() As String
      Get
        Return mvClassFields(FundraisingRequestFields.FundraisingRequestType).Value
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Return mvClassFields(FundraisingRequestFields.Source).Value
      End Get
    End Property
    Public ReadOnly Property TargetAmount() As Double
      Get
        Return mvClassFields(FundraisingRequestFields.TargetAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property PledgedAmount() As Double
      Get
        Return mvClassFields(FundraisingRequestFields.PledgedAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property PledgedDate() As String
      Get
        Return mvClassFields(FundraisingRequestFields.PledgedDate).Value
      End Get
    End Property
    Public ReadOnly Property ReceivedAmount() As Double
      Get
        Return mvClassFields(FundraisingRequestFields.ReceivedAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ReceivedDate() As String
      Get
        Return mvClassFields(FundraisingRequestFields.ReceivedDate).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(FundraisingRequestFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(FundraisingRequestFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(FundraisingRequestFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerate Code"

    Private mvTargetChanged As Boolean
    Private mvPreviousTarget As Double
    Private mvChangeReason As String

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvTargetChanged = False
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      mvPreviousTarget = TargetAmount
      MyBase.PreValidateUpdateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      mvTargetChanged = mvExisting AndAlso mvClassFields(FundraisingRequestFields.TargetAmount).ValueChanged
      If mvTargetChanged Then mvChangeReason = pParameterList("ChangeReason").Value
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      If mvTargetChanged Then
        Dim vTarget As New FundraisingRequestTarget(mvEnv)
        Dim vParams As New CDBParameters
        vParams.Add("FundraisingRequestNumber", FundraisingRequestNumber)
        vParams.Add("PreviousTargetAmount", mvPreviousTarget)
        vParams.Add("TargetAmount", TargetAmount)
        vParams.Add("ChangeReason", mvChangeReason)
        vTarget.Create(vParams)
        vTarget.Save(pAmendedBy, pAudit, pJournalNumber)
      End If
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      Dim vTarget As New FundraisingRequestTarget(mvEnv)
      vTarget.DeleteByForeignKey(New CDBField("fundraising_request_number", FundraisingRequestNumber))
    End Sub
#End Region

  End Class
End Namespace