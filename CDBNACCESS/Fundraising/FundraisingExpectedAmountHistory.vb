Namespace Access

  Public Class FundraisingExpectedAmountHistory
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum FundraisingExpectedAmountHistoryFields
      AllFields = 0
      FundraisingRequestNumber
      PreviousExpectedAmount
      ExpectedAmount
      ChangeReason
      IsIncomeAmount
      ChangedBy
      ChangedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("fundraising_request_number", CDBField.FieldTypes.cftLong)
        .Add("previous_expected_amount", CDBField.FieldTypes.cftNumeric)
        .Add("expected_amount", CDBField.FieldTypes.cftNumeric)
        .Add("change_reason")
        .Add("is_income_amount")
        .Add("changed_by")
        .Add("changed_on", CDBField.FieldTypes.cftTime)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "feah"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "fund_expected_amount_history"
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
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.FundraisingRequestNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property PreviousExpectedAmount() As Double
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.PreviousExpectedAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ExpectedAmount() As Double
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.ExpectedAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ChangeReason() As String
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.ChangeReason).Value
      End Get
    End Property
    Public ReadOnly Property IsIncomeAmount() As Boolean
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.IsIncomeAmount).Bool
      End Get
    End Property
    Public ReadOnly Property ChangedBy() As String
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.ChangedBy).Value
      End Get
    End Property
    Public ReadOnly Property ChangedOn() As String
      Get
        Return mvClassFields(FundraisingExpectedAmountHistoryFields.ChangedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerate Code"

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      mvClassFields(FundraisingExpectedAmountHistoryFields.ChangedBy).Value = mvEnv.User.UserID
      mvClassFields(FundraisingExpectedAmountHistoryFields.ChangedOn).Value = TodaysDateAndTime()
    End Sub

#End Region
  End Class
End Namespace
