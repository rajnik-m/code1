Namespace Access

  Public Class FundraisingPaymentSchedule
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum FundraisingPaymentScheduleFields
      AllFields = 0
      FundraisingRequestNumber
      ScheduledPaymentNumber
      ScheduledPaymentDesc
      PaymentAmount
      DueDate
      FundraisingPaymentType
      FundIncomePaymentType
      ReceivedAmount
      ReceivedDate
      SourceCode
      Notes
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("fundraising_request_number", CDBField.FieldTypes.cftLong)
        .Add("scheduled_payment_number", CDBField.FieldTypes.cftLong)
        .Add("scheduled_payment_desc")
        .Add("payment_amount", CDBField.FieldTypes.cftNumeric)
        .Add("due_date", CDBField.FieldTypes.cftDate)
        .Add("fundraising_payment_type")
        .Add("fund_income_payment_type")
        .Add("received_amount", CDBField.FieldTypes.cftNumeric)
        .Add("received_date", CDBField.FieldTypes.cftDate)
        .Add("source")
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(FundraisingPaymentScheduleFields.ScheduledPaymentNumber).PrimaryKey = True
        .SetControlNumberField(FundraisingPaymentScheduleFields.ScheduledPaymentNumber, "FS")

        .Item(FundraisingPaymentScheduleFields.FundraisingPaymentType).NonUpdatable = True
        .Item(FundraisingPaymentScheduleFields.ReceivedAmount).NonUpdatable = True
        .Item(FundraisingPaymentScheduleFields.ReceivedDate).NonUpdatable = True
        .Item(FundraisingPaymentScheduleFields.SourceCode).NonUpdatable = True
        .Item(FundraisingPaymentScheduleFields.CreatedBy).NonUpdatable = True
        .Item(FundraisingPaymentScheduleFields.CreatedOn).NonUpdatable = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "fps"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "fundraising_payment_schedule"
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
        Return mvClassFields(FundraisingPaymentScheduleFields.FundraisingRequestNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ScheduledPaymentNumber() As Integer
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.ScheduledPaymentNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ScheduledPaymentDesc() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.ScheduledPaymentDesc).Value
      End Get
    End Property
    Public ReadOnly Property PaymentAmount() As Double
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.PaymentAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property DueDate() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.DueDate).Value
      End Get
    End Property
    Public ReadOnly Property FundraisingPaymentType() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.FundraisingPaymentType).Value
      End Get
    End Property
    Public ReadOnly Property FundIncomePaymentType() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.FundIncomePaymentType).Value
      End Get
    End Property
    Public ReadOnly Property ReceivedAmount() As Double
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.ReceivedAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property ReceivedDate() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.ReceivedDate).Value
      End Get
    End Property
    Public ReadOnly Property SourceCode() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.SourceCode).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(FundraisingPaymentScheduleFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Private mvFundraisingRequest As FundraisingRequest
    Private mvPledgedAmountDiff As Double
    Private mvGIKPledgedAmountDiff As Double
    Private mvNumberOfPaymentsDiff As Integer

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return "FundraisingRequestNumber,ScheduledPaymentDesc,PaymentAmount,FundraisingPaymentType"
    End Function

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateCreateParameters(pParameterList)
      If Not pParameterList.ParameterExists("SkipValidation").Bool Then 'For WebServices only
        ValidateParameters(pParameterList("FundraisingRequestNumber").IntegerValue, pParameterList("FundraisingPaymentType").Value)
      End If
    End Sub
    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateCreateParameters(pParameterList)
      If mvFundraisingRequest Is Nothing Then
        mvFundraisingRequest = New FundraisingRequest(mvEnv)
        mvFundraisingRequest.Init(FundraisingRequestNumber)
      End If
      If FundraisingPaymentType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType) Then
        If Not pParameterList.ParameterExists("Automatic").Bool Then  'No need to update Pledged Amount and Number of Payments when creating FPS automatically (FR.CreatePaymentSchedule)
          If mvFundraisingRequest.PledgedAmount > 0 Then mvPledgedAmountDiff = PaymentAmount
          mvNumberOfPaymentsDiff = 1
        End If
      ElseIf mvFundraisingRequest.IsGIKPledged Then 'J1353: Only update GIK Pledged Amount if its already set
        mvGIKPledgedAmountDiff = PaymentAmount
      End If
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateUpdateParameters(pParameterList)
      If Not pParameterList.ParameterExists("SkipValidation").Bool Then 'For WebServices only
        ValidateParameters(pParameterList.ParameterExists("FundraisingRequestNumber").IntegerValue, "")
      End If
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PostValidateUpdateParameters(pParameterList)
      If mvClassFields(FundraisingPaymentScheduleFields.PaymentAmount).ValueChanged Then
        If mvFundraisingRequest Is Nothing Then
          mvFundraisingRequest = New FundraisingRequest(mvEnv)
          mvFundraisingRequest.Init(FundraisingRequestNumber)
        End If
        Dim vAmountDifference As Double = PaymentAmount - DoubleValue(mvClassFields(FundraisingPaymentScheduleFields.PaymentAmount).SetValue)
        If FundraisingPaymentType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType) Then
          If mvFundraisingRequest.PledgedAmount > 0 Then mvPledgedAmountDiff = vAmountDifference
        ElseIf mvFundraisingRequest.IsGIKPledged Then 'J1353: Only update GIK Pledged Amount if its already set
          mvGIKPledgedAmountDiff = vAmountDifference
        End If
      End If
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      If mvPledgedAmountDiff <> 0 Then
        mvFundraisingRequest.PledgedAmount += mvPledgedAmountDiff
      ElseIf mvGIKPledgedAmountDiff <> 0 AndAlso mvFundraisingRequest.IsGIKPledged Then 'J1353: Only update GIK Pledged Amount if its already set
        mvFundraisingRequest.GIKPledgedAmount += mvGIKPledgedAmountDiff
      End If
      If mvNumberOfPaymentsDiff <> 0 Then
        If mvFundraisingRequest.NumberOfPayments = 0 AndAlso mvFundraisingRequest.RequestEndDate.Length = 0 Then mvFundraisingRequest.RequestEndDate = TodaysDate()
        mvFundraisingRequest.NumberOfPayments += mvNumberOfPaymentsDiff
      End If
      If mvFundraisingRequest IsNot Nothing Then mvFundraisingRequest.Save(pAmendedBy, pAudit, pJournalNumber)
    End Sub

    Private Function ValidateParameters(ByVal pFundraisingRequestNumber As Integer, ByRef pFundraisingPaymentType As String) As Boolean
      If pFundraisingRequestNumber > 0 Then
        If pFundraisingRequestNumber <> FundraisingRequestNumber Then RaiseError(DataAccessErrors.daeInvalidParameter, "FundraisingRequestNumber")
        mvFundraisingRequest = New FundraisingRequest(mvEnv)
        mvFundraisingRequest.Init(pFundraisingRequestNumber)
        Dim vDefaultStatus As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundStatus)
        Dim vCanMaintain As Boolean = (mvFundraisingRequest.PledgedAmount > 0 OrElse mvFundraisingRequest.ExpectedAmount > 0 _
                                      OrElse mvFundraisingRequest.GIKPledgedAmount > 0 OrElse mvFundraisingRequest.GIKExpectedAmount > 0) _
                                      AndAlso vDefaultStatus.Length > 0 AndAlso vDefaultStatus = mvFundraisingRequest.FundraisingStatus
        If vCanMaintain Then
          If mvFundraisingRequest.PledgedAmount = 0 AndAlso mvFundraisingRequest.ExpectedAmount = 0 AndAlso pFundraisingPaymentType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType) Then
            RaiseError(DataAccessErrors.daeInvalidParameter, "FundraisingPaymentType")
          End If
        Else
          RaiseError(DataAccessErrors.daeCannotAddUpdateFundPaySchedule, mvFundraisingRequest.FundraisingRequestNumber.ToString)
        End If
        Return vCanMaintain
      Else
        RaiseError(DataAccessErrors.daeParameterNotFound, "FundraisingRequestNumber")
      End If
    End Function

    Public Sub AllocateAmount(ByVal pAmount As Double, ByVal pDate As String, ByVal pSourceCode As String)
      'Called from Batch.ProcessProduct
      mvClassFields(FundraisingPaymentScheduleFields.ReceivedAmount).DoubleValue = mvClassFields(FundraisingPaymentScheduleFields.ReceivedAmount).DoubleValue + pAmount
      mvClassFields(FundraisingPaymentScheduleFields.ReceivedDate).Value = pDate
      mvClassFields(FundraisingPaymentScheduleFields.SourceCode).Value = pSourceCode
      Save()
      Dim vWhereFields As New CDBFields(New CDBField("fundraising_request_number", FundraisingRequestNumber))
      Dim vUpdateFields As New CDBFields
      With vUpdateFields
        If FundraisingPaymentType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType) Then
          .Add("received_amount", CDBField.FieldTypes.cftLong, mvEnv.Connection.DBIsNull("received_amount", "0") & " + " & pAmount.ToString)
          .Add("received_date", CDBField.FieldTypes.cftDate, pDate)
        Else
          .Add("total_gik_received_amount", CDBField.FieldTypes.cftLong, mvEnv.Connection.DBIsNull("total_gik_received_amount", "0") & " + " & pAmount.ToString)
          .Add("latest_gik_received_date", CDBField.FieldTypes.cftDate, pDate)
        End If
      End With
      mvEnv.Connection.UpdateRecords("fundraising_requests", vUpdateFields, vWhereFields, False)
    End Sub

    Public Sub AddPaymentLink(ByVal pParams As CDBParameters)
      Dim vFPH As New FundraisingPaymentHistory(mvEnv)
      vFPH.Create(pParams)
      vFPH.Save(mvEnv.User.Logname)
      AllocateAmount(pParams("Amount").DoubleValue, pParams("TransactionDate").Value, pParams("Source").Value)
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      'before delete FPS record get all action_numbers on linked fundraising action records 
      'for deleted fundraising_payment_schedule for non-completed actions
      Dim vActionNumbers As String = ""
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("a.action_number", CDBField.FieldTypes.cftLong, "fa.action_number")
      vWhereFields.Add("scheduled_payment_number", ScheduledPaymentNumber)
      vWhereFields.Add("completed_on", CDBField.FieldTypes.cftDate, "")
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "a.action_number", "fundraising_actions fa,actions a", vWhereFields).GetRecordSet
      While vRecordSet.Fetch()
          If vActionNumbers.Length > 0 Then vActionNumbers = vActionNumbers & ","
          vActionNumbers = vActionNumbers & vRecordSet.Fields(1).Value
      End While
      vRecordSet.CloseRecordSet()
      If vActionNumbers.Length > 0 Then
        Dim vFields As New CDBFields
        Dim vInTrans As Boolean
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vInTrans = True
        End If
        vFields.Add("action_number", CDBField.FieldTypes.cftLong, vActionNumbers, CDBField.FieldWhereOperators.fwoIn)
        mvEnv.Connection.DeleteRecords("actions", vFields, False)
        mvEnv.Connection.DeleteRecords("contact_actions", vFields, False)
        mvEnv.Connection.DeleteRecords("organisation_actions", vFields, False)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataFundraisingPayments) Then
          mvEnv.Connection.DeleteRecords("fundraising_actions", vFields, False)
        End If
        If vInTrans Then mvEnv.Connection.CommitTransaction()
      End If
      MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)

      'Update Fundraising Request Pledged/GIK Pledged Amount/Number Of Payments
      If mvFundraisingRequest Is Nothing Then
        mvFundraisingRequest = New FundraisingRequest(mvEnv)
        mvFundraisingRequest.Init(FundraisingRequestNumber)
      End If
      Dim vUpdate As Boolean = True
      If FundraisingPaymentType = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefaultFundPayType) Then
        If mvFundraisingRequest.PledgedAmount > 0 Then mvFundraisingRequest.PledgedAmount -= PaymentAmount
        mvFundraisingRequest.NumberOfPayments -= 1
      ElseIf mvFundraisingRequest.GIKPledgedAmount > 0 Then 'Do not set the amount if already set to 0
        mvFundraisingRequest.GIKPledgedAmount -= PaymentAmount
        If mvFundraisingRequest.GIKPledgedAmount < 0 Then mvFundraisingRequest.GIKPledgedAmount = 0 'Always set the amount to zero if negative
      Else
        vUpdate = False
      End If
      If vUpdate Then mvFundraisingRequest.Save(pAmendedBy, pAudit, pJournalNumber)

      Dim vWhereField As New CDBField("scheduled_payment_number", ScheduledPaymentNumber)
      Dim vFPH As New FundraisingPaymentHistory(mvEnv)
      vFPH.DeleteByForeignKey(vWhereField)

      Dim vUpdateFields As New CDBFields
      vUpdateFields.Add("scheduled_payment_number", "")
      mvEnv.Connection.UpdateRecords("fundraising_actions", vUpdateFields, New CDBFields(vWhereField), False)
    End Sub

#End Region
  End Class
End Namespace
