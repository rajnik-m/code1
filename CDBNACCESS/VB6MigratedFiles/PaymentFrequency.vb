

Namespace Access
  Public Class PaymentFrequency

    Public Enum PaymentFrequencyRecordSetTypes 'These are bit values
      pfrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    Private Enum PaymentFrequencyFields
      pffAll = 0
      pffPaymentFrequency
      pffPaymentFrequencyDesc
      pffFrequency
      pffAmendedBy
      pffAmendedOn
      pffInterval
      pffPeriod
      pffOffsetMonths
    End Enum

    Public Enum PaymentFrequencyPeriods
      pfpDays
      pfpMonths
    End Enum

    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Public ReadOnly Property PaymentFrequencyCode() As String
      Get
        PaymentFrequencyCode = mvClassFields.Item(PaymentFrequencyFields.pffPaymentFrequency).Value
      End Get
    End Property
    Public ReadOnly Property PaymentFrequencyDesc() As String
      Get
        PaymentFrequencyDesc = mvClassFields.Item(PaymentFrequencyFields.pffPaymentFrequencyDesc).Value
      End Get
    End Property
    Public ReadOnly Property Frequency() As Integer
      Get
        Frequency = mvClassFields.Item(PaymentFrequencyFields.pffFrequency).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(PaymentFrequencyFields.pffAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(PaymentFrequencyFields.pffAmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property Interval() As Integer
      Get
        Interval = mvClassFields.Item(PaymentFrequencyFields.pffInterval).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Period() As PaymentFrequencyPeriods
      Get
        If mvClassFields.Item(PaymentFrequencyFields.pffPeriod).Value = "D" Then
          Period = PaymentFrequencyPeriods.pfpDays
        Else
          Period = PaymentFrequencyPeriods.pfpMonths
        End If
      End Get
    End Property

    Public ReadOnly Property OffsetMonths() As Integer
      Get
        Return mvClassFields.Item(PaymentFrequencyFields.pffOffsetMonths).IntegerValue
      End Get
    End Property
    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(PaymentFrequencyFields.pffAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As PaymentFrequencyRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(PaymentFrequencyFields.pffPaymentFrequency, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And PaymentFrequencyRecordSetTypes.pfrtAll) = PaymentFrequencyRecordSetTypes.pfrtAll Then
          .SetItem(PaymentFrequencyFields.pffPaymentFrequencyDesc, vFields)
          .SetItem(PaymentFrequencyFields.pffFrequency, vFields)
          .SetItem(PaymentFrequencyFields.pffAmendedBy, vFields)
          .SetItem(PaymentFrequencyFields.pffAmendedOn, vFields)
          .SetItem(PaymentFrequencyFields.pffInterval, vFields)
          .SetItem(PaymentFrequencyFields.pffPeriod, vFields)
          .SetItem(PaymentFrequencyFields.pffOffsetMonths, vFields)
        End If
      End With
    End Sub

    Public Function GetRecordSetFields(ByVal pRSType As PaymentFrequencyRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = PaymentFrequencyRecordSetTypes.pfrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "pf")
      Else
        'no other record set types defined
      End If
      Return vFields
    End Function
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "payment_frequencies"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("payment_frequency")
          .Add("payment_frequency_desc")
          .Add("frequency", CDBField.FieldTypes.cftInteger)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("interval", CDBField.FieldTypes.cftInteger)
          .Add("period")
          .Add("offset_months", CDBField.FieldTypes.cftInteger)
          .Item(PaymentFrequencyFields.pffPaymentFrequency).SetPrimaryKeyOnly()
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pPaymentFrequency As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pPaymentFrequency) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(PaymentFrequencyRecordSetTypes.pfrtAll) & " FROM payment_frequencies WHERE payment_frequency = '" & pPaymentFrequency & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, PaymentFrequencyRecordSetTypes.pfrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Private Sub SetDefaults()
      With mvClassFields
        .Item(PaymentFrequencyFields.pffPaymentFrequency).Value = ""
        .Item(PaymentFrequencyFields.pffFrequency).Value = CStr(1)
        .Item(PaymentFrequencyFields.pffInterval).Value = CStr(1)
        .Item(PaymentFrequencyFields.pffPeriod).Value = "M"
        .Item(PaymentFrequencyFields.pffOffsetMonths).Value = "0"
      End With
    End Sub
    Private Sub SetValid(ByVal pField As PaymentFrequencyFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(PaymentFrequencyFields.pffAmendedOn).Value = TodaysDate()
      mvClassFields.Item(PaymentFrequencyFields.pffAmendedBy).Value = mvEnv.User.Logname
      If mvClassFields.Item(PaymentFrequencyFields.pffOffsetMonths).IntegerValue <> 0 Then
        Dim vMaxOffset As Integer = 0
        If IsOffsetMonthsValid(mvClassFields.Item(PaymentFrequencyFields.pffPeriod).Value, Frequency, Interval, OffsetMonths, vMaxOffset) = False Then
          RaiseError(DataAccessErrors.daePayFrequencyOffsetMonthsInvalid, vMaxOffset.ToString)
        End If
      End If
    End Sub

    ''' <summary>Validate the OffsetMonths value.</summary>
    ''' <param name="pPeriod">PaymentFrequency Period ((D)ays or (M)onths</param>
    ''' <param name="pFrequency">PaymentFrequency Frequency</param>
    ''' <param name="pInterval">PaymentFrequency Interval</param>
    ''' <param name="pOffsetMonths">PaymentFrequency Offset</param>
    ''' <param name="pMaxOffset">Maximum Offset Months value which will be set and returned</param>
    ''' <returns>True if valid, otherwise False</returns>
    Public Shared Function IsOffsetMonthsValid(ByVal pPeriod As String, ByVal pFrequency As Integer, ByVal pInterval As Integer, ByVal pOffsetMonths As Integer, ByRef pMaxOffset As Integer) As Boolean
      'A shared function is used here so that the Offset Months can be validated without initialising the class
      Dim vValid As Boolean = True
      pMaxOffset = 0
      If pPeriod.Equals("M", StringComparison.InvariantCultureIgnoreCase) Then
        pMaxOffset = (pInterval - 1)
        If (pFrequency * pInterval) < 12 Then
          pMaxOffset = (12 - (pFrequency * pInterval))
          If pFrequency = 1 Then pMaxOffset = 0 'Invalid
        ElseIf (pFrequency * pInterval) = 12 Then
          If pFrequency = 12 Then pMaxOffset = 0 'Invalid
        Else
          pMaxOffset = 0    'Greater than 12, always invalid
        End If
      End If
      If (pOffsetMonths > pMaxOffset) OrElse (pOffsetMonths < 0) Then vValid = False

      Return vValid

    End Function

    ''' <summary>Get the offset months as they will be applied to this specific usage.</summary>
    ''' <returns>The calculated offset months as it will be applied.</returns>
    Public Function GetCalculatedOffsetMonths() As Integer
      Dim vOffset As Integer = 0
      If OffsetMonths > 0 _
      OrElse ((Frequency * Interval < 12) AndAlso Frequency > 1 AndAlso Period = PaymentFrequencyPeriods.pfpMonths) Then
        Dim vFreqPlusOffset As Integer = ((Frequency * Interval) + OffsetMonths)
        If vFreqPlusOffset = 12 Then
          'Leave it as it is
        Else
          vOffset = (vFreqPlusOffset - 12)
        End If
      End If

      Return vOffset

    End Function

  End Class
End Namespace
