

Namespace Access
  Public Class CreditSalesTerms

    Private mvTermsNumber As Integer
    Private mvTermsPeriod As String
    Private mvTermsFrom As String
    Private mvCompany As String
    Private mvUseSalesLedger As Boolean
    Private mvEnv As CDBEnvironment
    Private mvCreditCustomer As CreditCustomer

    Sub Init(ByVal pEnv As CDBEnvironment, ByRef pContactNumber As Integer, ByRef pCompany As String, ByRef pSalesLedgerAccount As String)
      Dim vFoundTerms As Boolean
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      mvUseSalesLedger = pEnv.GetConfigOption("fp_use_sales_ledger")
      mvCreditCustomer = New CreditCustomer
      mvCreditCustomer.Init(pEnv, pContactNumber, pCompany, pSalesLedgerAccount)
      If Not mvCreditCustomer.Existing Then
        RaiseError(DataAccessErrors.daeCreditCustomerMissing, CStr(pContactNumber), pCompany, pSalesLedgerAccount)
      Else
        If Len(mvCreditCustomer.TermsFrom) > 0 And Len(mvCreditCustomer.TermsPeriod) > 0 Then
          mvTermsFrom = mvCreditCustomer.TermsFrom
          mvTermsNumber = IntegerValue(mvCreditCustomer.TermsNumber)
          mvTermsPeriod = mvCreditCustomer.TermsPeriod
          mvCompany = ""
          vFoundTerms = True
        End If
      End If
      If Not vFoundTerms Then
        If pCompany <> mvCompany Then
          vRecordSet = pEnv.Connection.GetRecordSet("SELECT terms_from, terms_number, terms_period FROM company_credit_controls WHERE company = '" & pCompany & "'")
          If vRecordSet.Fetch() = True Then
            mvTermsFrom = vRecordSet.Fields(1).Value
            mvTermsNumber = vRecordSet.Fields(2).IntegerValue
            mvTermsPeriod = vRecordSet.Fields(3).Value
            mvCompany = pCompany
          Else
            RaiseError(DataAccessErrors.daeCompanyCreditControlsMissing, pCompany)
          End If
          vRecordSet.CloseRecordSet()
        End If
      End If
    End Sub

    Public Function PaymentDue(ByRef pTransactionDate As String) As Date
      Dim vRecordSet As CDBRecordSet

      If mvTermsFrom = "I" Or Not mvUseSalesLedger Then
        If mvTermsPeriod = "D" Then
          PaymentDue = DateAdd(Microsoft.VisualBasic.DateInterval.Day, mvTermsNumber, Today)
        Else
          PaymentDue = DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvTermsNumber, Today)
        End If
      Else
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT start_date FROM calendar WHERE start_date " & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, pTransactionDate) & " ORDER BY start_date")
        If vRecordSet.Fetch() = True Then
          If mvTermsPeriod = "D" Then
            PaymentDue = DateAdd(Microsoft.VisualBasic.DateInterval.Day, mvTermsNumber, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vRecordSet.Fields(1).Value)))
          Else
            PaymentDue = DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvTermsNumber, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vRecordSet.Fields(1).Value)))
          End If
        Else
          RaiseError(DataAccessErrors.daeNoFinancialPeriod, pTransactionDate)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Function

    Public ReadOnly Property Company() As String
      Get
        Company = mvCompany
      End Get
    End Property
    Public ReadOnly Property TermsFrom() As String
      Get
        TermsFrom = mvTermsFrom
      End Get
    End Property
    Public ReadOnly Property TermsPeriod() As String
      Get
        TermsPeriod = mvTermsPeriod
      End Get
    End Property
    Public ReadOnly Property TermsNumber() As Integer
      Get
        TermsNumber = mvTermsNumber
      End Get
    End Property

    Public ReadOnly Property CreditCustomer() As CreditCustomer
      Get
        CreditCustomer = mvCreditCustomer
      End Get
    End Property
  End Class
End Namespace
