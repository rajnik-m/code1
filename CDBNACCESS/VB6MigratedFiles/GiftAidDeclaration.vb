Imports System.Linq
Imports CARE.Access.ClassFields
Imports Advanced.LanguageExtensions

Namespace Access
  Public Class GiftAidDeclaration
    Implements IDbLoadable, IDbSelectable

    Public Enum GiftAidDeclarationRecordSetTypes 'These are bit values
      gadrtAll = &HFFFFS
      'ADD additional recordset types here
      gadrtNumbers = 1
      gadrtCancel = 2
      gadrtType = 4
      gadrtRemainder = 8
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum GiftAidDeclarationFields
      gadfAll = 0
      gadfDeclarationNumber
      gadfContactNumber
      gadfDeclarationDate
      gadfDeclarationType
      gadfSource
      gadfConfirmedOn
      gadfMethod
      gadfStartDate
      gadfEndDate
      gadfNotes
      gadfBatchNumber
      gadfTransactionNumber
      gadfPaymentPlanNumber
      gadfCancellationReason
      gadfCancelledBy
      gadfCancelledOn
      gadfAmendedBy
      gadfAmendedOn
      gadfCancellationSource
      gadfCreatedBy
      gadfCreatedOn
      gadPreviousCancelReason
    End Enum

    Public Enum GiftAidDeclarationTypes
      gadtDonation = 1
      gadtMember = 2
      gadtAll = 3
    End Enum

    Public Enum GiftAidDeclarationMethods
      gadmOral
      gadmWritten
      gadmAny
      gadmElectronic
    End Enum

    Private Enum DLUFields
      dfCDNumber = 1
      dfContactNumber
      dfBatchNumber
      dfTransactionNumber
      dfLineNumber
      dfDecOrCovNumber
      dfNetAmount
    End Enum

    Private Enum GACharityTaxStatus
      gaCTSCompany = 0
      gaCTSTrust
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvCharityTaxStatus As GACharityTaxStatus
    Private mvContact As Contact
    Private mvCAFPayMethod As String
    Private mvGiftAidMinimum As Double
    Private mvPrevTransDate As Date
    Private mvNewDluFromCov As Boolean
    Private mvGiftAidStartDate As String = ""
    Private mvGAPayPlanDecs As Boolean
    Private mvGATransDecs As Boolean
    Private mvAllGADecs As Collection
    Private mvTaxClaimsChecked As Boolean
    Private mvFirstClaimPayDate As String
    Private mvLastClaimPayDate As String
    Private mvGADControlsExists As Boolean = True

    Private mvAmendedValid As Boolean
    Private mvDataInfoGACreated As Boolean
    Private mvProcPaymentPlans As CDBCollection

    Public Sub New()

    End Sub

    Public Sub New(pEnv As CDBEnvironment)
      Me.Environment = pEnv
    End Sub

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        Me.ClassFields = New ClassFields
        With Me.ClassFields
          .DatabaseTableName = "gift_aid_declarations"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("declaration_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("declaration_date", CDBField.FieldTypes.cftDate)
          .Add("declaration_type")
          .Add("source")
          .Add("confirmed_on", CDBField.FieldTypes.cftDate)
          .Add("method")
          .Add("start_date", CDBField.FieldTypes.cftDate)
          .Add("end_date", CDBField.FieldTypes.cftDate)
          .Add("notes", CDBField.FieldTypes.cftMemo)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
          .Add("transaction_number", CDBField.FieldTypes.cftLong)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("cancellation_source")
          .Add("created_by")
          .Add("created_on", CDBField.FieldTypes.cftDate)
        End With

        Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).SetPrimaryKeyOnly()
        mvDataInfoGACreated = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidDecCreatedBy)
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfCreatedBy).InDatabase = mvDataInfoGACreated
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfCreatedOn).InDatabase = mvDataInfoGACreated

        mvCAFPayMethod = mvEnv.GetConfig("pm_caf")
        mvGiftAidMinimum = Val(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlCVGiftAidMinimum))
        mvNewDluFromCov = mvEnv.GetConfigOption("ga_dlu_from_op_cov", False)

        If IsDate(mvEnv.GetConfig("ga_operational_claim_date")) Then mvGiftAidStartDate = CDate(mvEnv.GetConfig("ga_operational_claim_date")).ToString(CAREDateFormat)
        If mvGiftAidStartDate.Length = 0 Then mvGiftAidStartDate = DateSerial(2000, 4, 6).ToString(CAREDateFormat)

        If mvGADControlsExists Then
          Select Case UCase(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGACharityTaxStatus))
            Case "T"
              mvCharityTaxStatus = GACharityTaxStatus.gaCTSTrust
            Case Else '"C"
              mvCharityTaxStatus = GACharityTaxStatus.gaCTSCompany
          End Select
        End If

      Else
        Me.ClassFields.ClearItems()
      End If
      mvAmendedValid = False
      mvExisting = False
      mvGAPayPlanDecs = False
      mvGATransDecs = False

      mvContact = New Contact(mvEnv)
      mvContact.Init()
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      Me.DeclarationDate = TodaysDate()
      Me.DeclarationTypeCode = Me.DeclarationTypeCodeFromType(GiftAidDeclarationTypes.gadtAll)
    End Sub

    Private Sub SetValid(ByRef pField As GiftAidDeclarationFields)
      'Add code here to ensure all values are valid before saving
      If pField = GiftAidDeclarationFields.gadfAll And Not mvAmendedValid Then
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedOn).Value = TodaysDate()
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedBy).Value = mvEnv.User.UserID
      End If
      If pField = GiftAidDeclarationFields.gadfAll And Len(Me.ClassFields(GiftAidDeclarationFields.gadfCreatedBy).Value) = 0 And mvExisting = False Then
        'For new record, set CreatedBy/On
        Me.ClassFields(GiftAidDeclarationFields.gadfCreatedOn).Value = TodaysDate()
        Me.ClassFields(GiftAidDeclarationFields.gadfCreatedBy).Value = mvEnv.User.UserID
      End If
      If Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value) = 0 Then
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value = CStr(mvEnv.GetControlNumber("GD"))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As GiftAidDeclarationRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = GiftAidDeclarationRecordSetTypes.gadrtAll Then
        If Me.ClassFields Is Nothing Then InitClassFields()
        vFields = Me.ClassFields.FieldNames(mvEnv, "gad")
      Else
        vFields = "declaration_number"
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtCancel) > 0 Then
          If Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationReason).InDatabase Then vFields = vFields & ",cancellation_reason,cancelled_by,cancelled_on"
          If Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationSource).InDatabase Then vFields = vFields & ",cancellation_source"
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtNumbers) > 0 Then
          vFields = vFields & ",contact_number"
          If Me.ClassFields.Item(GiftAidDeclarationFields.gadfBatchNumber).InDatabase Then vFields = vFields & ",batch_number,transaction_number,order_number"
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtType) > 0 Then
          vFields = vFields & ",declaration_type,declaration_date,start_date,end_date,method,confirmed_on"
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtRemainder) > 0 Then
          vFields = vFields & ",source,notes,amended_by,amended_on"
        End If
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pDeclarationNumber As Integer = 0, Optional ByVal pRaiseNoGAControlError As Boolean = True)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pRaiseNoGAControlError = False AndAlso mvEnv.Connection.GetCount("gift_aid_controls", Nothing) = 0 Then mvGADControlsExists = False

      If pDeclarationNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(GiftAidDeclarationRecordSetTypes.gadrtAll) & " FROM gift_aid_declarations WHERE declaration_number = " & pDeclarationNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, GiftAidDeclarationRecordSetTypes.gadrtAll)
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

    Friend Sub InitFromContact(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, Optional ByVal pTransactionDate As String = "", Optional ByVal pPlanNumber As Integer = 0, Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pRaiseNoGAControlError As Boolean = True)
      Dim vContact As New Contact(mvEnv)
      Dim vContactLink As ContactLink
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vSQLRestrict As String = ""
      Dim vContactNos As String = ""

      mvEnv = pEnv
      If pRaiseNoGAControlError = False AndAlso mvEnv.Connection.GetCount("gift_aid_controls", Nothing) = 0 Then mvGADControlsExists = False
      If mvGADControlsExists Then
        vContact.InitRecordSetType(mvEnv, Contact.ContactRecordSetTypes.crtName, pContactNumber)
        If vContact.ContactType = Contact.ContactTypes.ctcJoint Then
          For Each vContactLink In vContact.GetJointLinks(True)
            vContactNos = vContactNos & "," & vContactLink.ContactNumber2
          Next vContactLink
        End If
        If vContactNos.Length > 0 Then
          vContactNos = " IN (" & pContactNumber & vContactNos & ")"
        Else
          vContactNos = " = " & pContactNumber
        End If

        vSQL = "SELECT " & GetRecordSetFields(GiftAidDeclarationRecordSetTypes.gadrtAll) & " FROM gift_aid_declarations gad WHERE contact_number" & vContactNos
        If pTransactionDate.Length > 0 Then vSQL = vSQL & " AND start_date" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pTransactionDate) & " AND (end_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pTransactionDate) & " OR end_date IS NULL )"
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason).Length > 0 Then vSQL = vSQL & " AND (cancellation_reason IS NULL or cancellation_reason <> '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason) & "')"
        'Exclude any Declarations that are cancelled on or after the transaction date
        vSQL = vSQL & " AND (cancelled_on " & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, pTransactionDate) & " OR cancelled_on IS NULL)"

        If pBatchNumber > 0 Then vSQLRestrict = " (batch_number = " & pBatchNumber & " AND transaction_number = " & pTransactionNumber & " AND order_number IS NULL)"
        If pPlanNumber > 0 Then
          vSQLRestrict = " AND ((order_number = " & pPlanNumber & " AND batch_number IS NULL) OR" & vSQLRestrict & ")"
          If pBatchNumber = 0 Then vSQLRestrict = Replace(vSQLRestrict, "OR)", ")")
        ElseIf pBatchNumber > 0 Then
          vSQLRestrict = " AND " & vSQLRestrict
        End If

        '1st - Check for GAD linked to a specific PaymentPlan or Transaction
        vRS = mvEnv.Connection.GetRecordSet(vSQL & vSQLRestrict)
        If vRS.Fetch() = True Then InitFromRecordSet(mvEnv, vRS, GiftAidDeclarationRecordSetTypes.gadrtAll)
        vRS.CloseRecordSet()

        If DeclarationNumber = 0 And Len(vSQLRestrict) > 0 Then
          '2nd - Nothing found so check for a GAD that is not linked
          vSQLRestrict = " AND order_number IS NULL AND batch_number IS NULL"
          vRS = mvEnv.Connection.GetRecordSet(vSQL & vSQLRestrict)
          If vRS.Fetch() = True Then InitFromRecordSet(mvEnv, vRS, GiftAidDeclarationRecordSetTypes.gadrtAll)
          vRS.CloseRecordSet()
        End If
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As GiftAidDeclarationRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With Me.ClassFields
        'Always include the primary key attributes
        .SetItem(GiftAidDeclarationFields.gadfDeclarationNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtCancel) > 0 Then
          .SetOptionalItem(GiftAidDeclarationFields.gadfCancellationReason, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfCancelledBy, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfCancelledOn, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfCancellationSource, vFields)
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtNumbers) > 0 Then
          .SetItem(GiftAidDeclarationFields.gadfContactNumber, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfBatchNumber, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfTransactionNumber, vFields)
          .SetOptionalItem(GiftAidDeclarationFields.gadfPaymentPlanNumber, vFields)
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtType) > 0 Then
          .SetItem(GiftAidDeclarationFields.gadfDeclarationDate, vFields)
          .SetItem(GiftAidDeclarationFields.gadfDeclarationType, vFields)
          .SetItem(GiftAidDeclarationFields.gadfConfirmedOn, vFields)
          .SetItem(GiftAidDeclarationFields.gadfMethod, vFields)
          .SetItem(GiftAidDeclarationFields.gadfStartDate, vFields)
          .SetItem(GiftAidDeclarationFields.gadfEndDate, vFields)
        End If
        If (pRSType And GiftAidDeclarationRecordSetTypes.gadrtRemainder) > 0 Then
          .SetItem(GiftAidDeclarationFields.gadfSource, vFields)
          .SetItem(GiftAidDeclarationFields.gadfNotes, vFields)
          .SetItem(GiftAidDeclarationFields.gadfAmendedBy, vFields)
          .SetItem(GiftAidDeclarationFields.gadfAmendedOn, vFields)
        End If
      End With

      mvContact.InitRecordSetType(mvEnv, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress Or Contact.ContactRecordSetTypes.crtDetail, ContactNumber)

    End Sub

    Public Sub Save()
      'Save with defaults
      Save(New SaveOptions())
    End Sub

    Public Sub Save(pSaveOptions As SaveOptions)
      Dim vTransaction As Boolean
      Dim vWhereFields As New CDBFields
      Dim vCreateLines As Boolean
      Dim vJournalOperation As JournalOperations
      Dim vEndDateChanged As Boolean
      Dim vStartDateChanged As Boolean

      SetValid(GiftAidDeclarationFields.gadfAll)

      Dim vCancelInfo As CancellationInfo = pSaveOptions.CancellationInfo
      If vCancelInfo IsNot Nothing AndAlso vCancelInfo.IsValid Then
        If mvExisting Then
          RaiseError(DataAccessErrors.daeCannotCancelPreviousGADOnUpdate)
        End If
        CancelPreviousDeclaration(vCancelInfo)
      End If

      If mvExisting Then
        If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).SetValue), CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) <> 0 Then
          vCreateLines = True
        End If
        If Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then
          If Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue), CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) <> 0 Then
              vCreateLines = True
            End If
          Else
            vCreateLines = True
          End If
        Else
          If Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then
            vCreateLines = True
          End If
        End If
        If CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value) <> CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue) Then
          vStartDateChanged = True
        End If
        If Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationType).SetValue <> Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationType).Value Then
          vCreateLines = True
        End If
      Else
        vCreateLines = True
      End If

      'Decide whether we need to check batch/order numbers
      If vCreateLines Then LinkedDeclarationChecks()

      'Try and figure out whether we really need to check for claim adjustments
      If HasTaxClaims(False) Then
        'Only need to check for claim adjustments if we have some tax claims
        If mvExisting Then
          '(1) Deal with changed StartDate
          If vStartDateChanged Then
            If CDate(StartDate) > CDate(EarliestClaimedPaymentDate) Then
              'StartDate is after the earliest claimed TransactionDate
            ElseIf (CDate(EarliestClaimedPaymentDate) > CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).SetValue)) Then
              'Original StartDate was after the earliest claimed TransactionDate
            Else
              vStartDateChanged = False
            End If
          End If
          '(2) Check for and deal with changed EndDate
          vEndDateChanged = False
          If IsDate(EndDate) Then
            'EndDate is set
            If CDate(LatestClaimedPaymentDate) >= CDate(EndDate) Then
              'EndDate has been moved to be before the latest claimed TransactionDate
              vEndDateChanged = True
            Else
              'EndDate is on/after the latest claimed TransactionDate, so generally no adjustments required
              'But (Due to old data) original Enddate may have been before the latest claimed TransactionDate
              If IsDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) Then
                If CDate(LatestClaimedPaymentDate) > CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) Then
                  vEndDateChanged = True
                End If
              End If
            End If
          Else
            'No EndDate set
            If IsDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) Then
              'Although we no longer support changing GAD's such that adjustments need adjusting, we do need to support existing GAD's in this state
              If CDate(LatestClaimedPaymentDate) > CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) Then
                vEndDateChanged = True
              End If
            End If
          End If
        Else
          vStartDateChanged = False
          vEndDateChanged = False
        End If
      Else
        'No tax claims
        vStartDateChanged = False
        vEndDateChanged = False
      End If

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      If mvExisting Then
        If vCreateLines = True _
        AndAlso (mvClassFields.Item(GiftAidDeclarationFields.gadfStartDate).ValueChanged = True OrElse mvClassFields.Item(GiftAidDeclarationFields.gadfEndDate).ValueChanged = True) Then
          DeleteUnclaimedAdjustments(False)
        End If

        If vEndDateChanged Then ProcessChangeOfEndDate()
        If vStartDateChanged Then ProcessChangeOfStartDate()
      ElseIf BatchNumber = 0 Then
        ProcessNewDeclarationPayments(False)
      End If

      If vCreateLines Then
        PopulateUnclaimedLines()
      End If

      'Currently only the journaling of the creation of declaration and the cancellation of a declaration is required
      If Not mvExisting Then
        vJournalOperation = JournalOperations.jnlInsert
      ElseIf pSaveOptions.Cancel Then
        vJournalOperation = JournalOperations.jnlCancel
      End If
      If vJournalOperation > 0 Then mvEnv.AddJournalRecord(JournalTypes.jnlGiftAidDeclaration, vJournalOperation, ContactNumber, mvContact.AddressNumber, DeclarationNumber, 0, 0, pSaveOptions.BatchNumber, pSaveOptions.TransactionNumber)

      Me.ClassFields.Save(mvEnv, mvExisting, pSaveOptions.AmendedBy, pSaveOptions.Audit)
      If vTransaction Then mvEnv.Connection.CommitTransaction()
      If Not Me.IsMandatoryDataComplete Then
        ExecutionContext.GetInstance.InformationMessage = ProjectText.GiftAidDataMissing
      End If
    End Sub

    Private Sub CancelPreviousDeclaration(pCancelInfo As CancellationInfo)
      Dim vPreviousGAD As GiftAidDeclaration = Me.GetPreviousDeclaration()
      If vPreviousGAD IsNot Nothing Then
        Dim vErrorText As String = String.Empty
        Dim vCancelDate As Date = DateAdd(DateInterval.Day, -1, Date.Parse(Me.StartDate))
        Dim vPreviousStartDate As Date = Date.Parse(vPreviousGAD.StartDate)
        vCancelDate = New Date(Math.Max(vCancelDate.Ticks, vPreviousStartDate.Ticks))
        If Not vPreviousGAD.CanCancel(vErrorText, vCancelDate.ToString()) Then
          RaiseError(DataAccessErrors.daePreviousGADCancelFailed, vErrorText)
        End If
        vPreviousGAD.Cancel(pCancelInfo.CancellationReason, pCancelInfo.CancellationSource, pNewEndDate:=vCancelDate.ToString())
      End If
    End Sub

    Friend Function GetPreviousDeclaration() As GiftAidDeclaration
      Dim vWhere As CDBFields = GetPreviousDeclarationWhere(True)
      Dim vOrderByItems As New Dictionary(Of Integer, OrderByDirection) From
        {
          {GiftAidDeclarationFields.gadfStartDate, OrderByDirection.Descending}
        }
      Return CARERecordFactory.SelectInstance(Of GiftAidDeclaration)(Me.Environment, vWhere, Me.ClassFields.OrderByClause(vOrderByItems))
    End Function

    ''' <summary>
    ''' Generates a where clause for getting the previous declaration.  In future, please just use this function to get the previous GAD.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>The intention is that if there is one central place that decided what a previous GAD is, then we can consitently use it rather than copy the same code everywhere</remarks>
    Private Function GetPreviousDeclarationWhere(pIncludeCancelledGAD As Boolean) As CDBFields
      Dim vWhere As New CDBFields()

      'Match the existing GAD's scope.  If my values are null (except Contact number of course) then the previous GAD must also be null
      'contact number = my contact number
      'batch number = my batch number
      'trans number = my trans number
      'PP number = my PP Number
      If ClassFields(GiftAidDeclarationFields.gadfContactNumber).Value.Length = 0 Then
        RaiseError(DataAccessErrors.daeContactNumberInvalid)
      End If
      Dim vStandardFields As New List(Of ClassField)
      vStandardFields.AddRange({ClassFields(GiftAidDeclarationFields.gadfContactNumber),
                               ClassFields(GiftAidDeclarationFields.gadfBatchNumber),
                               ClassFields(GiftAidDeclarationFields.gadfPaymentPlanNumber)})
      vStandardFields.ForEach(Sub(vField As ClassField) vWhere.Add(vField.Name, vField.FieldType, vField.Value))

      Dim vMyStartDateColumn As String = String.Format("COALESCE({0},'{1}')",
                                                       Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Name,
                                                       If(Not String.IsNullOrWhiteSpace(ClassFields(GiftAidDeclarationFields.gadfStartDate).Value),
                                                                                        ClassFields(GiftAidDeclarationFields.gadfStartDate).Value,
                                                                                        DateTime.MaxValue.ToString(CAREDateFormat))
                                                       )
      'clause 1: where previous GAD between my start date
      vWhere.Add(
                  vMyStartDateColumn,
                  ClassFields(GiftAidDeclarationFields.gadfStartDate).FieldType,
                  ClassFields(GiftAidDeclarationFields.gadfStartDate).Value,
                  CDBField.FieldWhereOperators.fwoBetweenFrom Or CDBField.FieldWhereOperators.fwoOpenBracket
                 )
      'clause 1: ...and my end date
      vWhere.Add(
                  String.Format("{0}_2", vMyStartDateColumn),
                  ClassFields(GiftAidDeclarationFields.gadfEndDate).FieldType,
                  If(Not String.IsNullOrWhiteSpace(ClassFields(GiftAidDeclarationFields.gadfEndDate).Value),
                                                   ClassFields(GiftAidDeclarationFields.gadfEndDate).Value,
                                                   DateTime.MaxValue.ToString(CAREDateFormat)),
                  CDBField.FieldWhereOperators.fwoBetweenTo
                 )

      Dim vTheirStartDateColumn As String = String.Format("COALESCE({0},'{1}')",
                                                          ClassFields(GiftAidDeclarationFields.gadfStartDate).Name,
                                                          If(Not String.IsNullOrWhiteSpace(ClassFields(GiftAidDeclarationFields.gadfStartDate).Value),
                                                                                           ClassFields(GiftAidDeclarationFields.gadfStartDate).Value,
                                                                                           DateTime.MaxValue.ToString(CAREDateFormat))
                                                         )
      Dim vTheirEndDateColumn As String = String.Format("COALESCE({0},'{1}')",
                                                        ClassFields(GiftAidDeclarationFields.gadfEndDate).Name,
                                                          If(Not String.IsNullOrWhiteSpace(ClassFields(GiftAidDeclarationFields.gadfEndDate).Value),
                                                                                           ClassFields(GiftAidDeclarationFields.gadfEndDate).Value,
                                                                                           DateTime.MaxValue.ToString(CAREDateFormat))
                                                         )
      Dim vMyStartDateValue As String = String.Format("'{0}'",
                                                      ClassFields(GiftAidDeclarationFields.gadfStartDate).Value)
      'clause 2: or my start date between previous GAD start date
      vWhere.Add(
                  vMyStartDateValue,
                  CDBField.FieldTypes.cftUnknown,
                  vTheirStartDateColumn,
                  CDBField.FieldWhereOperators.fwoBetweenFrom Or CDBField.FieldWhereOperators.fwoOR
                 )
      'clause 1: ...and previous GAD end date
      vWhere.Add(
                  String.Format("{0}_2", vMyStartDateValue),
                  CDBField.FieldTypes.cftUnknown,
                  vTheirEndDateColumn,
                  CDBField.FieldWhereOperators.fwoBetweenTo Or CDBField.FieldWhereOperators.fwoCloseBracket
                 )
      If Not pIncludeCancelledGAD Then
        vWhere.Add(ClassFields(GiftAidDeclarationFields.gadfCancellationReason).Name, ClassFields(GiftAidDeclarationFields.gadfCancellationReason).Value)
      End If
      Return vWhere
    End Function

    Friend Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      'This is only used by Contact Merge to just save the Declaration
      'IT WILL NOT AND MUST NEVER re-create the unclaimed lines
      Dim vJournalOperation As JournalOperations
      Dim vTrans As Boolean

      SetValid(GiftAidDeclarationFields.gadfAll)

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      'Currently only the journalling of the creation of declaration and the cancellation of a declaration is required
      If Not mvExisting Then
        vJournalOperation = JournalOperations.jnlInsert
      End If
      If vJournalOperation > 0 Then mvEnv.AddJournalRecord(JournalTypes.jnlGiftAidDeclaration, vJournalOperation, ContactNumber, mvContact.AddressNumber, DeclarationNumber)

      Me.ClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)

      If vTrans Then mvEnv.Connection.CommitTransaction()

    End Sub

    Public Sub Delete()
      Dim vWhereFields As New CDBFields

      mvEnv.Connection.StartTransaction()

      vWhereFields.Add("cd_number", CDBField.FieldTypes.cftLong, DeclarationNumber)
      vWhereFields.Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")

      'Delete unclaimed and potential claim lines
      mvEnv.Connection.DeleteRecords("declaration_lines_unclaimed", vWhereFields, False)
      mvEnv.Connection.DeleteRecords("declaration_potential_lines", vWhereFields, False)

      'Delete any unclaimed claim-adjustment transactions
      DeleteUnclaimedAdjustments(True)

      'Delete unfulfilled contact mailing documents
      With vWhereFields
        .Clear()
        .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        .Add("declaration_number", CDBField.FieldTypes.cftLong, DeclarationNumber)
        .Add("fulfillment_number", CDBField.FieldTypes.cftLong, "")
      End With
      mvEnv.Connection.DeleteRecords("contact_mailing_documents", vWhereFields, False)

      'Now delete the Declaration itself
      Me.ClassFields.Delete(mvEnv.Connection)

      mvEnv.Connection.CommitTransaction()

    End Sub

    Public ReadOnly Property IsValid As Boolean
      Get
        IsValid = True

        If IsDate(StartDate) Then
          If CDate(StartDate) > CDate(TodaysDate()) Then IsValid = False
        End If
        If IsDate(EndDate) Then
          If CDate(EndDate) < CDate(TodaysDate()) Then IsValid = False
        End If
      End Get
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy As String
      Get
        AmendedBy = Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn As String
      Get
        AmendedOn = Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CAFPaymentMethod As String
      Get
        CAFPaymentMethod = mvCAFPayMethod
      End Get
    End Property

    Public ReadOnly Property ConfirmedOn As String
      Get
        ConfirmedOn = Me.ClassFields.Item(GiftAidDeclarationFields.gadfConfirmedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber As Integer
      Get
        ContactNumber = Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).IntegerValue
      End Get
    End Property

    Public Property DeclarationDate As String
      Get
        Return Me.ClassFields(GiftAidDeclarationFields.gadfDeclarationDate).Value
      End Get
      Private Set(value As String)
        Me.ClassFields(GiftAidDeclarationFields.gadfDeclarationDate).Value = value
      End Set
    End Property

    Public ReadOnly Property DeclarationNumber As Integer
      Get
        DeclarationNumber = Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).IntegerValue
      End Get
    End Property

    Public Property DeclarationTypeCode As String
      Get
        Return DeclarationTypeCodeFromType(GiftAidDeclarationTypes.gadtAll)
      End Get
      Private Set(value As String)
        Me.ClassFields(GiftAidDeclarationFields.gadfDeclarationType).Value = "A"
      End Set
    End Property

    Public Property EndDate As String
      Get
        Return Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
      End Get
      Friend Set(value As String)
        Dim currentStartDate As Date = Nothing
        If Date.TryParse(Me.StartDate, currentStartDate) Then
          Dim newEndDate As Date = Nothing
          If Date.TryParse(value, newEndDate) Then
            newEndDate = If(newEndDate < currentStartDate, currentStartDate, newEndDate)
            value = newEndDate.ToString(Utilities.Common.CAREDateFormat())
          End If
        End If
        If Me.CancellationReason.IsNullOrWhitespace = False AndAlso
          Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value <> value Then
          Dim vError As String = Resources.ErrorText.DaeRecordIsCancelled.Replace("%1", Me.ClassFields.Caption)
          RaiseError(DataAccessErrors.daeInvalidOperation, vError)
        End If
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value = value
      End Set
    End Property

    Public ReadOnly Property GiftAidMinimum As Double
      Get
        GiftAidMinimum = mvGiftAidMinimum
      End Get
    End Property

    Public ReadOnly Property MethodCode As String
      Get
        MethodCode = Me.ClassFields.Item(GiftAidDeclarationFields.gadfMethod).Value
      End Get
    End Property

    Public ReadOnly Property Notes As String
      Get
        Notes = Me.ClassFields.Item(GiftAidDeclarationFields.gadfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property Source As String
      Get
        Source = Me.ClassFields.Item(GiftAidDeclarationFields.gadfSource).Value
      End Get
    End Property

    Public ReadOnly Property StartDate As String
      Get
        StartDate = Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value
      End Get
    End Property

    Public ReadOnly Property CanAddDeclaration As Boolean
      Get
        Dim vSQL As String

        vSQL = "contact_number = " & Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value
        If Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then
          vSQL = vSQL & " AND ((start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) & ") OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) & " OR end_date IS NULL)) OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) & " AND end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) & ") OR (start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value))
          vSQL = vSQL & " AND start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value)) & " OR end_date IS NULL)))"
        Else
          vSQL = vSQL & " AND (end_date IS NULL OR end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value)) & ")"
        End If
        CanAddDeclaration = mvEnv.Connection.GetCount("gift_aid_declarations", Nothing, vSQL) = 0
      End Get
    End Property

    Public ReadOnly Property DeclarationType As GiftAidDeclarationTypes
      Get
        Return GiftAidDeclarationTypes.gadtAll
      End Get
    End Property

    Public ReadOnly Property DeclarationTypeDesc As String
      Get
        Return "All"
      End Get
    End Property

    Public ReadOnly Property Method As GiftAidDeclarationMethods
      Get
        Dim vMethod As GiftAidDeclarationMethods

        Select Case Me.ClassFields.Item(GiftAidDeclarationFields.gadfMethod).Value
          Case "O"
            vMethod = GiftAidDeclarationMethods.gadmOral
          Case "W"
            vMethod = GiftAidDeclarationMethods.gadmWritten
          Case "E"  'BR19026
            vMethod = GiftAidDeclarationMethods.gadmElectronic
        End Select
        Method = vMethod
      End Get
    End Property

    Public ReadOnly Property BatchNumber As Integer
      Get
        BatchNumber = Me.ClassFields.Item(GiftAidDeclarationFields.gadfBatchNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TransactionNumber As Integer
      Get
        TransactionNumber = Me.ClassFields.Item(GiftAidDeclarationFields.gadfTransactionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property PaymentPlanNumber As Integer
      Get
        PaymentPlanNumber = Me.ClassFields.Item(GiftAidDeclarationFields.gadfPaymentPlanNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CancellationReason As String
      Get
        CancellationReason = Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationReason).Value
      End Get
    End Property

    Public ReadOnly Property CancellationSource As String
      Get
        CancellationSource = Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationSource).Value
      End Get
    End Property

    Public ReadOnly Property CancelledBy As String
      Get
        CancelledBy = Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn As String
      Get
        CancelledOn = Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancelledOn).Value
      End Get
    End Property

    Public ReadOnly Property GiftAidStartDate As String
      Get
        GiftAidStartDate = mvGiftAidStartDate
      End Get
    End Property

    Public ReadOnly Property HasTaxClaims(Optional ByVal pCheckClaimAdjustments As Boolean = False) As Boolean
      Get
        'Check whether Declaration has Tax Claims
        Dim vRS As CDBRecordSet
        Dim vFields As CDBFields
        Dim vSQL As String

        If mvTaxClaimsChecked = False Then
          vFields = New CDBFields
          With vFields
            .Add("cd_number", CDBField.FieldTypes.cftLong, Me.ClassFields(GiftAidDeclarationFields.gadfDeclarationNumber).IntegerValue)
            .Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
            .Add("fh.batch_number", CDBField.FieldTypes.cftLong, "dtcl.batch_number")
            .Add("fh.transaction_number", CDBField.FieldTypes.cftLong, "dtcl.transaction_number")
          End With
          vSQL = "SELECT MIN(transaction_date) AS min_trn_date" & ", MAX(transaction_date) AS max_trn_date" & " FROM declaration_tax_claim_lines dtcl, financial_history fh WHERE " & mvEnv.Connection.WhereClause(vFields) & " GROUP BY cd_number"
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          If vRS.Fetch() = True Then
            mvFirstClaimPayDate = vRS.Fields(1).Value
            mvLastClaimPayDate = vRS.Fields(2).Value
          End If
          vRS.CloseRecordSet()
          mvTaxClaimsChecked = True

          If pCheckClaimAdjustments Then
            'Also check any Claim Adjustment batches
            With vFields
              .Clear()
              .Add("batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment))
              .Add("bt.batch_number", CDBField.FieldTypes.cftLong, "b.batch_number")
              .Add("dtcl.batch_number", CDBField.FieldTypes.cftLong, "bt.batch_number")
              .Add("dtcl.transaction_number", CDBField.FieldTypes.cftLong, "bt.transaction_number")
              .Add("cd_number", CDBField.FieldTypes.cftLong, DeclarationNumber)
              .Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
            End With
            vSQL = "SELECT MIN(transaction_date) AS min_trn_date" & ", MAX(transaction_date) AS max_trn_date" & " FROM batches b, batch_transactions bt, declaration_tax_claim_lines dtcl WHERE " & mvEnv.Connection.WhereClause(vFields) & " GROUP BY cd_number"
            vRS = mvEnv.Connection.GetRecordSet(vSQL)
            If vRS.Fetch() = True Then
              If IsDate(mvFirstClaimPayDate) Then
                If CDate(mvFirstClaimPayDate) > CDate(vRS.Fields(1).Value) Then mvFirstClaimPayDate = vRS.Fields(1).Value
              Else
                mvFirstClaimPayDate = vRS.Fields(1).Value
              End If
              If IsDate(mvLastClaimPayDate) Then
                If CDate(vRS.Fields(2).Value) > CDate(mvLastClaimPayDate) Then mvLastClaimPayDate = vRS.Fields(2).Value
              Else
                mvLastClaimPayDate = vRS.Fields(2).Value
              End If
            End If
            vRS.CloseRecordSet()
          End If
        End If
        HasTaxClaims = (Len(mvLastClaimPayDate) > 0)
      End Get
    End Property

    Public ReadOnly Property EarliestClaimedPaymentDate As String
      Get
        EarliestClaimedPaymentDate = mvFirstClaimPayDate
      End Get
    End Property

    Public ReadOnly Property LatestClaimedPaymentDate As String
      Get
        LatestClaimedPaymentDate = mvLastClaimPayDate
      End Get
    End Property

    Public ReadOnly Property GiftAidEarliestStartDate As String
      Get
        'Return the earliest Declaration start date that can be set
        Dim vCalcDate As String = ""
        Dim vAccountingDate As String
        Dim vBackClaimYears As Integer
        Dim vTaxClaimLimitChangeDate As String
        Dim vNewRule As Boolean

        Select Case mvCharityTaxStatus
          Case GACharityTaxStatus.gaCTSTrust
            'Trust for Tax Purposes
            'First set the tax period end date
            vTaxClaimLimitChangeDate = mvEnv.GetConfig("tax_claim_limit_change_date", "01/04/2010")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBackClaimYears) Then
              vBackClaimYears = GetBackClaimYears()
            Else
              vBackClaimYears = 5
            End If
            vAccountingDate = CStr(DateSerial(Year(CDate(TodaysDate())), 4, 5)) '5th April

            If CDate(vTaxClaimLimitChangeDate) > CDate(TodaysDate()) Then
              'End next 31 January
              vCalcDate = CStr(DateSerial(Year(CDate(TodaysDate())), 1, 31))
            Else
              'End of this tax year
              vCalcDate = CStr(DateSerial(Year(CDate(TodaysDate())), 4, 5))
              vNewRule = True
            End If

            If CDate(TodaysDate()) > CDate(vCalcDate) Then vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, CDate(vCalcDate)))
            'Deduct years
            vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -vBackClaimYears, CDate(vCalcDate)))

            'End date of prior tax year
            vCalcDate = CStr(DateSerial(Year(CDate(vCalcDate)) - 1, Month(CDate(vAccountingDate)), Day(CDate(vAccountingDate))))
            If vNewRule Then
              'go to start of next tax year - NEW
              vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vCalcDate)))
            Else
              'Start date of prior tax year - OLD
              vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vCalcDate))))
            End If

          'Adding GAD up to 31/01/2007 will set earliest start date to 06/04/2000
          'Adding GAD between 01/02/2007 and 31/01/2008 will set earliest start date to 06/04/2001
          '.....

          Case GACharityTaxStatus.gaCTSCompany
            'Company for Tax Purposes
            'First set the current accounting period start date
            vAccountingDate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAAccountingPeriodStart)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBackClaimYears) Then
              vBackClaimYears = GetBackClaimYears()
            Else
              vBackClaimYears = 6
            End If
            If Len(vAccountingDate) = 0 Then vAccountingDate = "06/04/2006" 'Database not upgraded!
            vAccountingDate = CStr(DateSerial(Year(CDate(TodaysDate())), Month(CDate(vAccountingDate)), Day(CDate(vAccountingDate))))
            If CDate(vAccountingDate) > CDate(TodaysDate()) Then vAccountingDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, CDate(vAccountingDate)))

            'Start date of current accounting period
            vCalcDate = vAccountingDate
            'Deduct years
            vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -vBackClaimYears, CDate(vCalcDate)))

            '1st January to 31st December accounting period
            'Adding GAD up to 31/12/2006 will set earliest start date to 06/04/2000
            'Adding GAD between 01/01/2007 and 31/12/2007 will set earliest start date to 01/01/2001
            '.....

            '6th April to 5th April accounting period
            'Adding GAD up to 05/04/2007 will set earliest start date to 06/04/2000
            'Adding GAD between 06/04/2007 and 05/04/2008 will set earliest start date to 06/04/2001
            '.....

        End Select

        'Set minimum date
        If CDate(mvGiftAidStartDate) > CDate(vCalcDate) Then vCalcDate = mvGiftAidStartDate

        GiftAidEarliestStartDate = CDate(vCalcDate).ToString(CAREDateFormat)
      End Get
    End Property

    Public ReadOnly Property EarliestClaimableTransactionDate As String
      Get
        'Used by the Tax Claim processing to set the earlest transaction date that can be claimed
        Dim vCalcDate As String = ""
        Dim vAccountingDate As String
        Dim vBackClaimYears As Integer
        Dim vTaxClaimLimitChangeDate As String
        Dim vNewRule As Boolean

        Select Case mvCharityTaxStatus
          Case GACharityTaxStatus.gaCTSTrust
            'Trust for Tax Purposes
            'First set the tax period end date
            vTaxClaimLimitChangeDate = mvEnv.GetConfig("tax_claim_limit_change_date", "01/04/2010")
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBackClaimYears) Then
              vBackClaimYears = GetBackClaimYears()
            Else
              vBackClaimYears = 5
            End If
            vAccountingDate = CStr(DateSerial(Year(CDate(TodaysDate())), 4, 5)) '5 April current year
            If CDate(vTaxClaimLimitChangeDate) > CDate(TodaysDate()) Then
              'End next 31 January
              vCalcDate = CStr(DateSerial(Year(CDate(TodaysDate())), 1, 31))
            Else
              'End of this tax year
              vCalcDate = CStr(DateSerial(Year(CDate(TodaysDate())), 4, 5))
              vNewRule = True
            End If
            If CDate(TodaysDate()) > CDate(vCalcDate) Then vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, CDate(vCalcDate)))
            'Less years
            vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -vBackClaimYears, CDate(vCalcDate)))
            'End prev tax year
            vCalcDate = CStr(DateSerial(Year(CDate(vCalcDate)) - 1, Month(CDate(vAccountingDate)), Day(CDate(vAccountingDate))))
            If vNewRule Then
              'go to start of next tax year - NEW
              vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vCalcDate)))
            Else
              'Start date of prior tax year - OLD
              vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vCalcDate))))
            End If

          'Payments dated between 06/04/2000 to 05/04/2001 to be claimed by 31/01/2007
          'Payments dated between 06/04/2001 to 05/04/2002 to be claimed by 31/01/2008
          '.....

          Case GACharityTaxStatus.gaCTSCompany
            'Company for Tax Purposes
            'Start Accounting Period
            vAccountingDate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAAccountingPeriodStart)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataBackClaimYears) Then
              vBackClaimYears = GetBackClaimYears()
            Else
              vBackClaimYears = 6
            End If
            If Len(vAccountingDate) = 0 Then vAccountingDate = "06/04/2006" 'Database not upgraded!
            vAccountingDate = CStr(DateSerial(Year(CDate(TodaysDate())), Month(CDate(vAccountingDate)), Day(CDate(vAccountingDate))))
            If CDate(vAccountingDate) > CDate(TodaysDate()) Then vAccountingDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, CDate(vAccountingDate)))
            'Less years
            vCalcDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -vBackClaimYears, CDate(vAccountingDate)))

            '1st January to 31st December accounting period
            'Payment dated between 06/04/2000 to 31/12/2000 to be claimed by 31/12/2006
            'Payments dated between 01/01/2001 to 31/12/2001 to be claimed by 31/12/2007
            '.....

            '6th April to 5th April accounting period
            'Payments dated between 06/04/2000 to 05/04/2001 to be claimed by 05/04/2007
            'Payments dated between 06/04/2001 to 05/04/2002 to be claimed by 05/04/2008
            '.....

        End Select

        'Set minimum date
        If CDate(vCalcDate) < CDate(mvGiftAidStartDate) Then vCalcDate = mvGiftAidStartDate
        EarliestClaimableTransactionDate = CDate(vCalcDate).ToString(CAREDateFormat)
      End Get
    End Property

    Public Sub UpdateFields(ByVal pContact As Integer, ByVal pDecDate As String, ByVal pSource As String, ByVal pConfirmedOn As String, ByVal pMethod As String, ByVal pStartDate As String, ByVal pEndDate As String, ByVal pNotes As String, Optional ByVal pPayPlanNumber As Integer = 0, Optional ByVal pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0)

      Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value = CStr(pContact)
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfSource).Value = pSource
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfMethod).Value = pMethod
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value = pStartDate
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfNotes).Value = pNotes
      If pDecDate.Length > 0 Then Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationDate).Value = pDecDate
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfConfirmedOn).Value = pConfirmedOn
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value = pEndDate

      If pPayPlanNumber > 0 Then
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfPaymentPlanNumber).Value = CStr(pPayPlanNumber)
      Else
        If pBatchNumber > 0 Then
          Me.ClassFields.Item(GiftAidDeclarationFields.gadfBatchNumber).Value = CStr(pBatchNumber)
          Me.ClassFields.Item(GiftAidDeclarationFields.gadfTransactionNumber).Value = CStr(pTransactionNumber)
        End If
      End If

      If mvExisting = False And Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).IntegerValue > 0 Then
        'Creating a new GAD in Rich Client comes in here without the Contact having been initialised.
        If Me.Contact Is Nothing Then
          mvContact = New Contact(mvEnv)
          mvContact.Init(pContact)
        End If
      End If
    End Sub

    Private Sub GenerateUnclaimedLines()
      Dim vContactLink As ContactLink
      Dim vDLU As New DeclarationLinesUnclaimed(Me.Environment)
      Dim vDLUFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vRow As CDBDataRow
      Dim vTable As CDBDataTable
      Dim vContactNos As String
      Dim vPrevCovEligGA As String
      Dim vPrevCovExpiry As String
      Dim vPrevCovNo As Integer
      Dim vPrevCovOphPayNo As Integer
      Dim vPrevCovPayNo As Integer
      Dim vPrevCovPPNo As Integer
      Dim vPrevCovType As String
      Dim vPrevGiftMem As Boolean
      Dim vPrevPPNo As Integer
      Dim vSQL As String
      Dim vSQLBTWhere As String
      Dim vSQLDon As String
      Dim vSQLInnerJoin As String
      Dim vSQLInnerJoinTT As String
      Dim vSQLJoinCov As String
      Dim vSQLOuterJoin As String

      vContactNos = CStr(ContactNumber)
      For Each vContactLink In mvContact.GetJointLinks(False)
        vContactNos = vContactNos & "," & vContactLink.ContactNumber1
      Next vContactLink

      'Create declaration_lines_unclaimed records
      With vDLUFields
        .Add("cd_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value)
        .Add("contact_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value)
        .Add("batch_number", CDBField.FieldTypes.cftLong, 0)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger, 0)
        .Add("line_number", CDBField.FieldTypes.cftInteger, 0)
        .Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
        .Add("net_amount", CDBField.FieldTypes.cftNumeric)
      End With

      'Build WhereFields collection for both SQL statements
      With vWhereFields
        .Add("fh.contact_number", CDBField.FieldTypes.cftLong, vContactNos, CType(IIf(InStr(vContactNos, ",") > 0, CDBField.FieldWhereOperators.fwoIn, CDBField.FieldWhereOperators.fwoEqual), CDBField.FieldWhereOperators))
        If BatchNumber > 0 Then
          .Add("fh.batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
          .Add("fh.transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
        End If
        If Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
          .Add("fh.transaction_date", CDBField.FieldTypes.cftDate, StartDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
          .Add("fh.transaction_date#2", CDBField.FieldTypes.cftDate, EndDate, CDBField.FieldWhereOperators.fwoBetweenTo)
        Else
          .Add("fh.transaction_date", CDBField.FieldTypes.cftDate, StartDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
        If IsDate(CancelledOn) Then .Add("fh.transaction_date#3", CDBField.FieldTypes.cftDate, CancelledOn, CDBField.FieldWhereOperators.fwoLessThan)
        .Add("fh.payment_method", CDBField.FieldTypes.cftCharacter, mvCAFPayMethod, CDBField.FieldWhereOperators.fwoNotEqual)
        .Add("fh.amount", CDBField.FieldTypes.cftNumeric, mvGiftAidMinimum.ToString, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        .Add("tt.transaction_sign", CDBField.FieldTypes.cftCharacter, "C")
        .Add("fhd.amount", CDBField.FieldTypes.cftNumeric, mvGiftAidMinimum.ToString, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        .Add("p.eligible_for_gift_aid", CDBField.FieldTypes.cftCharacter, "Y")
        .Add("lbr.batch_number", CDBField.FieldTypes.cftLong, "")
        .Add("gph.batch_number", CDBField.FieldTypes.cftLong, "")
      End With

      'Build Inner & Outer Joins used by both SQL statements
      vSQLInnerJoinTT = " INNER JOIN transaction_types tt ON fh.transaction_type = tt.transaction_type"
      vSQLInnerJoin = " INNER JOIN financial_history_details fhd ON %1.batch_number = fhd.batch_number AND %2.transaction_number = fhd.transaction_number %3"
      vSQLInnerJoin = vSQLInnerJoin & " INNER JOIN products p ON fhd.product = p.product"

      vSQLOuterJoin = " LEFT OUTER JOIN batch_transactions bt ON fh.batch_number = bt.batch_number AND fh.transaction_number = bt.transaction_number"
      vSQLOuterJoin = vSQLOuterJoin & " LEFT OUTER JOIN legacy_bequest_receipts lbr ON fhd.batch_number = lbr.batch_number AND fhd.transaction_number = lbr.transaction_number AND fhd.line_number = lbr.line_number"
      vSQLOuterJoin = vSQLOuterJoin & " LEFT OUTER JOIN gaye_pledge_payment_history gph ON fh.batch_number = gph.batch_number AND fh.transaction_number = gph.transaction_number"

      vSQLJoinCov = " LEFT OUTER JOIN (SELECT oph.batch_number, oph.transaction_number, oph.line_number, covenant_number, o.order_number AS cov_order_number, order_type AS cov_order_type,"
      vSQLJoinCov = vSQLJoinCov & " oph.payment_number AS oph_payment_number, c.payment_number AS cov_payment_number,"
      vSQLJoinCov = vSQLJoinCov & " o.eligible_for_gift_aid AS cov_eligible_ga, " & mvEnv.Connection.DBAddYears("c.start_date", "c.covenant_term") & " AS cov_expiry"
      vSQLJoinCov = vSQLJoinCov & " FROM order_payment_history oph INNER JOIN financial_history fh ON oph.batch_number = fh.batch_number AND oph.transaction_number = fh.transaction_number"
      vSQLJoinCov = vSQLJoinCov & " INNER JOIN covenants c ON oph.order_number = c.order_number INNER JOIN orders o ON c.order_number = o.order_number"
      vSQLJoinCov = vSQLJoinCov & " WHERE c.start_date <= fh.transaction_date"
      vSQLJoinCov = vSQLJoinCov & " AND ((c.cancelled_on IS NULL) OR (c.cancelled_on IS NOT NULL AND c.cancelled_on >= fh.transaction_date))"
      vSQLJoinCov = vSQLJoinCov & " ) cv ON fhd.batch_number = cv.batch_number AND fhd.transaction_number = cv.transaction_number AND fhd.line_number = cv.line_number"

      vSQLBTWhere = " AND (bt.eligible_for_gift_aid IS NULL OR (bt.eligible_for_gift_aid IS NOT NULL AND bt.eligible_for_gift_aid = 'Y'))"

      'Donation payments
      If (Me.DeclarationType And GiftAidDeclarationTypes.gadtDonation) = GiftAidDeclarationTypes.gadtDonation Then 'Declaration Type is Donation or All
        vDLUFields(DLUFields.dfBatchNumber).Value = CStr(0) 'Force it to be different
        vPrevGiftMem = False
        vPrevPPNo = 0
        vPrevCovNo = 0
        vPrevCovPPNo = 0
        vPrevCovType = ""
        vPrevCovOphPayNo = 0
        vPrevCovPayNo = 0
        vPrevCovEligGA = ""
        vPrevCovExpiry = ""

        vSQLDon = "SELECT /* SQLServerCSC */ fhd.batch_number,fhd.transaction_number,fhd.line_number,fhd.amount, fh.transaction_date, oph.order_number, covenant_number, cov_order_number, cov_order_type, oph_payment_number, cov_payment_number, cov_eligible_ga, cov_expiry"
        vSQLDon = vSQLDon & " FROM financial_history fh"
        vSQLDon = vSQLDon & vSQLInnerJoinTT
        vSQLDon = vSQLDon & Replace(Replace(Replace(vSQLInnerJoin, "%1", "fh"), "%2", "fh"), "%3", "")
        vSQLDon = vSQLDon & vSQLOuterJoin
        vSQLDon = vSQLDon & " LEFT OUTER JOIN (SELECT batch_number, transaction_number, line_number, oph.order_number, eligible_for_gift_aid FROM order_payment_history oph INNER JOIN orders o ON oph.order_number = o.order_number) oph ON fhd.batch_number = oph.batch_number AND fhd.transaction_number = oph.transaction_number AND fhd.line_number = oph.line_number"
        vSQLDon = vSQLDon & vSQLJoinCov
        vSQLDon = vSQLDon & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & vSQLBTWhere
        vSQLDon = vSQLDon & " AND (oph.eligible_for_gift_aid IS NULL OR (oph.eligible_for_gift_aid IS NOT NULL AND oph.eligible_for_gift_aid = 'Y'))"
        vSQLDon = vSQLDon & " ORDER BY fhd.batch_number,fhd.transaction_number,fhd.line_number"

        vTable = New CDBDataTable
        vTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQLDon))

        For Each vRow In vTable.Rows
          ProcessUnclaimed(vDLUFields, vRow.IntegerItem("batch_number"), vRow.IntegerItem("transaction_number"), vRow.IntegerItem("line_number"), Val(vRow.Item("amount")), vRow.Item("transaction_date"), vPrevCovNo, vPrevCovPPNo, vPrevCovType, vPrevCovOphPayNo, vPrevCovPayNo, vPrevCovEligGA, vPrevCovExpiry, vPrevPPNo)
          'ProcessUnclaimed will process the previous transaction data so need to store some values for checking next time
          vPrevPPNo = vRow.IntegerItem("order_number")
          vPrevCovNo = vRow.IntegerItem("covenant_number")
          vPrevCovPPNo = vRow.IntegerItem("cov_order_number")
          vPrevCovType = vRow.Item("cov_order_type")
          vPrevCovOphPayNo = vRow.IntegerItem("oph_payment_number")
          vPrevCovPayNo = vRow.IntegerItem("cov_payment_number")
          vPrevCovEligGA = vRow.Item("cov_eligible_ga")
          vPrevCovExpiry = vRow.Item("cov_expiry")
        Next vRow
        ProcessUnclaimed(vDLUFields, 0, 0, 0, 0, "", vPrevCovNo, vPrevCovPPNo, vPrevCovType, vPrevCovOphPayNo, vPrevCovPayNo, vPrevCovEligGA, vPrevCovExpiry, vPrevPPNo) 'Write any remaining dlu line details as this will group batch/trans numbers
      End If


      'Related adjustments - only create line if reversed transaction is for the contact we are looking at
      'and the line is not already in DLU for another Declaration (could happen when transaction paid by joint Contact)
      vSQL = "INSERT INTO declaration_lines_unclaimed(cd_number,declaration_or_covenant_number,contact_number,batch_number,transaction_number,line_number,net_amount)"
      vSQL = vSQL & " SELECT cd_number,'D'," & Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value & ",r.batch_number,r.transaction_number,r.line_number,(dlu.net_amount*-1)"
      vSQL = vSQL & " FROM declaration_lines_unclaimed dlu, reversals r, batch_transactions bt"
      vSQL = vSQL & " WHERE dlu.cd_number = " & Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).IntegerValue & " AND dlu.declaration_or_covenant_number = 'D'"
      vSQL = vSQL & " AND r.was_batch_number = dlu.batch_number AND r.was_transaction_number = dlu.transaction_number AND r.was_line_number = dlu.line_number"
      vSQL = vSQL & " AND bt.batch_number = r.batch_number AND bt.transaction_number = r.transaction_number AND bt.contact_number"
      If InStr(vContactNos, ",") > 0 Then
        vSQL = vSQL & " IN (" & vContactNos & ")"
      Else
        vSQL = vSQL & " = " & vContactNos
      End If
      vSQL = vSQL & " AND r.batch_number NOT IN (SELECT dlu1.batch_number FROM declaration_lines_unclaimed dlu1 Where dlu1.batch_number = r.batch_number"
      vSQL = vSQL & " AND dlu1.transaction_number = r.transaction_number AND dlu1.line_number = r.line_number AND dlu1.declaration_or_covenant_number = 'D')"
      mvEnv.Connection.ExecuteSQL(vSQL)

      'BR 9563 related CLAIMED adjustments - only create line if reversed transaction is for the contact we are looking at
      vSQL = " fh.batch_number = r.batch_number AND fh.transaction_number = r.transaction_number AND fh.contact_number"
      If InStr(vContactNos, ",") > 0 Then
        vSQL = vSQL & " IN (" & vContactNos & ")"
      Else
        vSQL = vSQL & " = " & vContactNos
      End If
      vSQL = vSQL & " AND r.batch_number NOT IN (SELECT batch_number FROM declaration_lines_unclaimed dlu WHERE dlu.batch_number = r.batch_number AND dlu.transaction_number = r.transaction_number AND dlu.line_number = r.line_number)"
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT r.batch_number,r.transaction_number,r.line_number,dtcl.net_amount FROM declaration_tax_claim_lines dtcl, reversals r, financial_history fh WHERE dtcl.cd_number = " & Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).IntegerValue & " AND dtcl.declaration_or_covenant_number = 'D' AND r.was_batch_number = dtcl.batch_number AND r.was_transaction_number = dtcl.transaction_number AND r.was_line_number = dtcl.line_number AND " & vSQL)
      While vRecordSet.Fetch() = True
        vDLUFields(DLUFields.dfBatchNumber).Value = CStr(vRecordSet.Fields("batch_number").IntegerValue)
        vDLUFields(DLUFields.dfTransactionNumber).Value = vRecordSet.Fields("transaction_number").Value
        vDLUFields(DLUFields.dfLineNumber).Value = vRecordSet.Fields("line_number").Value
        vDLUFields(DLUFields.dfNetAmount).Value = CStr(CDbl(vRecordSet.Fields("net_amount").Value) * -1) 'reverse out any reversals
        mvEnv.Connection.InsertRecord("declaration_lines_unclaimed", vDLUFields)
      End While
      vRecordSet.CloseRecordSet()

      'BR18914 Add any orphaned negative Adjustments, these can only be created by import.
      vDLUFields = New CDBFields
      vWhereFields = New CDBFields

      vDLUFields.Add("cd_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value)
      vDLUFields.Add("contact_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value)
      vDLUFields.Add("batch_number", CDBField.FieldTypes.cftLong, 0)
      vDLUFields.Add("transaction_number", CDBField.FieldTypes.cftInteger, 0)
      vDLUFields.Add("line_number", CDBField.FieldTypes.cftInteger, 0)
      vDLUFields.Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
      vDLUFields.Add("net_amount", CDBField.FieldTypes.cftNumeric)

      'Build WhereFields collection for both SQL statements
      vWhereFields.Add("fh.contact_number", CDBField.FieldTypes.cftLong, vContactNos, CType(IIf(InStr(vContactNos, ",") > 0, CDBField.FieldWhereOperators.fwoIn, CDBField.FieldWhereOperators.fwoEqual), CDBField.FieldWhereOperators))
      If BatchNumber > 0 Then
        vWhereFields.Add("fh.batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
        vWhereFields.Add("fh.transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
      End If
      If Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
        vWhereFields.Add("fh.transaction_date", CDBField.FieldTypes.cftDate, StartDate, CDBField.FieldWhereOperators.fwoBetweenFrom)
        vWhereFields.Add("fh.transaction_date#2", CDBField.FieldTypes.cftDate, EndDate, CDBField.FieldWhereOperators.fwoBetweenTo)
      Else
        vWhereFields.Add("fh.transaction_date", CDBField.FieldTypes.cftDate, StartDate, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      End If
      If IsDate(CancelledOn) Then
        vWhereFields.Add("fh.transaction_date#3", CDBField.FieldTypes.cftDate, CancelledOn, CDBField.FieldWhereOperators.fwoLessThan)
      End If
      vWhereFields.Add("fh.payment_method", CDBField.FieldTypes.cftCharacter, mvCAFPayMethod, CDBField.FieldWhereOperators.fwoNotEqual)
      vWhereFields.Add("fh.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoLessThan)
      vWhereFields.Add("tt.transaction_sign", CDBField.FieldTypes.cftCharacter, "C")
      vWhereFields.Add("fhd.amount", CDBField.FieldTypes.cftNumeric, "0", CDBField.FieldWhereOperators.fwoLessThan)
      vWhereFields.Add("p.eligible_for_gift_aid", CDBField.FieldTypes.cftCharacter, "Y")
      vWhereFields.Add("lbr.batch_number", CDBField.FieldTypes.cftLong, "")
      vWhereFields.Add("gph.batch_number", CDBField.FieldTypes.cftLong, "")
      vWhereFields.Add("r.line_number", CDBField.FieldTypes.cftNumeric, "")

      'Build Inner & Outer Joins used by both SQL statements
      vSQLInnerJoinTT = " INNER JOIN transaction_types tt ON fh.transaction_type = tt.transaction_type"
      vSQLInnerJoin = " INNER JOIN financial_history_details fhd ON %1.batch_number = fhd.batch_number AND %2.transaction_number = fhd.transaction_number %3"
      vSQLInnerJoin = vSQLInnerJoin & " INNER JOIN products p ON fhd.product = p.product"

      vSQLOuterJoin = " LEFT OUTER JOIN batch_transactions bt ON fh.batch_number = bt.batch_number AND fh.transaction_number = bt.transaction_number"
      vSQLOuterJoin = vSQLOuterJoin & " LEFT OUTER JOIN legacy_bequest_receipts lbr ON fhd.batch_number = lbr.batch_number AND fhd.transaction_number = lbr.transaction_number AND fhd.line_number = lbr.line_number"
      vSQLOuterJoin = vSQLOuterJoin & " LEFT OUTER JOIN gaye_pledge_payment_history gph ON fh.batch_number = gph.batch_number AND fh.transaction_number = gph.transaction_number"
      vSQLOuterJoin = vSQLOuterJoin & " LEFT OUTER JOIN reversals r ON fh.batch_number = r.batch_number AND fh.transaction_number = r.transaction_number"

      vSQLJoinCov = " LEFT OUTER JOIN (SELECT oph.batch_number, oph.transaction_number, oph.line_number, covenant_number, o.order_number AS cov_order_number, order_type AS cov_order_type,"
      vSQLJoinCov = vSQLJoinCov & " oph.payment_number AS oph_payment_number, c.payment_number AS cov_payment_number,"
      vSQLJoinCov = vSQLJoinCov & " o.eligible_for_gift_aid AS cov_eligible_ga, " & mvEnv.Connection.DBAddYears("c.start_date", "c.covenant_term") & " AS cov_expiry"
      vSQLJoinCov = vSQLJoinCov & " FROM order_payment_history oph INNER JOIN financial_history fh ON oph.batch_number = fh.batch_number AND oph.transaction_number = fh.transaction_number"
      vSQLJoinCov = vSQLJoinCov & " INNER JOIN covenants c ON oph.order_number = c.order_number INNER JOIN orders o ON c.order_number = o.order_number"
      vSQLJoinCov = vSQLJoinCov & " WHERE c.start_date <= fh.transaction_date"
      vSQLJoinCov = vSQLJoinCov & " AND ((c.cancelled_on IS NULL) OR (c.cancelled_on IS NOT NULL AND c.cancelled_on >= fh.transaction_date))"
      vSQLJoinCov = vSQLJoinCov & " ) cv ON fhd.batch_number = cv.batch_number AND fhd.transaction_number = cv.transaction_number AND fhd.line_number = cv.line_number"

      vSQLBTWhere = " AND (bt.eligible_for_gift_aid IS NULL OR (bt.eligible_for_gift_aid IS NOT NULL AND bt.eligible_for_gift_aid = 'Y'))"

      'Donation payments
      If (Me.DeclarationType And GiftAidDeclarationTypes.gadtDonation) = GiftAidDeclarationTypes.gadtDonation Then 'Declaration Type is Donation or All
        vDLUFields(DLUFields.dfBatchNumber).Value = CStr(0) 'Force it to be different
        vPrevGiftMem = False
        vPrevPPNo = 0
        vPrevCovNo = 0
        vPrevCovPPNo = 0
        vPrevCovType = ""
        vPrevCovOphPayNo = 0
        vPrevCovPayNo = 0
        vPrevCovEligGA = ""
        vPrevCovExpiry = ""

        vSQLDon = "SELECT /* SQLServerCSC */ fhd.batch_number,fhd.transaction_number,fhd.line_number,fhd.amount, fh.transaction_date, oph.order_number, covenant_number, cov_order_number, cov_order_type, oph_payment_number, cov_payment_number, cov_eligible_ga, cov_expiry"
        vSQLDon = vSQLDon & " FROM financial_history fh"
        vSQLDon = vSQLDon & vSQLInnerJoinTT
        vSQLDon = vSQLDon & Replace(Replace(Replace(vSQLInnerJoin, "%1", "fh"), "%2", "fh"), "%3", "")
        vSQLDon = vSQLDon & vSQLOuterJoin
        vSQLDon = vSQLDon & " LEFT OUTER JOIN (SELECT batch_number, transaction_number, line_number, oph.order_number, eligible_for_gift_aid FROM order_payment_history oph INNER JOIN orders o ON oph.order_number = o.order_number) oph ON fhd.batch_number = oph.batch_number AND fhd.transaction_number = oph.transaction_number AND fhd.line_number = oph.line_number"
        vSQLDon = vSQLDon & vSQLJoinCov
        vSQLDon = vSQLDon & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & vSQLBTWhere
        vSQLDon = vSQLDon & " AND (oph.eligible_for_gift_aid IS NULL OR (oph.eligible_for_gift_aid IS NOT NULL AND oph.eligible_for_gift_aid = 'Y'))"
        vSQLDon = vSQLDon & " ORDER BY fhd.batch_number,fhd.transaction_number,fhd.line_number"

        Dim v As New SQLStatement(mvEnv.Connection, mvEnv.Connection.ProcessAnsiJoins(vSQLDon))
        vTable = New CDBDataTable
        vTable.FillFromSQLDONOTUSE(mvEnv, mvEnv.Connection.ProcessAnsiJoins(vSQLDon))

        For Each vRow In vTable.Rows
          ProcessUnclaimed(vDLUFields, vRow.IntegerItem("batch_number"), vRow.IntegerItem("transaction_number"), vRow.IntegerItem("line_number"), Val(vRow.Item("amount")), vRow.Item("transaction_date"), vPrevCovNo, vPrevCovPPNo, vPrevCovType, vPrevCovOphPayNo, vPrevCovPayNo, vPrevCovEligGA, vPrevCovExpiry, vPrevPPNo)
          'ProcessUnclaimed will process the previous transaction data so need to store some values for checking next time
          vPrevPPNo = vRow.IntegerItem("order_number")
          vPrevCovNo = vRow.IntegerItem("covenant_number")
          vPrevCovPPNo = vRow.IntegerItem("cov_order_number")
          vPrevCovType = vRow.Item("cov_order_type")
          vPrevCovOphPayNo = vRow.IntegerItem("oph_payment_number")
          vPrevCovPayNo = vRow.IntegerItem("cov_payment_number")
          vPrevCovEligGA = vRow.Item("cov_eligible_ga")
          vPrevCovExpiry = vRow.Item("cov_expiry")
        Next vRow
        ProcessUnclaimed(vDLUFields, 0, 0, 0, 0, "", vPrevCovNo, vPrevCovPPNo, vPrevCovType, vPrevCovOphPayNo, vPrevCovPayNo, vPrevCovEligGA, vPrevCovExpiry, vPrevPPNo) 'Write any remaining dlu line details as this will group batch/trans numbers
      End If


      'Delete any of the just created unclaimed lines data if they exist in the claimed lines table
      'for same batch/transaction/line numbers and declaration_or_covenant_number = 'D'
      vDLU.Init()
      vSQL = "SELECT " & vDLU.GetRecordSetFields() & " FROM declaration_lines_unclaimed dlu, declaration_tax_claim_lines dtcl"
      vSQL = vSQL & " WHERE dlu.cd_number = " & Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value & " AND dlu.declaration_or_covenant_number = 'D'"
      vSQL = vSQL & " AND dlu.batch_number = dtcl.batch_number AND dlu.transaction_number = dtcl.transaction_number AND dlu.line_number = dtcl.line_number"
      vSQL = vSQL & " AND dtcl.declaration_or_covenant_number = 'D'"
      vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
      While vRecordSet.Fetch() = True
        vDLU = New DeclarationLinesUnclaimed(Me.Environment)
        vDLU.InitFromRecordSet(vRecordSet)
        vDLU.Delete()
      End While
      vRecordSet.CloseRecordSet()

    End Sub

    Private Sub ProcessUnclaimed(ByRef pDLUFields As CDBFields, ByVal pBatchNo As Integer, ByVal pTransactionNo As Integer, ByVal pLineNo As Integer, ByVal pAmount As Double, ByVal pTransactionDate As String, ByVal pCovNumber As Integer, ByVal pCvPlanNumber As Integer, ByVal pCvPlanTpe As String, ByVal pCvOphPayNumber As Integer, ByVal pCvPayNumber As Integer, ByVal pCvEligibleGA As String, ByVal pCvExpiry As String, Optional ByVal pPayPlanNumber As Integer = 0, Optional ByVal pGiftMembership As Boolean = False)
      Dim vIgnore As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vGAD As GiftAidDeclaration
      Dim vNewAmount As Double
      Dim vPPNumber As Integer
      Dim vPayPlan As PaymentPlan

      If pDLUFields(DLUFields.dfBatchNumber).IntegerValue = pBatchNo And CDbl(pDLUFields(DLUFields.dfTransactionNumber).Value) = pTransactionNo And CDbl(pDLUFields(DLUFields.dfLineNumber).Value) = pLineNo Then
        'It is the same fhd line as last time so just total the amount
        pDLUFields(DLUFields.dfNetAmount).Value = CStr(Val(pDLUFields(DLUFields.dfNetAmount).Value) + Val(CStr(pAmount)))
      Else
        'It is a new fhd line so check if we need to save the old one
        If CDbl(pDLUFields(DLUFields.dfBatchNumber).Value) > 0 Then
          'Check for a valid covenant at the time of the transaction
          If pCvPlanNumber > 0 Then
            'We have a Covenant
            If DateAdd(Microsoft.VisualBasic.DateInterval.Day, Val(mvEnv.GetConfig("cv_no_days_claim_grace")), CDate(pCvExpiry)) >= mvPrevTransDate Then
              'Transaction date is on/before Covenant expiry date + grace days so payment claimed under Covenant
              'Note: This is the same test used by Batch Posting when posting the payment.
              vIgnore = True
            End If
            If vIgnore Then
              If pCvEligibleGA <> "N" Then
                'Create new lines for overpaid Covenants
                vNewAmount = CheckForNewDluLines(pDLUFields, pDLUFields("net_amount").DoubleValue, pCovNumber, pCvPlanNumber, pCvPlanTpe, pCvPayNumber, pCvOphPayNumber)
                If vNewAmount <> 0 Then
                  vIgnore = False
                  pDLUFields("net_amount").Value = CStr(vNewAmount)
                End If
              End If
            End If
          End If

          If Not vIgnore And pPayPlanNumber > 0 Then
            'This is a membership Payment Plan payment
            If pGiftMembership Then
              If mvProcPaymentPlans.Exists(CStr(pPayPlanNumber)) Then
                vPayPlan = CType(mvProcPaymentPlans.Item(CStr(pPayPlanNumber)), PaymentPlan)
              Else
                vPayPlan = New PaymentPlan
                vPayPlan.Init(mvEnv, pPayPlanNumber)
                mvProcPaymentPlans.Add(vPayPlan, CStr(pPayPlanNumber))
              End If
              vIgnore = Not (vPayPlan.EligibleForGiftAid)
              If vIgnore = False Then vIgnore = Not (vPayPlan.MembershipEligibleForGiftAid(CStr(mvPrevTransDate)))
            End If
          End If

          If Not vIgnore Then
            'Check if there are linked Declarations which this payment may be claimed under
            If mvGAPayPlanDecs = True Or mvGATransDecs = True Then
              '1. Select any Order Payment History for this transaction
              If mvGAPayPlanDecs Then vPPNumber = pPayPlanNumber

              '2. See if this transaction is linked to the current Declaration
              '   Declarations linked to transactions only select that transaction
              If Me.ClassFields(GiftAidDeclarationFields.gadfPaymentPlanNumber).IntegerValue > 0 Then
                If Me.ClassFields(GiftAidDeclarationFields.gadfPaymentPlanNumber).IntegerValue <> vPPNumber Then vIgnore = True
              End If

              '3. Now see if this transaction is linked to any existing Declaration
              If vIgnore = False Then
                vGAD = New GiftAidDeclaration
                vGAD.Init(mvEnv)
                For Each vGAD In mvAllGADecs
                  'The current Declaration will be excluded if adding a new Declaration
                  If vGAD.DeclarationNumber = Me.ClassFields(GiftAidDeclarationFields.gadfDeclarationNumber).IntegerValue Then
                    '
                  Else
                    If mvGAPayPlanDecs = True And vPPNumber > 0 Then
                      If vGAD.PaymentPlanNumber = vPPNumber Then vIgnore = True
                    End If
                    If mvGATransDecs = True And vIgnore = False Then
                      If vGAD.BatchNumber = pDLUFields(DLUFields.dfBatchNumber).IntegerValue And vGAD.TransactionNumber = pDLUFields(DLUFields.dfTransactionNumber).IntegerValue Then
                        vIgnore = True
                      End If
                    End If
                  End If
                  If vIgnore = True Then Exit For
                Next vGAD
              End If
            End If
          End If

          If Not vIgnore Then
            'Check if unclaimed line already exists
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT cd_number, net_amount, declaration_or_covenant_number FROM declaration_lines_unclaimed dlu WHERE dlu.batch_number = " & pDLUFields(3).IntegerValue & " AND dlu.transaction_number = " & pDLUFields(4).Value & " AND dlu.line_number = " & pDLUFields(5).Value & " AND declaration_or_covenant_number = 'D'")
            If vRecordSet.Fetch() = True Then
              'Batch/Trans/Line combination exists - only ignore if signs are same
              If (CDbl(vRecordSet.Fields("net_amount").Value) >= 0 And CDbl(pDLUFields(7).Value) >= 0) Or (CDbl(vRecordSet.Fields("net_amount").Value) < 0 And CDbl(pDLUFields(7).Value) < 0) Then
                'This batch/trans/line exists - update the net_amount
                vIgnore = True
                If pDLUFields(1).IntegerValue = vRecordSet.Fields(1).IntegerValue And vRecordSet.Fields(3).Value = "D" Then
                  'Only update if the current declaration matches the one the trans is against
                  vFields = New CDBFields
                  vWhereFields = New CDBFields
                  vFields.Add("net_amount", CDBField.FieldTypes.cftNumeric, Val(pDLUFields(7).Value) + Val(vRecordSet.Fields("net_amount").Value))
                  With vWhereFields
                    .Add("batch_number", pDLUFields(3).IntegerValue, CDBField.FieldWhereOperators.fwoEqual)
                    .Add("transaction_number", CDBField.FieldTypes.cftInteger, pDLUFields(4).Value, CDBField.FieldWhereOperators.fwoEqual)
                    .Add("line_number", CDBField.FieldTypes.cftInteger, pDLUFields(5).Value, CDBField.FieldWhereOperators.fwoEqual)
                  End With
                  mvEnv.Connection.UpdateRecords("declaration_lines_unclaimed", vFields, vWhereFields)
                End If
              End If
            End If
            vRecordSet.CloseRecordSet()
          End If

          If Not vIgnore Then mvEnv.Connection.InsertRecord("declaration_lines_unclaimed", pDLUFields)
        End If
        'now set up the new one
        pDLUFields(DLUFields.dfBatchNumber).Value = CStr(pBatchNo)
        pDLUFields(DLUFields.dfTransactionNumber).Value = CStr(pTransactionNo)
        pDLUFields(DLUFields.dfLineNumber).Value = CStr(pLineNo)
        pDLUFields(DLUFields.dfNetAmount).Value = CStr(pAmount)
        If pTransactionDate.Length > 0 Then mvPrevTransDate = CDate(pTransactionDate)
      End If
    End Sub

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedOn).Value = pAmendedOn
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub
    Public Sub SetCreated(ByRef pCreatedOn As String, ByRef pCreatedBy As String)
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfCreatedOn).Value = pCreatedOn
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfCreatedBy).Value = pCreatedBy
    End Sub
    Private Function CheckForNewDluLines(ByVal pDLUFields As CDBFields, ByVal pAmount As Double, ByVal pCovNumber As Integer, ByVal pCvPlanNumber As Integer, ByVal pCvPlanTpe As String, ByVal pCvPayNumber As Integer, ByVal pCvOphPayNumber As Integer) As Double
      Dim vRS As CDBRecordSet
      Dim vAmount As Double
      Dim vDiff As Double
      Dim vFound As Boolean
      Dim vFrom As String
      Dim vSQL As String
      Dim vWhere As String

      vFrom = "FROM orders o, order_payment_history oph, declaration_tax_claim_lines dtcl"
      vWhere = "AND oph.order_number = o.order_number AND oph.batch_number = " & pDLUFields("batch_number").IntegerValue
      vWhere = vWhere & " AND oph.transaction_number = " & pDLUFields("transaction_number").IntegerValue & " And oph.line_number = " & pDLUFields("line_number").IntegerValue
      vWhere = vWhere & " AND dtcl.batch_number = oph.batch_number"
      vWhere = vWhere & " AND dtcl.transaction_number = oph.transaction_number AND dtcl.line_number = oph.line_number"
      vWhere = vWhere & " AND dtcl.cd_number = " & pCovNumber & " AND dtcl.declaration_or_covenant_number = 'C'"

      If (Me.DeclarationType And GiftAidDeclarationTypes.gadtMember) = GiftAidDeclarationTypes.gadtMember Then 'Declaration Type is Member or All
        'Check for a membership order
        If pCvPlanTpe = "M" Then
          vSQL = "SELECT oph.amount, dtcl.net_amount"
          vSQL = vSQL & " " & vFrom & " WHERE o.order_number = " & pCvPlanNumber
          vSQL = vSQL & " AND o.order_type = 'M' " & vWhere
          vRS = mvEnv.Connection.GetRecordSet(vSQL)
          While vRS.Fetch() = True
            vFound = True
            vDiff = vRS.Fields("amount").DoubleValue - vRS.Fields("net_amount").DoubleValue
            If vDiff <> 0 Then
              vAmount = vAmount + vDiff
            End If
          End While
          vRS.CloseRecordSet()
        End If
      End If

      If Me.DeclarationType = GiftAidDeclarationTypes.gadtDonation Or (Me.DeclarationType = GiftAidDeclarationTypes.gadtAll And vFound = False) Then
        'Must have some donation products
        vSQL = "SELECT oph.amount,dtcl.net_amount,fhd.amount AS fhd_amount "
        vSQL = vSQL & vFrom & ",financial_history_details fhd,products p"
        vSQL = vSQL & " WHERE o.order_number = " & pCvPlanNumber & " " & vWhere
        vSQL = vSQL & " AND fhd.batch_number = dtcl.batch_number AND fhd.transaction_number = dtcl.transaction_number"
        vSQL = vSQL & " AND fhd.line_number = dtcl.line_number AND p.product = fhd.product"
        vSQL = vSQL & " AND p.donation = 'Y'"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataProductEligibleGA) Then
          vSQL = vSQL & " AND p.eligible_for_gift_aid = 'Y'"
        End If
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch() = True
          vFound = True
          vDiff = vRS.Fields("amount").DoubleValue - vRS.Fields("net_amount").DoubleValue
          If vDiff <> 0 Then
            vAmount = vAmount + vRS.Fields("fhd_amount").DoubleValue
          End If
        End While
        If vAmount > vDiff Then vAmount = vDiff
        vRS.CloseRecordSet()
      End If

      If vAmount > pDLUFields("net_amount").DoubleValue Then vAmount = pDLUFields("net_amount").DoubleValue

      If vFound = False Then
        'Looks like no dtcl line for this payment
        If pCvPayNumber >= pCvOphPayNumber Then
          'Payment included in Tax Claim but none of the amount used
          vAmount = pAmount
        End If
      End If

      CheckForNewDluLines = vAmount

    End Function
    Public Sub InitNumberOnly(ByVal pEnv As CDBEnvironment, ByRef pDeclarationNumber As Integer)
      Init(pEnv)
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value = CStr(pDeclarationNumber)
    End Sub

    Friend Sub InitNewFromMergedDeclaration(ByVal pEnv As CDBEnvironment, ByVal pMergedDeclaration As GiftAidDeclaration, ByVal pStartDate As String, ByVal pEndDate As String)
      Dim vNotes As String

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()

      With pMergedDeclaration
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).Value = CStr(.ContactNumber)
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationDate).Value = .DeclarationDate
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationType).Value = .DeclarationTypeCode
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfSource).Value = .Source
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfConfirmedOn).Value = .ConfirmedOn
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfMethod).Value = .MethodCode
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfStartDate).Value = pStartDate
        If pEndDate.Length > 0 Then Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value = pEndDate
        vNotes = .Notes
        If vNotes.Length > 0 Then vNotes = vNotes & vbCrLf & vbCrLf
        vNotes = vNotes & String.Format(ProjectText.String18543, CStr(.DeclarationNumber)) 'Supercedes Gift Aid Declaration %s following Contact merge.
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfNotes).Value = vNotes
      End With

      mvContact.InitRecordSetType(mvEnv, Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtAddress, ContactNumber)

    End Sub

    Public Sub Update(ByVal pStartDate As String, ByVal pEndDate As String, ByVal pDeclarationType As GiftAidDeclarationTypes, ByVal pNotes As String)
      'Update the Gift Aid Declaration (only used by Contact Merge at the moment)
      'Note - the EndDate is allowed to be null

      With Me.ClassFields
        .Item(GiftAidDeclarationFields.gadfStartDate).Value = pStartDate
        .Item(GiftAidDeclarationFields.gadfEndDate).Value = pEndDate
        .Item(GiftAidDeclarationFields.gadfDeclarationType).Value = DeclarationTypeCodeFromType(pDeclarationType)
        .Item(GiftAidDeclarationFields.gadfNotes).Value = pNotes
      End With
    End Sub
    Public Function DeclarationTypeCodeFromType(ByRef pDeclarationType As GiftAidDeclarationTypes) As String
      Return "A"
    End Function
    Public Function DeclarationTypeFromCode(ByRef pDeclarationTypeCode As String) As GiftAidDeclarationTypes
      Return GiftAidDeclarationTypes.gadtAll
    End Function
    Public Function CheckOverlapDeclaration(ByVal pContactNumber As Integer, ByVal pStartDate As String, Optional ByVal pEndDate As String = "", Optional ByVal pPayPlanNumber As Integer = 0, Optional ByVal pBatchNumber As Integer = 0) As Integer
      Dim vCount As Integer = 0
      Dim vSQL As String = "select * from gift_aid_declarations where "

      If pPayPlanNumber = 0 And pBatchNumber = 0 Then
        vSQL &= "contact_number = " & pContactNumber & " AND declaration_number <> " & DeclarationNumber
        If pEndDate.Length > 0 Then
          vSQL &= " AND ((start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pStartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pEndDate) & ") OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pStartDate) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pEndDate) & " OR end_date IS NULL)) OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pStartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pStartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pEndDate) & ") OR (start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pStartDate) & " AND start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pEndDate) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pEndDate) & " OR end_date IS NULL)))"
        Else
          vSQL &= " AND (end_date IS NULL OR end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pStartDate) & ")"
        End If

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataGiftAidMergeCancellation) Then
          If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason).Length > 0 Then
            vSQL &= " AND (cancellation_reason IS NULL or cancellation_reason <> '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason) & "')"
          End If
        End If
        vSQL &= " AND order_number IS NULL AND batch_number IS NULL"
        Dim vData As DataTable = New SQLStatement(mvEnv.Connection, vSQL).GetDataTable
        vCount = (From vRow As DataRow In vData
                  Where IsDBNull(vRow("end_date")) OrElse
                        IsDBNull(vRow("cancelled_on")) OrElse
                        IsDBNull(vRow("cancellation_reason")) OrElse
                        Not vRow.Field(Of Date)("end_date").Equals(vRow.Field(Of Date)("cancelled_on")) OrElse
                        Not vRow.Field(Of Date)("end_date").Equals(CDate(pStartDate)) OrElse
                        String.IsNullOrWhiteSpace(vRow.Field(Of String)("cancellation_reason")) OrElse
                        vRow.Field(Of String)("cancellation_reason").Equals(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason))
                  Select vRow).Count
      End If
      Return vCount
    End Function

    Private Function GetBackClaimYears() As Integer
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vBackClaimYears As Integer

      vSQL = "SELECT back_claim_years FROM gift_aid_controls"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      If vRS.Fetch() = True Then
        vBackClaimYears = vRS.Fields(1).IntegerValue
      End If
      vRS.CloseRecordSet()

      GetBackClaimYears = vBackClaimYears

    End Function

    Public Sub Cancel(ByVal pCancelReason As String, ByVal pCancellationSource As String, Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByVal pNewEndDate As String = "", Optional ByVal pPreventDLURecreation As Boolean = False)
      Dim vContacts As New Collection
      Dim vContact As New Contact(mvEnv)
      Dim vCreateLines As Boolean
      Dim vCancellationRS As CDBRecordSet

      vCancellationRS = mvEnv.Connection.GetRecordSet("SELECT status,cancellation_reason_desc FROM cancellation_reasons WHERE cancellation_reason = '" & pCancelReason & "'")
      If vCancellationRS.Fetch() = True Then
        If vCancellationRS.Fields(1).Value.Length > 0 Then
          vContact.Init((Me.ClassFields.Item(GiftAidDeclarationFields.gadfContactNumber).IntegerValue))
          If vContact.Existing Then
            vContact.Status = vCancellationRS.Fields(1).Value
            If String.IsNullOrWhiteSpace(vContact.StatusReason) Then
              vContact.StatusReason = vCancellationRS.Fields(2).Value
            End If
            vContact.StatusDate = TodaysDate()
            vContact.Save()
          End If
        End If
      End If
      vCancellationRS.CloseRecordSet()

      Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationReason).Value = pCancelReason
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancelledBy).Value = mvEnv.User.UserID
      Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancelledOn).Value = TodaysDate()
      If Len(pCancellationSource) > 0 Then Me.ClassFields.Item(GiftAidDeclarationFields.gadfCancellationSource).Value = pCancellationSource
      If Len(pNewEndDate) = 0 Then pNewEndDate = CancelledOn

      If EndDate.Length > 0 Then
        If CDate(EndDate) > CDate(pNewEndDate) Then 'CDate(CancelledOn) Then
          Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value = pNewEndDate 'CancelledOn
          vCreateLines = True
        ElseIf CDate(CancelledOn) <= CDate(EndDate) Then
          vCreateLines = True
        End If
      Else
        Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value = pNewEndDate 'CancelledOn
        vCreateLines = True
      End If

      If vCreateLines = True And pPreventDLURecreation = False Then
        RemoveUnclaimedLines()
        GenerateUnclaimedLines()
      End If
      Dim vSaveOptions As New SaveOptions() With
      {
          .AmendedBy = pAmendedBy,
          .Audit = pAudit,
          .BatchNumber = pBatchNumber,
          .TransactionNumber = pTransactionNumber,
          .Cancel = True
      }
      Save(vSaveOptions)
    End Sub

    Public Sub RemoveUnclaimedLines()
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("cd_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value)
      vWhereFields.Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
      mvEnv.Connection.DeleteRecords("declaration_lines_unclaimed", vWhereFields, False)
    End Sub

    Public Function CanCancel(ByRef pCancelMessage As String, Optional ByVal pNewEndDate As String = "") As Boolean
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vCanCancel As Boolean = True


      If Not String.IsNullOrWhiteSpace(Me.CancellationReason) Then
        pCancelMessage = ProjectText.GiftAidAlreadyCancelled
        vCanCancel = False
      End If

      If vCanCancel Then
        'Check the declaration_tax_claim_lines where the transaction date is >= the new end date
        'and see if the line has been reversed
        If Len(pNewEndDate) = 0 Then pNewEndDate = TodaysDate()

        vSQL = "SELECT fh.status AS fh_status, fhd.status AS fhd_status FROM declaration_tax_claim_lines dtcl, financial_history fh, financial_history_details fhd"
        vSQL = vSQL & " WHERE cd_number = " & DeclarationNumber & " AND declaration_or_covenant_number = 'D' AND fh.batch_number = dtcl.batch_number"
        vSQL = vSQL & " AND fh.transaction_number = dtcl.transaction_number AND transaction_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pNewEndDate) ' TodaysDate)
        vSQL = vSQL & " AND fhd.batch_number = fh.batch_number AND fhd.transaction_number = fh.transaction_number AND fhd.line_number = dtcl.line_number"
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch() = True And vCanCancel = True
          If Len(vRS.Fields("fh_status").Value) > 0 Then
            If vRS.Fields("fh_status").Value <> "R" Then
              If vRS.Fields("fhd_status").Value <> "R" Then
                vCanCancel = False
              End If
            End If
          Else
            vCanCancel = False
          End If
        End While
        vRS.CloseRecordSet()
        If Not vCanCancel Then pCancelMessage = (ProjectText.String18544) 'There are claimed transactions dated on or after the new end date.  The Declaration can not be cancelled.
      End If
      CanCancel = vCanCancel
    End Function

    Public Function Validate(ByRef pMessage As String) As Boolean
      'Checks that there are no over-lapping Declarations
      'Returns any messages showing why it can not be saved
      Dim vCov As New Covenant
      Dim vFields As New CDBFields
      Dim vPayPlan As New PaymentPlan
      Dim vExpiry As String
      Dim vMsg As String
      Dim vValid As Boolean
      Dim vWhere As String

      vValid = True
      vMsg = (ProjectText.String18545) 'This Declaration can not be %1 as

      If EndDate.Length > 0 Then
        vWhere = " AND ((start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, EndDate) & ") OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, StartDate) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, EndDate) & " OR end_date IS NULL)) OR (start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, StartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartDate) & " AND end_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, EndDate) & ") OR (start_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartDate) & " AND start_date " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, EndDate) & " AND (end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, EndDate) & " OR end_date IS NULL)))"
      Else
        vWhere = " AND (end_date IS NULL OR end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, StartDate) & ")"
      End If

      If PaymentPlanNumber > 0 Then
        '1. Check no un-linked declarations
        With vFields
          .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          .Add("declaration_number", DeclarationNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          .Add("order_number", CDBField.FieldTypes.cftLong, PaymentPlanNumber)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
        End With
        If mvEnv.Connection.GetCount("gift_aid_declarations", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
          vValid = False
          vMsg = vMsg & (ProjectText.String18546) ' the date range has already been covered by another Declaration
        End If

        If vValid Then
          '2. Check no over-lapping Dec's linked to a payment plan
          vFields.FieldExists("order_number").Value = CStr(PaymentPlanNumber)
          If mvEnv.Connection.GetCount("gift_aid_declarations", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
            vValid = False
            vMsg = vMsg & (ProjectText.String18547) ' there is already another Declaration linking to this Payment Plan covered by the date range.
          End If
        End If

        If vValid Then
          '3. Check no Dec's already linked to a payment made for this payment plan
          With vFields
            .Remove("order_number")
            .Remove("batch_number")
            .Add("gad.order_number", CDBField.FieldTypes.cftLong)
            .Add("gad.batch_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
            .Add("oph.batch_number", CDBField.FieldTypes.cftLong, "gad.batch_number")
            .Add("oph.transaction_number", CDBField.FieldTypes.cftLong, "gad.transaction_number")
            .Add("oph.order_number", CDBField.FieldTypes.cftLong, PaymentPlanNumber)
          End With
          If mvEnv.Connection.GetCount("gift_aid_declarations gad, order_payment_history oph", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
            vValid = False
            vMsg = vMsg & (ProjectText.String18548) ' some payments under the Payment Plan are already included under another Declaration.
          End If
        End If

        If vValid Then
          '4.(a) Check the dates are within the payment plan's dates
          With vPayPlan
            .Init(mvEnv, PaymentPlanNumber)
            If CDate(StartDate) < CDate(.StartDate) Then
              vValid = False
            ElseIf CDate(StartDate) > CDate(.ExpiryDate) Then
              vValid = False
            Else
              If EndDate.Length > 0 Then
                If CDate(EndDate) > CDate(.ExpiryDate) Then
                  vValid = False
                End If
              End If
            End If
          End With
          If vValid = False Then
            vMsg = vMsg & (ProjectText.String18549) ' the dates are outside of the Payment Plan dates.
          End If

          '(b) If has/had Covenant attached then ensure Declaration nor entirely covered by Covenant
          If vValid = True And vPayPlan.CovenantStatus <> PaymentPlan.ppCovenant.ppcNo Then
            vCov.InitFromPaymentPlan(mvEnv, (vPayPlan.PlanNumber))
            With vCov
              If .Existing Then
                vExpiry = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Year, .CovenantTerm, CDate(.StartDate))))
                If CDate(StartDate) >= CDate(.StartDate) Then
                  If CDate(StartDate) < CDate(vExpiry) Then
                    If EndDate.Length > 0 Then
                      If CDate(EndDate) <= CDate(vExpiry) Then
                        vValid = False
                      End If
                    End If
                  End If
                End If
              End If
            End With
            If vValid = False Then
              vMsg = vMsg & (ProjectText.String18550) ' the dates are already covered by a Covenant
            End If
          End If
        End If

      ElseIf BatchNumber > 0 Then
        '1. Check no un-linked Declarations
        With vFields
          .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          .Add("declaration_number", DeclarationNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("batch_number", CDBField.FieldTypes.cftLong)
        End With
        If mvEnv.Connection.GetCount("gift_aid_declarations", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
          vValid = False
          vMsg = vMsg & (ProjectText.String18546) ' the date range has already been covered by another Declaration
        End If

        If vValid Then
          '2. Check no Dec's already linked to this batch/transaction
          With vFields
            .FieldExists("batch_number").Value = CStr(BatchNumber)
            .Add("transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
          End With
          If mvEnv.Connection.GetCount("gift_aid_declarations", vFields) > 0 Then
            vValid = False
            vMsg = vMsg & (ProjectText.String18551) ' this payment is already linked to a Declaration.
          End If
        End If

        If vValid Then
          '3. Check no Dec's already linked to a pay plan that this payment may be paying
          'a. First, check processed transactions using order_payment_history
          With vFields
            .Clear()
            .Add("oph.batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
            .Add("oph.transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
            .Add("gad.order_number", CDBField.FieldTypes.cftLong, "oph.order_number")
            .Add("declaration_number", DeclarationNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          End With
          If mvEnv.Connection.GetCount("order_payment_history oph, gift_aid_declarations gad", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
            vValid = False
            vMsg = vMsg & (ProjectText.String18552) ' this payment is already linked to a Declaration for a Payment Plan.
          End If

          If vValid Then
            'b. Second, check unprocessed transactions using batch_transaction_analysis
            With vFields
              .Clear()
              .Add("bta.batch_number", CDBField.FieldTypes.cftLong, BatchNumber)
              .Add("bta.transaction_number", CDBField.FieldTypes.cftLong, TransactionNumber)
              .Add("bta.order_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoNotEqual)
              .Add("gad.order_number", CDBField.FieldTypes.cftLong, "bta.order_number")
              .Add("declaration_number", DeclarationNumber, CDBField.FieldWhereOperators.fwoNotEqual)
            End With
            If mvEnv.Connection.GetCount("batch_transaction_analysis bta, gift_aid_declarations gad", Nothing, mvEnv.Connection.WhereClause(vFields) & vWhere) > 0 Then
              vValid = False
              vMsg = vMsg & (ProjectText.String18552) ' this payment is already linked to a Declaration for a Payment Plan.
            End If
          End If
        End If
      Else
        'This Declaration is not linked to a payment or payment plan
      End If

      Validate = vValid
      If vValid = False Then
        pMessage = vMsg
      End If

    End Function

    Private Sub LinkedDeclarationChecks()
      'See if this contact has any Declarations that are linked to a Transaction or a Payment Plan
      Dim vRS As CDBRecordSet
      Dim vGAD As GiftAidDeclaration
      Dim vSQL As String

      mvGAPayPlanDecs = False
      mvGATransDecs = False
      mvAllGADecs = New Collection

      vGAD = New GiftAidDeclaration
      vGAD.Init(mvEnv)

      'Create a collection of all the Declarations for this Contact
      'to be used in ProcessUnclaimed
      vSQL = "SELECT " & vGAD.GetRecordSetFields(GiftAidDeclarationRecordSetTypes.gadrtNumbers) & " FROM gift_aid_declarations gad"
      vSQL = vSQL & " WHERE contact_number = " & Me.ClassFields(GiftAidDeclarationFields.gadfContactNumber).IntegerValue _
        & " and (end_date IS NULL OR (end_date IS NOT NULL AND end_date " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, TodaysDate()) & " )) " _
        & " ORDER BY declaration_number"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)

      While vRS.Fetch() = True
        vGAD = New GiftAidDeclaration
        vGAD.InitFromRecordSet(mvEnv, vRS, GiftAidDeclarationRecordSetTypes.gadrtNumbers)
        mvAllGADecs.Add(vGAD)

        If vGAD.PaymentPlanNumber <> 0 Then mvGAPayPlanDecs = True
        If vGAD.BatchNumber <> 0 Then mvGATransDecs = True
      End While
      vRS.CloseRecordSet()

      'Make sure that current Declaration is also included
      If BatchNumber > 0 Then mvGATransDecs = True
      If PaymentPlanNumber > 0 Then mvGAPayPlanDecs = True
    End Sub

    Public Sub PopulateUnclaimedLines(Optional ByVal pCheckLinkedDeclaration As Boolean = False)
      'This will be called here from Save, and from the Contact class during a merge
      Dim vWhereFields As New CDBFields
      Dim vTrans As Boolean

      If pCheckLinkedDeclaration Then LinkedDeclarationChecks()

      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If

      If mvExisting Then
        'delete recreate unclaimed lines
        vWhereFields.Add("cd_number", CDBField.FieldTypes.cftLong, Me.ClassFields.Item(GiftAidDeclarationFields.gadfDeclarationNumber).Value)
        vWhereFields.Add("declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D")
        mvEnv.Connection.DeleteRecords("declaration_lines_unclaimed", vWhereFields, False)
      End If
      If CancellationReason <> mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason) Then
        GenerateUnclaimedLines()
      End If

      If vTrans Then mvEnv.Connection.CommitTransaction()
    End Sub

    ''' <summary>If the End Date has changed, then find any claimed payments that have now become invalidated and create adjustment transactions.
    ''' The End Date will have already been changed.</summary>
    Private Sub ProcessChangeOfEndDate()
      Dim vRS As CDBRecordSet = Nothing
      Dim vDTCL As DeclarationTaxClaimLine 'Original claim line
      Dim vFH As FinancialHistory 'Original history
      Dim vProcess As Boolean
      Dim vSQL As String
      Dim vWhereFields As New CDBFields
      Dim vWhereNIFields As New CDBFields
      Dim vDelete As Boolean

      'Probably always want to delete any unclaimed adjustments
      If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) = 0 Then
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then vDelete = True
      Else
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then
          If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value), CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue)) > 0 Then vDelete = True
        End If
      End If
      If vDelete = False Then
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) = 0 Then
          If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then vDelete = True
        Else
          If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue), CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value)) > 0 Then vDelete = True
          End If
        End If
      End If
      'First delete existing unclaimed adjustments
      If vDelete Then
        GetReversalUnClaimedAdjustmentData(DeclarationNumber, "bta.batch_number,bta.transaction_number,bta.line_number", vRS, False)

        vWhereFields.Clear()
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        While vRS.Fetch() = True
          vWhereFields(1).Value = CStr(vRS.Fields(1).IntegerValue)
          vWhereFields(2).Value = CStr(vRS.Fields(2).IntegerValue)
          vWhereFields(3).Value = CStr(vRS.Fields(3).IntegerValue)
          mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields, False)
          mvEnv.Connection.DeleteRecords("financial_history_details", vWhereFields, False)
        End While
        vRS.CloseRecordSet()
      End If

      'Only need to process if the End date is now earlier than the previous End Date
      If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) = 0 Then
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then vProcess = True
      Else
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) > 0 Then
          If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value), CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue)) > 0 Then vProcess = True
        End If
      End If

      If vProcess Then
        'Create reversals
        vDTCL = New DeclarationTaxClaimLine
        vFH = New FinancialHistory
        vDTCL.Init(mvEnv)
        vFH.Init(mvEnv)
        vSQL = vFH.GetRecordSetFields(FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers) & ", " & vDTCL.GetRecordSetFields(DeclarationTaxClaimLine.DeclarationTaxClaimLineRecordSetTypes.dtclrtAll)
        vSQL = Replace(vSQL, "dtcl.batch_number", "dtcl.batch_number AS dtcl_batch_number")
        vSQL = Replace(vSQL, "dtcl.transaction_number", "dtcl.transaction_number AS dtcl_transaction_number")
        vSQL = vSQL & ", fhd.product AS fhd_product, fhd.amount AS fhd_amount"
        vSQL = "SELECT " & vSQL & " FROM declaration_tax_claim_lines dtcl, financial_history fh, financial_history_details fhd, transaction_types tt "
        vSQL = vSQL & " WHERE dtcl.cd_number = " & DeclarationNumber & " AND declaration_or_covenant_number = 'D'"
        vSQL = vSQL & " AND fh.batch_number = dtcl.batch_number AND fh.transaction_number = dtcl.transaction_number"
        If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue.Length > 0 Then
          vSQL = vSQL & " AND fh.transaction_date BETWEEN" & mvEnv.Connection.SQLLiteral("", CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value)) & "AND" & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue))
        Else
          vSQL = vSQL & " AND fh.transaction_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value))
        End If
        vSQL = vSQL & " AND fhd.batch_number = fh.batch_number AND fhd.transaction_number = fh.transaction_number"
        vSQL = vSQL & " AND fhd.line_number = dtcl.line_number "
        vSQL = vSQL & " AND fhd.status IS NULL "
        vSQL = vSQL & " AND tt.transaction_type = fh.transaction_type "
        vSQL = vSQL & " AND tt.transaction_sign = 'C' "
        vSQL = vSQL & " AND (tt.negatives_allowed = 'N' OR (tt.negatives_allowed = 'Y' AND fhd.amount > 0)) "
        vSQL = vSQL & " ORDER BY dtcl.batch_number, dtcl.transaction_number, dtcl.line_number" & mvEnv.Connection.DBForceOrder
        vRS = mvEnv.Connection.GetRecordSet(vSQL)

        CreateAdjustmentBatches(vRS, True)

      End If

      If vProcess = False Then
        'If dates have been moved forward, may need to delete previously created adjustments
        'But only if they have not already been claimed
        If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value) = 0 Then
          If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then vProcess = True
        Else
          If Len(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue), CDate(Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value)) > 0 Then vProcess = True
          End If
        End If
      End If

      If vProcess Then
        'Now create new adjustments for existing claimed adjustments
        GetReversalClaimedAdjustmentData(DeclarationNumber, vRS, False)

        CreateAdjustmentBatches(vRS, False)
      End If

    End Sub

    Private Sub CreateAdjustmentBatches(ByVal pRS As CDBRecordSet, ByVal pDatesBackwards As Boolean, Optional ByVal pUnclaimedAdjustments As Boolean = False)
      Dim vBatch As Batch = Nothing
      Dim vBT As BatchTransaction = Nothing
      Dim vBTA As BatchTransactionAnalysis = Nothing
      Dim vFHD As FinancialHistoryDetail = Nothing
      Dim vDLU As DeclarationLinesUnclaimed 'Original representation of unclaimed line
      Dim vDTCL As DeclarationTaxClaimLine 'Original claim line
      Dim vOldFH As FinancialHistory 'Original history
      Dim vOldBTA As BatchTransactionAnalysis
      Dim vBTASaved As Boolean
      Dim vOldBN As Integer 'Original Batch Number
      Dim vOldTN As Integer 'Original Transaction Number
      Dim vOldLN As Integer 'Original Line Number

      Dim vDecLineBatchNumber As Integer
      Dim vDecLinetransactionNumber As Integer
      Dim vDecLineLineNumber As Integer
      Dim vDecLineCdNumber As Integer
      Dim vDecLineNetAmount As Double

      While pRS.Fetch() = True
        If pUnclaimedAdjustments Then
          vDLU = New DeclarationLinesUnclaimed(Me.Environment)
          vDLU.InitFromRecordSet(pRS)
          vDecLineBatchNumber = vDLU.BatchNumber
          vDecLinetransactionNumber = vDLU.TransactionNumber
          vDecLineLineNumber = vDLU.LineNumber
          vDecLineCdNumber = vDLU.CdNumber
          vDecLineNetAmount = vDLU.NetAmount
        Else
          vDTCL = New DeclarationTaxClaimLine
          vDTCL.InitFromRecordSet(mvEnv, pRS, DeclarationTaxClaimLine.DeclarationTaxClaimLineRecordSetTypes.dtclrtAll)
          vDecLineBatchNumber = vDTCL.BatchNumber
          vDecLinetransactionNumber = vDTCL.TransactionNumber
          vDecLineLineNumber = vDTCL.LineNumber
          vDecLineCdNumber = vDTCL.CdNumber
          vDecLineNetAmount = vDTCL.NetAmount
        End If
        vOldBTA = New BatchTransactionAnalysis(mvEnv)
        vOldFH = New FinancialHistory
        If pDatesBackwards Then
          vOldFH.InitFromRecordSet(mvEnv, pRS, FinancialHistory.FinancialHistoryRecordSetTypes.fhrtNumbers)
        Else
          vOldBTA.InitFromRecordSet(pRS)
        End If
        vBTASaved = False

        If vOldBN = 0 Then
          'Set up a new batch
          vBatch = New Batch(mvEnv)
          vBT = New BatchTransaction(mvEnv)
          vBTA = New BatchTransactionAnalysis(mvEnv)
          vFHD = New FinancialHistoryDetail(mvEnv)
          With vBatch
            .InitNewBatch(mvEnv)
            .BatchType = Batch.BatchTypes.GiftAidClaimAdjustment
            .BankAccount = mvEnv.Connection.GetValue("SELECT bank_account FROM batches WHERE batch_number = " & vDecLineBatchNumber)
            .ReadyForBanking = True
          End With
          vBT.Init()
          vBTA.Init()
          vFHD.Init(mvEnv)
        End If

        If ((vDecLineBatchNumber = vOldBN) And (vDecLinetransactionNumber = vOldTN) And (vDecLineLineNumber = vOldLN)) Then
          'Same batch/transaction/line
          vBTA.ProductCode = "N/A"
          vFHD.Save()
        ElseIf ((vDecLineBatchNumber <> vOldBN) Or (vDecLinetransactionNumber <> vOldTN)) Then
          'Transaction has changed
          If vBT.BatchNumber > 0 Then
            vBTA.Save()
            vBT.Save()
            vBTASaved = True
            vFHD.Save()
          End If
          vOldLN = 0 'Force it to be different

          vBT = New BatchTransaction(mvEnv)
          With vBT
            .InitFromBatch(mvEnv, vBatch)
            If pDatesBackwards Then
              .ContactNumber = vOldFH.ContactNumber
              .AddressNumber = vOldFH.AddressNumber
              .TransactionDate = vOldFH.TransactionDate
              .TransactionType = vOldFH.TransactionType
              .PaymentMethod = vOldFH.PaymentMethod
            Else
              .ContactNumber = vOldBTA.ContactNumber
              .AddressNumber = vOldBTA.AddressNumber
              .TransactionDate = pRS.Fields("transaction_date").Value
              .TransactionType = pRS.Fields("transaction_type").Value
              .PaymentMethod = pRS.Fields("payment_method").Value
            End If
            .Receipt = "N"
            .EligibleForGiftAid = True
          End With
        End If

        If (vDecLineLineNumber <> vOldLN) Then
          If vBTA.BatchNumber > 0 And vBTASaved = False Then
            vBTA.Save()
            vFHD.Save()
          End If

          vBTA = New BatchTransactionAnalysis(mvEnv)
          With vBTA
            .InitFromTransaction(vBT)
            .LineType = "P" '?
            .ProductCode = pRS.Fields("fhd_product").Value
            '.ContactNumber = vBT.ContactNumber no longer required as now set in InitFromTransaction
            '.AddressNumber = vBT.AddressNumber
            .MemberNumber = vDecLineCdNumber.ToString 'DeclarationNumber
            .Amount = (vDecLineNetAmount * -1)
            .Quantity = 1
            .Source = Source
            .Notes = CStr(vDecLineBatchNumber) & "/" & CStr(vDecLinetransactionNumber) & "/" & CStr(vDecLineLineNumber)
          End With
        End If

        'Always set financial history details
        'This is required as there is no other way of keeping track of the products when this is a pay plan payment
        vFHD = New FinancialHistoryDetail(mvEnv)
        With vFHD
          .Init(mvEnv)
          .BatchNumber = vBTA.BatchNumber
          .TransactionNumber = vBTA.TransactionNumber
          .LineNumber = vBTA.LineNumber
          .ProductCode = pRS.Fields("fhd_product").Value
          .Amount = (pRS.Fields("fhd_amount").DoubleValue * -1)
          .CurrencyAmount = (pRS.Fields("fhd_amount").DoubleValue * -1)
          .CurrencyVatAmount = 0
          .Source = vBTA.Source
          .InvoicePayment = False
        End With

        vOldBN = vDecLineBatchNumber
        vOldTN = vDecLinetransactionNumber
        vOldLN = vDecLineLineNumber

      End While
      pRS.CloseRecordSet()

      If Not vBatch Is Nothing Then
        vBTA.Save()
        vBT.Save()
        vFHD.Save()
        vBatch.SetBatchTotals()
        vBatch.SetDetailComplete(Nothing, False)
        vBatch.Save()
      End If

    End Sub

    Private Sub GetReversalClaimedAdjustmentData(ByVal pDeclarationNumber As Integer, ByRef pRS As CDBRecordSet, ByVal pUseStartDate As Boolean, Optional ByVal pNegativesOnly As Boolean = False)
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vDTCL As New DeclarationTaxClaimLine
      Dim vWhereFields As New CDBFields
      Dim vDateFrom As String = ""
      Dim vDateTo As String = ""
      Dim vSQL As String

      vDTCL.Init(mvEnv)
      vBTA.Init()
      vSQL = vBTA.GetRecordSetFields()
      vSQL = Replace(Replace(Replace(vSQL, "bta.batch_number,", ""), "bta.transaction_number,", ""), "line_number,", "")
      vSQL = Replace(Replace(vSQL, "currency_vat_amount", "bta.currency_vat_amount"), "sales_contact_number", "bta.sales_contact_number")
      vSQL = vSQL & ", " & vDTCL.GetRecordSetFields(DeclarationTaxClaimLine.DeclarationTaxClaimLineRecordSetTypes.dtclrtAll) & ", transaction_date, bt.transaction_type, bt.payment_method"
      vSQL = vSQL & ", fhd.product AS fhd_product, fhd.amount AS fhd_amount"
      If DeclarationNumber <> pDeclarationNumber Then
        vSQL = Replace(vSQL, "cd_number,", "")
        vSQL = vSQL & ", " & DeclarationNumber & " AS cd_number"
      End If
      vSQL = "SELECT " & vSQL & " FROM batches b, batch_transactions bt, batch_transaction_analysis bta, declaration_tax_claim_lines dtcl, financial_history_details fhd"
      vSQL = vSQL & " WHERE b.batch_type = '" & Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment) & "'"
      vSQL = vSQL & " AND bt.batch_number = b.batch_number"
      If pUseStartDate Then
        'Start date has been changed on a Declaration before claim adjustments have been claimed
        If CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value) > CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue) Then
          vWhereFields.Add("bt.transaction_date", CDBField.FieldTypes.cftDate, Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue, CDBField.FieldWhereOperators.fwoBetweenFrom)
          vWhereFields.Add("bt.transaction_date2", CDBField.FieldTypes.cftDate, Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value, CDBField.FieldWhereOperators.fwoBetweenTo)
        Else
          vWhereFields.Add("bt.transaction_date", CDBField.FieldTypes.cftDate, Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value, CDBField.FieldWhereOperators.fwoBetweenFrom)
          vWhereFields.Add("bt.transaction_date2", CDBField.FieldTypes.cftDate, Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue, CDBField.FieldWhereOperators.fwoBetweenTo)
        End If
      Else
        If DeclarationNumber = pDeclarationNumber Then
          'The currect Declaration - look for payments for this Declaration dated after the new End Date
          If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
            If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue.Length > 0 Then
              If CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) > CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value) Then
                vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
                vDateTo = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
              Else
                vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
                vDateTo = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
              End If
            Else
              vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
            End If
          ElseIf Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then
            vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
          End If
          If vDateTo.Length > 0 Then
            vSQL = vSQL & " AND bt.transaction_date" & mvEnv.Connection.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, vDateFrom) & mvEnv.Connection.SQLLiteral("AND", CDBField.FieldTypes.cftDate, vDateTo)
          ElseIf vDateFrom.Length > 0 Then
            vSQL = vSQL & " AND bt.transaction_date" & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, vDateFrom)
          End If
        Else
          'Some other Declaration - look for payments linked to the other Declaration and dated on/after the Start Date
          If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
            vSQL = vSQL & " AND bt.transaction_date BETWEEN" & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value)) & " AND " & mvEnv.Connection.SQLLiteral("", CDBField.FieldTypes.cftDate, (Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value))
          Else
            vSQL = vSQL & " AND bt.transaction_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value))
          End If
        End If
      End If
      vSQL = vSQL & " AND bta.batch_number = bt.batch_number AND bta.transaction_number = bt.transaction_number"
      If pNegativesOnly Then vSQL = vSQL & " AND bta.amount < 0"
      vSQL = vSQL & " AND dtcl.batch_number = bta.batch_number AND dtcl.transaction_number = bta.transaction_number"
      vSQL = vSQL & " AND dtcl.line_number = bta.line_number AND dtcl.declaration_or_covenant_number = 'D'"
      vSQL = vSQL & " AND dtcl.cd_number = " & pDeclarationNumber
      vSQL = vSQL & " AND fhd.batch_number = dtcl.batch_number AND fhd.transaction_number = dtcl.transaction_number AND fhd.line_number = dtcl.line_number"
      vSQL = vSQL & " ORDER BY dtcl.batch_number, dtcl.transaction_number, dtcl.line_number"
      vSQL = vSQL & mvEnv.Connection.DBForceOrder
      pRS = mvEnv.Connection.GetRecordSet(vSQL)

    End Sub

    Private Sub GetReversalUnClaimedAdjustmentData(ByVal pDeclarationNumber As Integer, ByVal pSelectAttrs As String, ByRef pRS As CDBRecordSet, ByVal pUseStartDate As Boolean)
      Dim vWhereFields As New CDBFields
      Dim vDateFrom As String = ""
      Dim vDateTo As String = ""
      Dim vSQL As String

      If pUseStartDate Then
        'Start date has been changed on a Declaration before claim adjustments have been claimed
        If CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value) > CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue) Then
          vDateFrom = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue
          vDateTo = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value
        Else
          vDateFrom = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value
          vDateTo = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue
        End If
      Else
        If DeclarationNumber = pDeclarationNumber Then
          'Current Declaration
          'We are selecting the adjustments to be deleted = no longer supported but here to handle old data.
          If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
            If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).SetValue.Length > 0 Then
              If CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) > CDate(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value) Then
                vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
                vDateTo = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
              Else
                vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
                vDateTo = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
              End If
            Else
              vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).Value
            End If
          ElseIf Len(Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) > 0 Then
            vDateFrom = Me.ClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue
          End If
        Else
          'A new Declaration has been added and need to pick up the adjustments created under a cancelled Declaration & not yet claimed
          If Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value.Length > 0 Then
            vDateFrom = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value
            vDateTo = Me.ClassFields(GiftAidDeclarationFields.gadfEndDate).Value
          Else
            vDateFrom = Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value
          End If
        End If
      End If

      With vWhereFields
        .Add("b.batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment), CDBField.FieldWhereOperators.fwoEqual)
        If IsDate(vDateTo) Then
          .Add("bt.transaction_date", CDBField.FieldTypes.cftDate, vDateFrom, CDBField.FieldWhereOperators.fwoBetweenFrom)
          .Add("bt.transaction_date#2", CDBField.FieldTypes.cftDate, vDateTo, CDBField.FieldWhereOperators.fwoBetweenTo)
        Else
          .Add("bt.transaction_date", CDBField.FieldTypes.cftDate, vDateFrom, CDBField.FieldWhereOperators.fwoGreaterThan)
        End If
        .Add("bta.member_number", CDBField.FieldTypes.cftCharacter, pDeclarationNumber)
      End With

      vSQL = "SELECT " & pSelectAttrs & " FROM batches b"
      vSQL = vSQL & " INNER JOIN batch_transactions bt ON b.batch_number = bt.batch_number"
      vSQL = vSQL & " INNER JOIN batch_transaction_analysis bta ON bt.batch_number = bta.batch_number AND bt.transaction_number = bta.transaction_number"
      If DeclarationNumber <> pDeclarationNumber Then
        vSQL = vSQL & " INNER JOIN financial_history_details fhd ON bta.batch_number = fhd.batch_number AND bta.transaction_number = fhd.transaction_number AND bta.line_number = fhd.line_number"
      End If
      vSQL = vSQL & " LEFT OUTER JOIN declaration_tax_claim_lines dtcl ON bta.batch_number = dtcl.batch_number AND bta.transaction_number = dtcl.transaction_number AND bta.line_number = dtcl.line_number"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
      vSQL = vSQL & " AND dtcl.batch_number IS NULL"
      vSQL = vSQL & mvEnv.Connection.DBForceOrder

      pRS = mvEnv.Connection.GetRecordSetAnsiJoins(vSQL)

    End Sub

    Private Sub ProcessNewDeclarationPayments(ByVal pChangedStartDate As Boolean)
      'If we have created a new GAD and there are other GAD's dated before this one starts
      'Need to check for any claim adjustments created for the other GAD that now need to be claimed under this one
      Dim vRS As CDBRecordSet
      Dim vFields As New CDBFields
      Dim vBTA As New BatchTransactionAnalysis(mvEnv)
      Dim vSQL As String

      'Need to do two selections to see if there are claim adjustments to be reclaimed (Oracle has problems with bta.MemberNumber as it is a character field)
      'First find any GAD's, and if one found then
      'Second find any claim adjustments
      Dim vGADNumbers As New List(Of Integer)
      With vFields
        .Add("gad.contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        .Add("end_date", CDBField.FieldTypes.cftDate, Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value, CDBField.FieldWhereOperators.fwoLessThan)
        .Add("gad.batch_number")
      End With
      vSQL = "SELECT declaration_number FROM gift_aid_declarations gad"
      vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vFields) & " ORDER BY end_date desc"
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch
        vGADNumbers.Add(vRS.Fields(1).IntegerValue)
      End While
      vRS.CloseRecordSet()

      Dim vGADNo As Integer = 0
      Dim vAdjustmentsProcessed As Boolean = False
      For vIndex As Integer = 0 To vGADNumbers.Count - 1
        vGADNo = vGADNumbers(vIndex)
        With vFields
          .Clear()
          .Add("bta.member_number", CDBField.FieldTypes.cftCharacter, vGADNo)
          .Add("b.batch_number", CDBField.FieldTypes.cftLong, "bta.batch_number")
          .Add("batch_type", CDBField.FieldTypes.cftCharacter, Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment))
        End With

        If mvEnv.Connection.GetCount("batch_transaction_analysis bta, batches b", vFields) > 0 Then
          vAdjustmentsProcessed = True    'Only process data from the first GAD that had adjustment transactions

          'First check for unclaimed adjustment transactions
          vBTA.Init()
          vSQL = vBTA.GetRecordSetFields()
          vSQL = Replace(Replace(vSQL, "currency_vat_amount", "bta.currency_vat_amount"), "sales_contact_number", "bta.sales_contact_number")
          vSQL = vSQL & ", transaction_date, bt.transaction_type, bt.payment_method"
          vSQL = vSQL & ", fhd.product AS fhd_product, fhd.amount AS fhd_amount"
          vSQL = vSQL & ", " & DeclarationNumber & " AS cd_number, 'D' AS declaration_or_covenant_number, bta.amount AS net_amount"
          vSQL = Replace(vSQL, "dlu.declaration_or_covenant_number", "'D' AS dlu.declaration_or_covenant_number")

          GetReversalUnClaimedAdjustmentData(vGADNo, vSQL, vRS, pChangedStartDate)
          CreateAdjustmentBatches(vRS, False, True)

          'Second check for claimed adjustment transactions (negatives only)
          GetReversalClaimedAdjustmentData(vGADNo, vRS, pChangedStartDate, True)
          CreateAdjustmentBatches(vRS, False)
        End If
        If vAdjustmentsProcessed Then Exit For
      Next

    End Sub

    ''' <summary>If the Start Date has changed, then find any unclaimed adjustment payments that have now become invalidated and create adjustment transactions.
    ''' The Start Date will have already been changed.</summary>
    Private Sub ProcessChangeOfStartDate()
      Dim vRS As CDBRecordSet = Nothing
      Dim vProcess As Boolean
      Dim vWhereFields As CDBFields

      'Only need to process if the End date is now earlier than the previous End Date
      If CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).Value) > CDate(Me.ClassFields(GiftAidDeclarationFields.gadfStartDate).SetValue) Then
        vProcess = True
      End If

      If vProcess Then
        'Start Date is now later than before so remove any unwanted unclaimed adjustments
        GetReversalUnClaimedAdjustmentData(DeclarationNumber, "bta.batch_number,bta.transaction_number,bta.line_number", vRS, True)

        vWhereFields = New CDBFields
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("line_number", CDBField.FieldTypes.cftLong, "", CDBField.FieldWhereOperators.fwoEqual)
        While vRS.Fetch() = True
          vWhereFields(1).Value = CStr(vRS.Fields(1).IntegerValue)
          vWhereFields(2).Value = CStr(vRS.Fields(2).IntegerValue)
          vWhereFields(3).Value = CStr(vRS.Fields(3).IntegerValue)
          mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields, False)
        End While
        vRS.CloseRecordSet()

      Else
        'Start Date is now earlier then before so pick up any claim adjustments from a previous Declaration.
        ProcessNewDeclarationPayments(True)
      End If

    End Sub

    ''' <summary>If the GAD is being deleted or the dates have changed then need to delete any unclaimed adjustment transactions.
    ''' These will have been created if the GAD start over-lapped with the adjustments created for a previous GAD.</summary>
    ''' <param name="pDeleteDeclaration">The entire GAD is being deleted.</param>
    Private Sub DeleteUnclaimedAdjustments(ByVal pDeleteDeclaration As Boolean)
      Dim vStartDateChanged As Boolean = False
      Dim vEndDateChanged As Boolean = False

      If pDeleteDeclaration = False Then
        'If we are not deleting the Declaration then we only need to do this if Declaration period is now shorter than before.
        Dim vStartDate As Date = CDate(StartDate)
        If mvClassFields.Item(GiftAidDeclarationFields.gadfStartDate).ValueChanged Then
          vStartDateChanged = (vStartDate.CompareTo(CDate(mvClassFields.Item(GiftAidDeclarationFields.gadfStartDate).SetValue)) > 0)
        End If
        If mvClassFields.Item(GiftAidDeclarationFields.gadfEndDate).ValueChanged Then
          If String.IsNullOrWhiteSpace(EndDate) = False Then
            'We have an EndDate
            If String.IsNullOrWhiteSpace(mvClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue) = False Then
              'We originally had a EndDate
              Dim vEndDate As Date = Date.Parse(EndDate)
              Dim vOriginalEndDate As Date = Date.Parse(mvClassFields.Item(GiftAidDeclarationFields.gadfEndDate).SetValue)
              vEndDateChanged = (vEndDate.CompareTo(vOriginalEndDate) < 0)
            Else
              'There was originally no EndDate so date range has shortened
              vEndDateChanged = True
            End If
          End If
        End If
      End If

      If mvExisting = True AndAlso (pDeleteDeclaration = True OrElse (vStartDateChanged = True OrElse vEndDateChanged = True)) Then
        Dim vAnsiJoins As New AnsiJoins()
        vAnsiJoins.Add("batch_transactions bt", "b.batch_number", "bt.batch_number")
        vAnsiJoins.Add("batch_transaction_analysis bta", "bt.batch_number", "bta.batch_number", "bt.transaction_number", "bta.transaction_number")
        vAnsiJoins.AddLeftOuterJoin("declaration_tax_claim_lines dtcl", "bta.batch_number", "dtcl.batch_number", "bta.transaction_number", "dtcl.transaction_number", "bta.line_number", "dtcl.line_number")

        Dim vWhereFields As New CDBFields(New CDBField("b.batch_type", Batch.GetBatchTypeCode(Batch.BatchTypes.GiftAidClaimAdjustment)))
        If pDeleteDeclaration = False Then
          Dim vWhereOperator As CDBField.FieldWhereOperators
          If vStartDateChanged Then
            vWhereOperator = CDBField.FieldWhereOperators.fwoLessThan
            If vEndDateChanged Then vWhereOperator = (vWhereOperator Or CDBField.FieldWhereOperators.fwoOpenBracket)
            vWhereFields.Add("bt.transaction_date", CDBField.FieldTypes.cftDate, StartDate, vWhereOperator)
          End If
          If vEndDateChanged Then
            vWhereOperator = CDBField.FieldWhereOperators.fwoGreaterThan
            If vStartDateChanged Then vWhereOperator = (vWhereOperator Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
            vWhereFields.Add("bt.transaction_date#2", CDBField.FieldTypes.cftDate, EndDate, vWhereOperator)
          End If
        End If
        vWhereFields.Add("bta.member_number", CDBField.FieldTypes.cftCharacter, DeclarationNumber.ToString())
        vWhereFields.Add("dtcl.cd_number", CDBField.FieldTypes.cftInteger, String.Empty, CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("dtcl.cd_number#2", CDBField.FieldTypes.cftInteger, DeclarationNumber.ToString, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("dtcl.declaration_or_covenant_number", CDBField.FieldTypes.cftCharacter, "D", CDBField.FieldWhereOperators.fwoCloseBracketTwice)

        Dim vSQLStatement As New SQLStatement(Me.Environment.Connection, "bta.batch_number,bta.transaction_number,bta.line_number,dtcl.claim_number", "batches b", vWhereFields, String.Empty, vAnsiJoins)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()

        vWhereFields.Clear()
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftInteger)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftInteger)

        While vRS.Fetch() = True
          If String.IsNullOrEmpty(vRS.Fields("claim_number").Value) Then
            vWhereFields(1).Value = vRS.Fields("batch_number").Value
            vWhereFields(2).Value = vRS.Fields("transaction_number").Value
            mvEnv.Connection.DeleteRecords("batch_transactions", vWhereFields, False)

            vWhereFields.Add("line_number", vRS.Fields("line_number").IntegerValue)
            mvEnv.Connection.DeleteRecords("batch_transaction_analysis", vWhereFields, False)
            mvEnv.Connection.DeleteRecords("financial_history_details", vWhereFields, False)
            vWhereFields.Remove(3)
          End If
        End While
        vRS.CloseRecordSet()
      End If

    End Sub

    Public ReadOnly Property GADControlsExists As Boolean
      Get
        Return mvGADControlsExists
      End Get
    End Property

    Public Property Environment As CDBEnvironment
      Get
        Return mvEnv
      End Get
      Private Set(value As CDBEnvironment)
        mvEnv = value
      End Set
    End Property

    Public Property Contact As Contact
      Get
        If mvContact Is Nothing OrElse
           String.IsNullOrWhiteSpace(mvContact.Surname) Then
          mvContact.InitWithPrimaryKey(New CDBFields({New CDBField("contact_number", Me.ContactNumber)}))
        End If
        Return mvContact
      End Get
      Set(value As Contact)
        mvContact = value
      End Set
    End Property

    Public ReadOnly Property IsMandatoryDataComplete As Boolean
      Get
        Dim vResult As Boolean = True
        Me.Contact.RefreshAddress()
        If Me.Contact.Address.IsUk AndAlso
           String.IsNullOrWhiteSpace(Me.Contact.Address.Postcode) Then
          vResult = False
        End If
        If String.IsNullOrWhiteSpace(Me.Contact.Address.AddressLine) AndAlso
           String.IsNullOrWhiteSpace(Me.Contact.Address.HouseName) Then
          vResult = False
        End If
        If String.IsNullOrWhiteSpace(Me.Contact.Forenames) AndAlso
           String.IsNullOrWhiteSpace(Me.Contact.Initials) Then
          vResult = False
        End If
        Return vResult
      End Get
    End Property

    Public ReadOnly Property TableName As String
      Get
        Return Me.ClassFields.DatabaseTableName
      End Get
    End Property

    Public Sub InitWithPrimaryKey(ByVal pWhereFields As CDBFields, Optional ByVal pRSType As GiftAidDeclarationRecordSetTypes = GiftAidDeclarationRecordSetTypes.gadrtAll)
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, GetRecordSetFields(pRSType), Me.ClassFields.TableNameAndAlias, pWhereFields).GetRecordSet
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(Me.Environment, vRecordSet, pRSType)
      Else
        Init(Me.Environment)
      End If
      vRecordSet.CloseRecordSet()
    End Sub


    Public Sub LoadFromRow(pRow As DataRow) Implements IDbLoadable.LoadFromRow
      Me.InitFromDataRow(pRow, False)
    End Sub

    Private Sub InitFromDataRow(ByVal pDataRow As DataRow, ByVal pUseProperName As Boolean) 'Copied from CARERecord
      InitClassFields()
      mvExisting = True
      Dim vName As String
      For Each vClassField As ClassField In mvClassFields
        If pUseProperName Then vName = vClassField.ProperName Else vName = vClassField.Name
        If pDataRow.Table.Columns.Contains(vName) Then
          vClassField.SetValue = pDataRow.Item(vName).ToString
        End If
      Next
    End Sub

    Protected Friend Property ClassFields As ClassFields
      Get
        If mvClassFields Is Nothing Then InitClassFields()
        Return mvClassFields
      End Get
      Private Set(value As ClassFields)
        mvClassFields = value
      End Set
    End Property

    Public ReadOnly Property FieldNames As String Implements IDbSelectable.DbFieldNames
      Get
        Return Me.ClassFields.FieldNames(mvEnv, Me.ClassFields.TableAlias)
      End Get
    End Property

    Public ReadOnly Property AliasedTableName As String Implements IDbSelectable.DbAliasedTableName
      Get
        Return Me.ClassFields.TableNameAndAlias
      End Get
    End Property

    Public Class CancellationInfo
      Public Sub New(vCancelReason As String, vCancelSource As String)
        Me.CancellationReason = vCancelReason
        Me.CancellationSource = vCancelSource
      End Sub
      Public Property CancellationReason As String
      Public Property CancellationSource As String
      Public ReadOnly Property IsValid As Boolean
        Get
          Return Me.Validate()
        End Get
      End Property

      Private Function Validate() As Boolean
        Return Not String.IsNullOrWhiteSpace(Me.CancellationReason)
      End Function
    End Class

    Public Class SaveOptions
      Public Property AmendedBy As String = String.Empty
      Public Property Audit As Boolean
      Public Property BatchNumber As Integer = 0
      Public Property TransactionNumber As Integer = 0

      Public Property Cancel As Boolean = False

      Public Property CancellationInfo As CancellationInfo = Nothing

    End Class

  End Class
End Namespace
