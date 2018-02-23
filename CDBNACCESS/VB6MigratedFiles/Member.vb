

Namespace Access
  Public Class Member

    Public Enum MemberRecordSetTypes 'These are bit values
      mrtAll = &HFFS
      'ADD additional recordset types here
      mrtNumber = 1
      mrtDetails = 2
      mrtContactDetails = &H100S 'In addition to mrtAll
        End Enum

    Public Enum FutureMembershipTypeErrors
      fmteNone
      fmteHistoricRenewalDate
      fmteMembershipType
    End Enum

    Public Enum SetCardIssueNumberTypes
      scintReinitialise
      scintIncrement
      scintDecrement
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum MemberFields
      mfAll = 0
      mfContactNumber
      mfMembershipType
      mfPaymentPlanNumber
      mfNumberOfMembers
      mfAgeOverride
      mfBranch
      mfJoined
      mfBranchMember
      mfApplied
      mfAccepted
      mfVotingRights
      mfMembershipCardExpires
      mfCancellationReason
      mfCancelledBy
      mfCancelledOn
      mfAmendedBy
      mfAmendedOn
      mfAddressNumber
      mfSource
      mfMembershipNumber
      mfMemberNumber
      mfReprintMshipCard
      mfCancellationSource
      mfMembershipCardIssueNumber
      mfMembershipStatus
      mfCmtDate
      mfLockBranch
    End Enum

    Private Enum MemberNumberTypes
      mntUnknown
      mntIsContactNumber
      mntSurnameJoinDate
      mntCharSeqInteger
      mntOrderNumber
      mntSequential
      mntMembershipNumber
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    Private mvMembershipType As MembershipType
    Private mvAddressDesc As String

    Private mvContact As Contact
    Private mvContactDesc As String
    Private mvDOBEstimated As Boolean
    Private mvDateOfBirth As String
    Private mvContactType As Contact.ContactTypes
    Private mvOwnershipGroup As String
    Private mvDistributionCode As String

    Private mvChangedDOBEstimated As Boolean
    Private mvChangedDateOfBirth As Boolean
    Private mvChangedOwnershipGroup As Boolean

    Private mvFutureMembershipTypeCode As String
    Private mvFutureMembershipType As MembershipType
    Private mvFutureChangeDate As String

    Private mvChangeMembershipType As Boolean
    Private mvCMTDate As String
    Private mvGiftMemberMaxJuniorAge As Integer

    'PayPlan Bits
    Private mvAutoPaymentMethod As Boolean
    Private mvRenewalDate As String
    Private mvRenewalPending As Boolean
    Private mvPayPlanTerm As Integer
    Private mvFixedRenewalCycle As Boolean
    Private mvAmendedValid As Boolean
    Private mvContinousRenewals As Boolean
    Private mvPayPlanTermUnits As PaymentPlan.OrderTermUnits
    Private mvPayPlanBalance As Double
    Private mvPayPlanBalanceSet As Boolean

    'Used for deciding which Smart Client menu items should be available
    Private mvPaymentPlan As PaymentPlan

    Private Const INVALID_NUMBER As Integer = -1

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub CancelOtherMembershipFlags()
      CancelOtherMembershipFlags(TodaysDate)
    End Sub
    Private Sub CancelOtherMembershipFlags(ByVal pValidToDate As String)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vTable As String

      If mvMembershipType.Activity.Length > 0 Then
        'End the Membership Activity
        Dim vCC As ContactCategory
        If mvContactType = Contact.ContactTypes.ctcOrganisation Then
          vCC = New OrganisationCategory(mvEnv)
          vWhereFields.Add("organisation_number", mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue)
        Else
          vCC = New ContactCategory(mvEnv)
          vWhereFields.Add("contact_number", mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue)
        End If
        vWhereFields.Add("activity", mvMembershipType.Activity)
        vWhereFields.Add("activity_value", mvMembershipType.ActivityValue)
        vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidToDate, CDBField.FieldWhereOperators.fwoGreaterThan)

        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vCC.FieldNames, vCC.AliasedTableName, vWhereFields)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          If mvContactType = Contact.ContactTypes.ctcOrganisation Then
            vCC = New OrganisationCategory(mvEnv)
          Else
            vCC = New ContactCategory(mvEnv)
          End If
          vCC.InitFromRecordSet(vRS)
          vCC.Update(vCC.ValidFrom, pValidToDate)
          If CDate(vCC.ValidTo).CompareTo(CDate(vCC.ValidFrom)) < 0 Then
            'ValidTo is before ValidFrom
            vCC.Update(vCC.ValidFrom, vCC.ValidFrom)
          End If
          If vCC.IsValidForUpdate Then
            vCC.Save(mvEnv.User.UserID, True)
          End If
        End While
        vRS.CloseRecordSet()
      End If
      If mvMembershipType.MailingSuppression.Length > 0 Then
        'End the Membership Suppression
        vWhereFields.Clear()
        If mvContactType = Contact.ContactTypes.ctcOrganisation Then
          vTable = "organisation_suppressions"
          vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue)
        Else
          vTable = "contact_suppressions"
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue)
        End If
        vWhereFields.Add("mailing_suppression", CDBField.FieldTypes.cftCharacter, mvMembershipType.MailingSuppression)
        vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidToDate, CDBField.FieldWhereOperators.fwoGreaterThan)
        vUpdateFields.Clear()
        vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, pValidToDate)
        mvEnv.Connection.UpdateRecords(vTable, vUpdateFields, vWhereFields, False)
      End If
      RemoveFutureRecord()
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) = True And mvChangeMembershipType = False Then
        'Do not do this for CMT as it needs to be done differently
        If mvMembershipType.UseMembershipGroups Then SetMembershipGroupsHistoric(mvEnv, MembershipNumber)
      End If
    End Sub

    Public Property ContactType() As Contact.ContactTypes
      Get
        ContactType = mvContactType
      End Get
      Set(ByVal Value As Contact.ContactTypes)
        mvContactType = Value
      End Set
    End Property

    Public WriteOnly Property PaymentPlanContinuousRenewals() As Boolean
      Set(ByVal Value As Boolean)
        mvContinousRenewals = Value
      End Set
    End Property

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property Accepted() As String
      Get
        Accepted = mvClassFields.Item(MemberFields.mfAccepted).Value
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        If Value <> mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue Then mvAddressDesc = ""
        mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue = Value
      End Set
    End Property

    Public Property AddressDesc() As String
      Get
        AddressDesc = mvAddressDesc
      End Get
      Set(ByVal Value As String)
        mvAddressDesc = Value
      End Set
    End Property

    Public Property AgeOverride() As String
      Get
        AgeOverride = mvClassFields.Item(MemberFields.mfAgeOverride).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfAgeOverride).Value = Value
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(MemberFields.mfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(MemberFields.mfAmendedOn).Value
      End Get
    End Property

    Public Property Applied() As String
      Get
        SetValid(MemberFields.mfApplied)
        Applied = mvClassFields.Item(MemberFields.mfApplied).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfApplied).Value = Value
      End Set
    End Property

    Public Property Branch() As String
      Get
        Branch = mvClassFields.Item(MemberFields.mfBranch).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfBranch).Value = Value
      End Set
    End Property

    Public Property BranchMember() As String
      Get
        BranchMember = mvClassFields.Item(MemberFields.mfBranchMember).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfBranchMember).Value = Value
      End Set
    End Property

    Public Property CancellationReason() As String
      Get
        CancellationReason = mvClassFields.Item(MemberFields.mfCancellationReason).Value
      End Get
      Set(ByVal Value As String)
        If Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) = 0 Then
          'Can only set Cancellation Reason if not already set
          mvClassFields.Item(MemberFields.mfCancellationReason).Value = Value
          mvClassFields.Item(MemberFields.mfCancelledBy).Value = mvEnv.User.UserID
          mvClassFields.Item(MemberFields.mfCancelledOn).Value = TodaysDate()
        End If
      End Set
    End Property

    Public ReadOnly Property CancellationSource() As String
      Get
        CancellationSource = mvClassFields.Item(MemberFields.mfCancellationSource).Value
      End Get
    End Property

    Public Property CancelledBy() As String
      Get
        CancelledBy = mvClassFields.Item(MemberFields.mfCancelledBy).Value
      End Get
      Set(ByVal Value As String)
        'Set by Let CancellationReason
        mvClassFields.Item(MemberFields.mfCancelledBy).Value = Value
      End Set
    End Property

    Public Property CancelledOn() As String
      Get
        CancelledOn = mvClassFields.Item(MemberFields.mfCancelledOn).Value
      End Get
      Set(ByVal Value As String)
        'Set by Let CancellationReason
        mvClassFields.Item(MemberFields.mfCancelledOn).Value = Value
      End Set
    End Property

    Public ReadOnly Property Contact() As Contact
      Get
        If mvContact Is Nothing Then
          mvContact = New Contact(mvEnv)
          mvContact.Init(CInt(mvClassFields.Item(MemberFields.mfContactNumber).Value))
        End If
        Contact = mvContact
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        If Value <> mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue Then mvContactDesc = ""
        mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public Property ContactDesc() As String
      Get
        ContactDesc = mvContactDesc
      End Get
      Set(ByVal Value As String)
        mvContactDesc = Value
      End Set
    End Property

    Public Property ContactDOBEstimated() As Boolean
      Get
        ContactDOBEstimated = mvDOBEstimated
      End Get
      Set(ByVal Value As Boolean)
        If mvDOBEstimated <> Value Then
          mvDOBEstimated = Value
          mvChangedDOBEstimated = True
        End If
      End Set
    End Property

    Public Property ContactDateOfBirth() As String
      Get
        ContactDateOfBirth = mvDateOfBirth
      End Get
      Set(ByVal Value As String)
        If IsDate(Value) Then
          If mvDateOfBirth <> Value Then
            mvDateOfBirth = Value
            mvChangedDateOfBirth = True
          End If
        End If
      End Set
    End Property

    Public WriteOnly Property DateOfBirth() As String
      Set(ByVal Value As String)
        If IsDate(Value) Then
          If mvDateOfBirth <> Value Then
            mvDateOfBirth = Value
          End If
        End If
      End Set
    End Property

    Public Property Joined() As String
      Get
        SetValid(MemberFields.mfJoined)
        Joined = mvClassFields.Item(MemberFields.mfJoined).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfJoined).Value = Value
      End Set
    End Property

    Public Property MemberNumber() As String
      Get
        If Len(mvClassFields.Item(MemberFields.mfMemberNumber).Value) = 0 Then
          If Len(mvClassFields.Item(MemberFields.mfContactNumber).Value) > 0 And Len(mvClassFields.Item(MemberFields.mfPaymentPlanNumber).Value) > 0 And Len(mvClassFields.Item(MemberFields.mfJoined).Value) > 0 Then
            'May need to use ContactNumber, PayPlanNumber and/or Joined Date
            mvClassFields.Item(MemberFields.mfMemberNumber).Value = GetMemberNumber("")
          End If
        End If
        MemberNumber = mvClassFields.Item(MemberFields.mfMemberNumber).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfMemberNumber).Value = Value
      End Set
    End Property

    Public Property MembershipCardExpires() As String
      Get
        MembershipCardExpires = mvClassFields.Item(MemberFields.mfMembershipCardExpires).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfMembershipCardExpires).Value = Value
      End Set
    End Property

    Public ReadOnly Property MembershipNumber() As Integer
      Get
        MembershipNumber = mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property MembershipType() As MembershipType
      Get
        If mvMembershipType Is Nothing Then
          If MembershipTypeCode.Length > 0 Then
            Dim vMTCode As String = MembershipTypeCode
            mvMembershipType = mvEnv.MembershipType(vMTCode)
          End If
        End If
        MembershipType = mvMembershipType
      End Get
    End Property

    Public ReadOnly Property MembershipTypeCode() As String
      Get
        MembershipTypeCode = mvClassFields.Item(MemberFields.mfMembershipType).Value
      End Get
    End Property

    Public Property NumberOfMembers() As Integer
      Get
        NumberOfMembers = mvClassFields.Item(MemberFields.mfNumberOfMembers).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(MemberFields.mfNumberOfMembers).IntegerValue = Value
      End Set
    End Property

    Public WriteOnly Property PaymentPlanAutoPayMethod() As Boolean
      Set(ByVal Value As Boolean)
        mvAutoPaymentMethod = Value
      End Set
    End Property

    Public Property PaymentPlanNumber() As Integer
      Get
        PaymentPlanNumber = mvClassFields.Item(MemberFields.mfPaymentPlanNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(MemberFields.mfPaymentPlanNumber).IntegerValue = Value
      End Set
    End Property

    Public Property PaymentPlanRenewalDate() As String
      Get
        Return mvRenewalDate
      End Get
      Set(ByVal Value As String)
        mvRenewalDate = Value
      End Set
    End Property

    Public WriteOnly Property PaymentPlanRenewalPending() As Boolean
      Set(ByVal Value As Boolean)
        mvRenewalPending = Value
      End Set
    End Property

    Public WriteOnly Property PaymentPlanTerm() As Integer
      Set(ByVal Value As Integer)
        mvPayPlanTerm = Value
      End Set
    End Property

    Public Property ReprintMshipCard() As Boolean
      Get
        ReprintMshipCard = mvClassFields.Item(MemberFields.mfReprintMshipCard).Bool
      End Get
      Set(ByVal Value As Boolean)
        If Value = False Then
          mvClassFields.Item(MemberFields.mfReprintMshipCard).Value = ""
        Else
          mvClassFields.Item(MemberFields.mfReprintMshipCard).Bool = Value
        End If
      End Set
    End Property
    Public Property Source() As String
      Get
        Source = mvClassFields.Item(MemberFields.mfSource).Value
      End Get
      Set(ByVal Value As String)
        mvClassFields.Item(MemberFields.mfSource).Value = Value
      End Set
    End Property

    Public Property VotingRights() As Boolean
      Get
        VotingRights = mvClassFields.Item(MemberFields.mfVotingRights).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(MemberFields.mfVotingRights).Bool = Value
      End Set
    End Property

    Public Property FutureChangeDate() As String
      Get
        FutureChangeDate = mvFutureChangeDate
      End Get
      Set(ByVal Value As String)
        mvFutureChangeDate = Value
      End Set
    End Property

    Public Property FutureMembershipTypeCode() As String
      Get
        FutureMembershipTypeCode = mvFutureMembershipTypeCode
      End Get
      Set(ByVal Value As String)
        mvFutureMembershipTypeCode = Value
        mvFutureMembershipType = mvEnv.MembershipType(Value)
      End Set
    End Property

    Friend ReadOnly Property FutureMembershipType() As MembershipType
      Get
        FutureMembershipType = mvFutureMembershipType
      End Get
    End Property

    Public WriteOnly Property PaymentPlantermUnits() As PaymentPlan.OrderTermUnits
      Set(ByVal Value As PaymentPlan.OrderTermUnits)
        mvPayPlanTermUnits = Value
      End Set
    End Property

    Public WriteOnly Property PaymentPlanBalance() As Double
      Set(ByVal Value As Double)
        mvPayPlanBalance = Value
        mvPayPlanBalanceSet = True
      End Set
    End Property
    Public Property PaymentPlanFixedRenewalCycle() As Boolean
      Get
        Return mvFixedRenewalCycle
      End Get
      Set(ByVal Value As Boolean)
        mvFixedRenewalCycle = Value
      End Set
    End Property

    Public ReadOnly Property GiftMemberMaxJuniorAge() As Integer
      Get
        GiftMemberMaxJuniorAge = mvGiftMemberMaxJuniorAge
      End Get
    End Property

    Public ReadOnly Property MembershipCardIssueNumber() As Integer
      Get
        If mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).IntegerValue = 0 Then mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
        MembershipCardIssueNumber = mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property MembershipStatus() As String
      Get
        MembershipStatus = mvClassFields.Item(MemberFields.mfMembershipStatus).Value
      End Get
    End Property
    ''' <summary>Following a change of membership type (CMT), this is the date the new membership is effective from.</summary>
    Public ReadOnly Property CMTDate() As String
      Get
        Return mvClassFields.Item(MemberFields.mfCmtDate).Value
      End Get
    End Property
    Public ReadOnly Property LockBranch() As Boolean
      Get
        Return mvClassFields.Item(MemberFields.mfLockBranch).Bool
      End Get
    End Property

    Public Property ContactOwnershipGroup() As String
      Get
        ContactOwnershipGroup = mvOwnershipGroup
      End Get
      Set(ByVal Value As String)
        If mvOwnershipGroup <> Value Then
          mvOwnershipGroup = Value
          mvChangedOwnershipGroup = True
        End If
      End Set
    End Property

    Public WriteOnly Property LineValue(ByVal pAttributeName As String) As String
      Set(ByVal Value As String)
        'Used by Smart Client to set members from MembershipMembersSummary grid
        Select Case pAttributeName
          Case "DateOfBirth"
            mvDateOfBirth = Value
          Case "DistributionCode"
            mvDistributionCode = Value
          Case "DOBEstimated"
            mvDOBEstimated = BooleanValue(Value)
          Case "MembershipType"
            mvClassFields.Item(MemberFields.mfMembershipType).Value = Value
            If Len(mvClassFields.Item(MemberFields.mfMembershipType).Value) > 0 Then mvMembershipType = mvEnv.MembershipType((mvClassFields.Item(MemberFields.mfMembershipType).Value))
          Case "ContactName", "AddressLine", "LineNumber"
            'Do nothing
          Case Else
            mvClassFields.ItemValue(pAttributeName) = Value
        End Select

        '<MemberLine AddressNumber="252893" MembershipNumber="4607" DOBEstimated="N" ContactNumber="888903205" MembershipType="ANG2"
        'ContactName="Mr R Trent" Joined="01/09/2006" Branch="HN" BranchMember="N" Applied="01/09/2006" AddressLine="1 Church Walk, BLACKBURN, BB1 9QS" LineNumber="1" />

        If mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue > 0 And mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue > 0 Then
          If mvContact Is Nothing Then
            mvContact = New Contact(mvEnv)
            mvContact.Init((mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue), (mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue))
          End If
        End If
      End Set
    End Property

    Friend ReadOnly Property DistributionCode() As String
      Get
        DistributionCode = mvDistributionCode
      End Get
    End Property

    Friend ReadOnly Property OriginalMembershipTypeCode() As String
      Get
        Dim vValue As String = mvClassFields.Item(MemberFields.mfMembershipType).SetValue
        If mvExisting = False OrElse vValue.Length = 0 Then vValue = mvClassFields.Item(MemberFields.mfMembershipType).Value
        Return vValue
      End Get
    End Property

    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        'There should be an entry here for each field in the table
        'Keep these in the same order as the Fields enum
        With mvClassFields
          .DatabaseTableName = "members"
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("membership_type")
          .Add("order_number", CDBField.FieldTypes.cftLong)
          .Add("number_of_members", CDBField.FieldTypes.cftInteger)
          .Add("age_override", CDBField.FieldTypes.cftInteger)
          .Add("branch")
          .Add("joined", CDBField.FieldTypes.cftDate)
          .Add("branch_member")
          .Add("applied", CDBField.FieldTypes.cftDate)
          .Add("accepted", CDBField.FieldTypes.cftDate)
          .Add("voting_rights")
          .Add("membership_card_expires", CDBField.FieldTypes.cftDate)
          .Add("cancellation_reason")
          .Add("cancelled_by")
          .Add("cancelled_on", CDBField.FieldTypes.cftDate)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("address_number", CDBField.FieldTypes.cftLong)
          .Add("source")
          .Add("membership_number", CDBField.FieldTypes.cftLong)
          .Add("member_number")
          .Add("reprint_mship_card")
          .Add("cancellation_source")
          .Add("membership_card_issue_number", CDBField.FieldTypes.cftInteger)
          .Add("membership_status")
          .Add("cmt_date", CDBField.FieldTypes.cftDate)
          .Add("lock_branch")
        End With
        mvClassFields.Item(MemberFields.mfMembershipNumber).SetPrimaryKeyOnly()

        mvClassFields.Item(MemberFields.mfMembershipNumber).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfContactNumber).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfMembershipType).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfPaymentPlanNumber).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfNumberOfMembers).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfBranch).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfVotingRights).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfCancellationReason).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfCancelledBy).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfCancelledOn).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfAmendedBy).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfAmendedOn).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfSource).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfAddressNumber).PrefixRequired = True
        mvClassFields.Item(MemberFields.mfCancellationSource).PrefixRequired = True
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipStatus) Then
          mvClassFields.Item(MemberFields.mfMembershipStatus).PrefixRequired = True
        End If
        mvClassFields.Item(MemberFields.mfCmtDate).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdvanceCMT)
        mvChangedDOBEstimated = False
        mvChangedDateOfBirth = False
        mvGiftMemberMaxJuniorAge = IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGiftMemberMaxJuniorAge))
        If mvGiftMemberMaxJuniorAge = 0 Then mvGiftMemberMaxJuniorAge = 16 'Default

        mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipCardIssueNumber)
        mvClassFields.Item(MemberFields.mfLockBranch).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbLockBranch)
      Else
        mvClassFields.ClearItems()
      End If
      mvAmendedValid = False
      mvExisting = False
      mvDistributionCode = ""
      mvCMTDate = ""
    End Sub
    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
      mvMembershipType = Nothing
      mvMembershipType = New MembershipType(mvEnv)
      mvMembershipType.Init()

      mvClassFields.Item(MemberFields.mfContactNumber).Value = CStr(INVALID_NUMBER)
      mvClassFields.Item(MemberFields.mfMembershipType).Value = ""
      mvClassFields.Item(MemberFields.mfPaymentPlanNumber).Value = CStr(INVALID_NUMBER)
      mvClassFields.Item(MemberFields.mfNumberOfMembers).Value = CStr(1)
      mvClassFields.Item(MemberFields.mfAgeOverride).Value = ""
      mvClassFields.Item(MemberFields.mfBranch).Value = ""
      mvClassFields.Item(MemberFields.mfJoined).Value = ""
      mvClassFields.Item(MemberFields.mfBranchMember).Value = "N"
      mvClassFields.Item(MemberFields.mfApplied).Value = ""
      mvClassFields.Item(MemberFields.mfAccepted).Value = ""
      mvClassFields.Item(MemberFields.mfVotingRights).Value = "N"
      mvClassFields.Item(MemberFields.mfMembershipCardExpires).Value = ""
      mvClassFields.Item(MemberFields.mfCancellationReason).Value = ""
      mvClassFields.Item(MemberFields.mfCancelledBy).Value = ""
      mvClassFields.Item(MemberFields.mfCancelledOn).Value = ""
      mvClassFields.Item(MemberFields.mfAmendedBy).Value = ""
      mvClassFields.Item(MemberFields.mfAmendedOn).Value = ""
      mvClassFields.Item(MemberFields.mfAddressNumber).Value = CStr(INVALID_NUMBER)
      mvClassFields.Item(MemberFields.mfSource).Value = ""
      mvClassFields.Item(MemberFields.mfMembershipNumber).Value = CStr(INVALID_NUMBER)
      mvClassFields.Item(MemberFields.mfMemberNumber).Value = ""
      mvClassFields.Item(MemberFields.mfReprintMshipCard).Value = "N"
      mvClassFields.Item(MemberFields.mfLockBranch).Value = "N"

      mvContactDesc = ""
      mvAddressDesc = ""
      mvDOBEstimated = False
      mvDateOfBirth = ""

      mvAutoPaymentMethod = False
      mvRenewalDate = ""
      mvRenewalPending = True
      mvPayPlanTerm = 0
      mvPayPlanBalance = 0
      mvPayPlanBalanceSet = False

      mvExisting = False
    End Sub

    Private Sub SetValid(ByRef pField As MemberFields)
      'Add code here to ensure all values are valid before saving
      '  If pField = mfall Then
      '    mvClassFields.Item(mfAmendedOn).Value = TodaysDate()
      '    mvClassFields.Item(mfAmendedBy).Value = mvEnv.User.Logname
      '  End If
      If pField = MemberFields.mfAll And Not mvAmendedValid Then
        mvClassFields.Item(MemberFields.mfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(MemberFields.mfAmendedBy).Value = mvEnv.User.UserID
      End If
      If (pField = MemberFields.mfAll Or pField = MemberFields.mfJoined) And Len(mvClassFields.Item(MemberFields.mfJoined).Value) = 0 Then
        SetJoinedDate()
      End If
      If (pField = MemberFields.mfAll Or pField = MemberFields.mfApplied) And Len(mvClassFields.Item(MemberFields.mfApplied).Value) = 0 Then
        mvClassFields.Item(MemberFields.mfApplied).Value = mvClassFields.Item(MemberFields.mfJoined).Value
      End If
      If (pField = MemberFields.mfAll Or pField = MemberFields.mfMemberNumber) Then
        If Len(mvClassFields.Item(MemberFields.mfMemberNumber).Value) = 0 Then
          mvClassFields.Item(MemberFields.mfMemberNumber).Value = GetMemberNumber("")
        ElseIf mvEnv.GetConfigOption("enter_member_number") Then
          mvClassFields.Item(MemberFields.mfMemberNumber).Value = GetMemberNumber(MemberNumber)
        End If
      End If
      If (pField = MemberFields.mfAll Or pField = MemberFields.mfMembershipCardIssueNumber) And mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).IntegerValue = 0 Then
        mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
      End If
    End Sub

    Private Function GetMemberNumber(ByVal pMemberNumber As String) As String
      Dim vFormat As String
      Dim vType As MemberNumberTypes
      Dim vDone As Boolean
      Dim vMemberNumber As String

      vFormat = mvEnv.GetConfig("member_number_format")
      If vFormat = "varies" Then
        'See if there is an existing member number
        If mvExisting Then
          If mvClassFields.Item(MemberFields.mfMemberNumber).Value.Length > 0 Then
            If mvEnv.GetConfig("me_exist_member_number_format") = "current_member_number" Then 'GEOLSOC
              Return mvClassFields.Item(MemberFields.mfMemberNumber).Value
            Else
              RaiseError(DataAccessErrors.daeCannotGetMemberNumber, XLAT("member_number_format incorrectly configured"))
              Return ""     'Fix compiler warning
            End If
            vDone = True
          End If
        End If
        If Not vDone Then
          vMemberNumber = mvEnv.Connection.GetValue("SELECT member_number FROM members WHERE contact_number = " & ContactNumber)
          If Len(vMemberNumber) > 0 Then
            Return vMemberNumber
            vDone = True
          Else
            vType = GetMemberNumberType(mvEnv.GetConfig("me_new_member_number_format"))
          End If
        End If
      Else
        vType = GetMemberNumberType(vFormat)
      End If
      If Not vDone Then
        If vType = MemberNumberTypes.mntUnknown Then
          RaiseError(DataAccessErrors.daeCannotGetMemberNumber, XLAT("member_number_format incorrectly configured"))
          Return ""     'Fix compiler warning
        Else
          Return GetMemberNumberByType(vType, pMemberNumber)
        End If
      End If
      Return ""     'Fix compiler warning
    End Function

    Private Function GetMemberNumberType(ByVal pConfig As String) As MemberNumberTypes
      Select Case pConfig
        Case "contact_number" 'RSPCA
          GetMemberNumberType = MemberNumberTypes.mntIsContactNumber
        Case "contact_no_surname_jndate"
          GetMemberNumberType = MemberNumberTypes.mntSurnameJoinDate
        Case "char_seq_integer"
          GetMemberNumberType = MemberNumberTypes.mntCharSeqInteger
        Case "order_number" 'RSPB
          GetMemberNumberType = MemberNumberTypes.mntOrderNumber
        Case "sequential"
          GetMemberNumberType = MemberNumberTypes.mntSequential
        Case ""
          GetMemberNumberType = MemberNumberTypes.mntMembershipNumber 'Default
        Case Else
          GetMemberNumberType = MemberNumberTypes.mntUnknown
      End Select
    End Function
    Private Function GetMemberNumberByType(ByVal pType As MemberNumberTypes, ByVal pMemberNumber As String) As String
      Dim vContactNumber As String
      Dim vLoMem As String
      Dim vHiMem As String
      Dim vIntMem As Integer
      Dim vCharMem As String = ""
      Dim vRecordSet As CDBRecordSet
      Dim vCount As Integer
      Dim vPos As Integer
      Dim vErrorMsg As String = ""
      Dim vSurname As String

      Select Case pType
        Case MemberNumberTypes.mntIsContactNumber 'RSPCA
          Return mvClassFields.Item(MemberFields.mfContactNumber).Value

        Case MemberNumberTypes.mntSurnameJoinDate 'IFST
          vContactNumber = Left(mvClassFields.Item(MemberFields.mfContactNumber).Value & "00000", 5)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT surname FROM contacts WHERE contact_number = " & mvClassFields.Item(MemberFields.mfContactNumber).Value)
          If vRecordSet.Fetch() = True Then
            vSurname = vRecordSet.Fields.Item("surname").Value
            Return Left(vSurname, 1) & Right(CStr(Year(CDate(mvClassFields.Item(MemberFields.mfJoined).Value))), 2) & vContactNumber
          Else
            vErrorMsg = XLAT("Failed to find contact number") 'LoadString(IDBTRADER + 2)
          End If
          vRecordSet.CloseRecordSet()

        Case MemberNumberTypes.mntCharSeqInteger 'NCF SEMA
          If Len(pMemberNumber) = 8 Then
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT cancellation_reason FROM members WHERE member_number = '" & pMemberNumber & "'")
            If vRecordSet.Fetch() = True Then
              If vRecordSet.Fields.Item("cancellation_reason").Value.Length > 0 Then
                Return pMemberNumber
              Else
                vErrorMsg = XLAT("Member Number already exists") 'LoadString(IDBTRADER + 3)
              End If
            Else
              Return pMemberNumber
            End If
            vRecordSet.CloseRecordSet()
          Else
            If Len(pMemberNumber) = 0 Then
              pMemberNumber = mvEnv.GetConfig("me_member_number_prefix")
            End If
            vLoMem = Left(pMemberNumber & "0000000", 8)
            vHiMem = Left(pMemberNumber & "9999999", 8)
            vCount = mvEnv.Connection.GetCount("members", Nothing, "member_number BETWEEN '" & vLoMem & "' AND '" & vHiMem & "'")
            If vCount = 0 Then
              GetMemberNumberByType = Left(pMemberNumber & "0000000", 7) & "1"
            Else
              vRecordSet = mvEnv.Connection.GetRecordSet("SELECT MAX(member_number)  AS  max_number FROM members WHERE member_number BETWEEN '" & vLoMem & "' AND '" & vHiMem & "'")
              If vRecordSet.Fetch() = True Then
                vCharMem = vRecordSet.Fields.Item("max_number").Value
                vPos = 1
                While Not Mid(vCharMem, vPos, 1) Like "#"
                  vPos = vPos + 1
                End While
                vIntMem = IntegerValue(Mid(vCharMem, vPos))
                vCharMem = Left(vCharMem, vPos - 1)
              Else
                vIntMem = 0
              End If
              vRecordSet.CloseRecordSet()
              vIntMem = vIntMem + 1
              Return vCharMem & Right("00000000" & vIntMem, 8 - Len(vCharMem))
            End If
          End If

        Case MemberNumberTypes.mntOrderNumber 'RSPB
          Return mvClassFields.Item(MemberFields.mfPaymentPlanNumber).Value

        Case MemberNumberTypes.mntSequential
          vErrorMsg = "The 'sequential' option used in the member_number_format or me_new_member_number_format config is not supported by this software"
          vErrorMsg = vErrorMsg & vbCrLf & "The Member Number control number (M) should be adjusted to the higher of member_number and membership_number and the config removed"
          vErrorMsg = vErrorMsg & vbCrLf & "This will result in new members having member_number and membership_number set to the same number"
#If OLD_SEQUENTIAL_MEMBER_NUMBER_CODE Then
				'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression OLD_SEQUENTIAL_MEMBER_NUMBER_CODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
				'This code will not work well (if at all) in the GUI as (in empress) it will lock all records in the members table
				Set vRecordSet = mvEnv.Connection.GetRecordSet("SELECT MAX(" & mvEnv.Connection.DBLength("member_number") & ")  AS  max_number FROM members")
				If vRecordSet.Fetch = rssOK Then
				vLen = Val(vRecordSet.Fields.Item("max_number").Value)
				vRecordSet.CloseRecordSet
				Set vRecordSet = mvEnv.Connection.GetRecordSet("SELECT member_number FROM members WHERE " & mvEnv.Connection.DBLength("member_number") & " = " & vLen & " ORDER BY member_number DESC")
				If vRecordSet.Fetch = rssOK Then
				vIntMem = Val(vRecordSet.Fields.Item("member_number").Value)
				vIntMem = vIntMem + 1
				vCharMem = Format$(vIntMem)
				If Len(vCharMem) < vLen Then vCharMem = Right$("00000000" & vCharMem, vLen)
				return vCharMem
				Else
				vErrorMsg = XLAT("Failed to retrieve member number record")      'LoadString(IDBTRADER + 5)
				End If
				Else
				vErrorMsg = XLAT("Failed to retrieve member number record")      'LoadString(IDBTRADER + 6)
				End If
				vRecordSet.CloseRecordSet
#End If
        Case Else 'mntMembershipNumber
          'Will be null and set to membership_number
      End Select
      If vErrorMsg.Length > 0 Then RaiseError(DataAccessErrors.daeCannotGetMemberNumber, vErrorMsg)
      Return ""       'Fix compiler warning message
    End Function

    Private Sub RemoveFutureRecord()
      Dim vMemberFutureType As New FutureMembershipType(mvEnv)

      Dim vActivity As String = String.Empty
      Dim vActivityValue As String = String.Empty

      Dim vAnsiJoins As New AnsiJoins({New AnsiJoin("membership_types mt", "fmt.future_membership_type", "mt.membership_type")})
      Dim vWhereFields As New CDBFields(New CDBField("fmt.membership_number", MembershipNumber))
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vMemberFutureType.FieldNames & ", mt.activity, mt.activity_value", vMemberFutureType.AliasedTableName, vWhereFields, "", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      If vRS.Fetch Then
        vMemberFutureType.InitFromRecordSet(vRS)
        vActivity = vRS.Fields("activity").Value
        vActivityValue = vRS.Fields("activity_value").Value
      End If
      vRS.CloseRecordSet()

      Dim vTrans As Boolean = False
      If String.IsNullOrWhiteSpace(vActivity) = False Then
        'Delete Trigger Activity
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", ContactNumber)
        vWhereFields.Add("activity", vActivity)
        vWhereFields.Add("activity_value", vActivityValue)
        vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vMemberFutureType.FutureChangeDate)
        vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, vMemberFutureType.FutureChangeDate)

        Dim vCC As ContactCategory
        If mvContactType = Contact.ContactTypes.ctcOrganisation Then
          vCC = New OrganisationCategory(mvEnv)
        Else
          vCC = New ContactCategory(mvEnv)
        End If
        vSQLStatement = New SQLStatement(mvEnv.Connection, vCC.FieldNames, vCC.AliasedTableName, vWhereFields)
        vRS = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          If vTrans = False AndAlso mvEnv.Connection.InTransaction = False Then
            vTrans = mvEnv.Connection.StartTransaction()
          End If
          If mvContactType = Contact.ContactTypes.ctcOrganisation Then
            vCC = New OrganisationCategory(mvEnv)
          Else
            vCC = New ContactCategory(mvEnv)
          End If
          vCC.InitFromRecordSet(vRS)
          vCC.Delete(mvEnv.User.UserID, True)
        End While
        vRS.CloseRecordSet()
      End If

      If vMemberFutureType IsNot Nothing AndAlso vMemberFutureType.Existing Then
        If vTrans = False AndAlso mvEnv.Connection.InTransaction = False Then
          vTrans = mvEnv.Connection.StartTransaction()
        End If
        vMemberFutureType.Delete(mvEnv.User.UserID)
      End If

      If vTrans Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As MemberRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And MemberRecordSetTypes.mrtAll) = MemberRecordSetTypes.mrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "m")
      Else
        If (pRSType And MemberRecordSetTypes.mrtNumber) > 0 Then vFields = "membership_number,member_number,m.contact_number,m.address_number,m.order_number,membership_type"
        If (pRSType And MemberRecordSetTypes.mrtDetails) > 0 Then vFields = vFields & ",m.source"
      End If
      If (pRSType And MemberRecordSetTypes.mrtContactDetails) > 0 Then vFields = vFields & ",dob_estimated,date_of_birth,label_name,contact_type"
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pMembershipNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      If pMembershipNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(MemberRecordSetTypes.mrtAll Or MemberRecordSetTypes.mrtContactDetails) & " FROM members m, contacts c WHERE membership_number = " & pMembershipNumber & " AND m.contact_number = c.contact_number")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, MemberRecordSetTypes.mrtAll)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As MemberRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Modify below to handle each recordset type as required
        If (pRSType And MemberRecordSetTypes.mrtNumber) > 0 Then
          .SetItem(MemberFields.mfContactNumber, vFields)
          .SetItem(MemberFields.mfMembershipType, vFields)
          .SetItem(MemberFields.mfPaymentPlanNumber, vFields)
          .SetItem(MemberFields.mfAddressNumber, vFields)
          .SetItem(MemberFields.mfMembershipNumber, vFields)
          .SetItem(MemberFields.mfMemberNumber, vFields)
        End If
        If (pRSType And MemberRecordSetTypes.mrtDetails) > 0 Then
          .SetItem(MemberFields.mfSource, vFields)
        End If
        If (pRSType And MemberRecordSetTypes.mrtContactDetails) > 0 Then
          mvContactDesc = vFields.Item("label_name").Value 'ANDY this really should be a Contact Class
          mvDOBEstimated = vFields.Item("dob_estimated").Bool
          mvDateOfBirth = vFields.Item("date_of_birth").Value
          'Select Case vFields.Item("contact_type").Value
          '  Case "O"
          '    mvContactType = Contact.ContactTypes.ctcOrganisation
          '  Case "J"
          '    mvContactType = Contact.ContactTypes.ctcJoint
          '  Case Else
          '    mvContactType = Contact.ContactTypes.ctcContact
          'End Select
        End If
        If (pRSType And MemberRecordSetTypes.mrtAll) = MemberRecordSetTypes.mrtAll Then
          .SetItem(MemberFields.mfNumberOfMembers, vFields)
          .SetItem(MemberFields.mfAgeOverride, vFields)
          .SetItem(MemberFields.mfBranch, vFields)
          .SetItem(MemberFields.mfJoined, vFields)
          .SetItem(MemberFields.mfBranchMember, vFields)
          .SetItem(MemberFields.mfApplied, vFields)
          .SetItem(MemberFields.mfAccepted, vFields)
          .SetItem(MemberFields.mfVotingRights, vFields)
          .SetItem(MemberFields.mfMembershipCardExpires, vFields)
          .SetItem(MemberFields.mfCancellationReason, vFields)
          .SetItem(MemberFields.mfCancelledBy, vFields)
          .SetItem(MemberFields.mfCancelledOn, vFields)
          .SetItem(MemberFields.mfAmendedBy, vFields)
          .SetItem(MemberFields.mfAmendedOn, vFields)
          .SetItem(MemberFields.mfReprintMshipCard, vFields)
          .SetOptionalItem(MemberFields.mfCancellationSource, vFields)
          .SetOptionalItem(MemberFields.mfMembershipCardIssueNumber, vFields)
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMembershipStatus) Then
            .SetOptionalItem(MemberFields.mfMembershipStatus, vFields)
          End If
          .SetOptionalItem(MemberFields.mfLockBranch, vFields)
          mvMembershipType = mvEnv.MembershipType((.Item(MemberFields.mfMembershipType).Value))

          'BR17976 mvContactType now set for MemberRecordSetTypes.mrtAll.  
          'CancelMemberData and CancelCategories now using correct code path for organisations.  Now organisation member category cancellation date set correctly.
          If vFields.ContainsKey("contact_type") OrElse ((pRSType And MemberRecordSetTypes.mrtContactDetails) > 0) Then
            Select Case vFields.Item("contact_type").Value
              Case "O"
                mvContactType = Contact.ContactTypes.ctcOrganisation
              Case "J"
                mvContactType = Contact.ContactTypes.ctcJoint
              Case Else
                mvContactType = Contact.ContactTypes.ctcContact
            End Select
          End If

          If mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).InDatabase And Len(mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value) = 0 Then mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
        End If
      End With
    End Sub
    Public Sub AddCancelledActivityAndSuppression(ByRef pValidFrom As String)
      Dim vToDate As String

      vToDate = mvClassFields(MemberFields.mfCancelledOn).Value
      If mvMembershipType.Activity.Length > 0 Then
        If vToDate.Length > 0 Then
          Dim vCC As New ContactCategory(mvEnv)
          vCC.ContactTypeSaveActivity(mvContactType, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvMembershipType.Activity, mvMembershipType.ActivityValue, mvClassFields.Item(MemberFields.mfSource).Value, pValidFrom, vToDate, "", ContactCategory.ActivityEntryStyles.aesNormal, "", mvClassFields.Item(MemberFields.mfAmendedOn).Value, mvClassFields.Item(MemberFields.mfAmendedBy).Value)
        End If
      End If
      If Len(mvMembershipType.MailingSuppression) > 0 And Len(vToDate) > 0 Then
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesNormal, mvContactType, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvMembershipType.MailingSuppression, pValidFrom, vToDate, mvClassFields.Item(MemberFields.mfAmendedOn).Value, mvClassFields.Item(MemberFields.mfAmendedBy).Value)
      End If
    End Sub

    Public Sub AddActivityAndSuppression(ByRef pValidFrom As String)
      Dim vToDate As String = ""

      If mvMembershipType.Activity.Length > 0 Then
        If mvContinousRenewals Then
          vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvMembershipType.SuspensionGrace, CDate(mvRenewalDate)))
        Else
          If mvAutoPaymentMethod Then
            vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, mvMembershipType.SuspensionGrace, CDate(mvRenewalDate)))
            If mvMembershipType.SuspensionGrace > 0 Then vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, -1, CDate(vToDate)))
          Else
            If mvPayPlanBalanceSet = True And mvPayPlanBalance = 0 Then
              'New Payment Plan balance = 0 so RenewalDate rolled forward already
              'Keep vToDate as RenewalDate
              vToDate = mvRenewalDate
            Else
              If mvMembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfMonthlyTerm Then
                vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, CDbl(mvMembershipType.MembershipTerm), CDate(mvRenewalDate)))
              ElseIf mvMembershipType.PaymentTerm = MembershipType.MembershipTypeTerms.mtfWeeklyTerm Then
                vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.WeekOfYear, CDbl(mvMembershipType.MembershipTerm), CDate(mvRenewalDate)))
              Else
                vToDate = CDate(mvRenewalDate).AddYears(mvMembershipType.MembershipTerm).ToString(CAREDateFormat)
              End If
            End If
          End If
          If mvChangeMembershipType Then
            'Valid To is to be the renewal date - 1 day but only if payment already received for this year
            If CDate(mvRenewalDate) > CDate(TodaysDate()) Then
              vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(mvRenewalDate)))
            Else
              vToDate = ""
            End If
          End If
        End If
        If vToDate.Length > 0 Then
          Dim vCC As New ContactCategory(mvEnv)
          vCC.ContactTypeSaveActivity(mvContactType, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvMembershipType.Activity, mvMembershipType.ActivityValue, mvClassFields.Item(MemberFields.mfSource).Value, pValidFrom, vToDate, "", ContactCategory.ActivityEntryStyles.aesNormal, "", mvClassFields.Item(MemberFields.mfAmendedOn).Value, mvClassFields.Item(MemberFields.mfAmendedBy).Value)
        End If
      End If
      If Len(mvMembershipType.MailingSuppression) > 0 And Len(vToDate) > 0 Then
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesNormal, mvContactType, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvMembershipType.MailingSuppression, pValidFrom, vToDate, mvClassFields.Item(MemberFields.mfAmendedOn).Value, mvClassFields.Item(MemberFields.mfAmendedBy).Value)
      End If
    End Sub

    Public Sub InitFutureMember(ByRef pCurrentMember As Member)
      With pCurrentMember
        mvClassFields.Item(MemberFields.mfContactNumber).Value = CStr(.ContactNumber)
        mvClassFields.Item(MemberFields.mfAddressNumber).Value = CStr(.AddressNumber)
        mvClassFields.Item(MemberFields.mfMembershipType).Value = .FutureMembershipTypeCode
        mvMembershipType = .FutureMembershipType
        mvClassFields.Item(MemberFields.mfSource).Value = .Source
        mvClassFields.Item(MemberFields.mfPaymentPlanNumber).Value = CStr(.PaymentPlanNumber)
        mvClassFields.Item(MemberFields.mfNumberOfMembers).Value = CStr(.NumberOfMembers)
        mvClassFields.Item(MemberFields.mfAgeOverride).Value = .AgeOverride
        mvClassFields.Item(MemberFields.mfBranch).Value = .Branch
        mvClassFields.Item(MemberFields.mfJoined).Value = .Joined
        mvClassFields.Item(MemberFields.mfBranchMember).Value = .BranchMember
        mvClassFields.Item(MemberFields.mfApplied).Value = .Applied
        mvClassFields.Item(MemberFields.mfAccepted).Value = .Accepted
        mvClassFields.Item(MemberFields.mfVotingRights).Bool = .VotingRights
        mvClassFields.Item(MemberFields.mfMembershipCardExpires).Value = .MembershipCardExpires
        mvClassFields.Item(MemberFields.mfMemberNumber).Value = .MemberNumber
        mvClassFields.Item(MemberFields.mfMembershipNumber).Value = CStr(mvEnv.GetControlNumber("M"))
        mvClassFields.Item(MemberFields.mfAmendedOn).Value = TodaysDate()
        mvClassFields.Item(MemberFields.mfAmendedBy).Value = "automatic"
        mvContactType = pCurrentMember.ContactType
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False, Optional ByRef pBatchNumber As Integer = 0, Optional ByRef pTransactionNumber As Integer = 0, Optional ByRef pWarningMessage As String = "")
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vDate As String
      Dim vCancellation As Boolean
      Dim vTransaction As Boolean

      SetValid(MemberFields.mfAll)
      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      'UPDATE CONTACT DOB & DOBESTIMATED
      If mvChangedDOBEstimated Then vUpdateFields.Add("dob_estimated", CDBField.FieldTypes.cftCharacter, BooleanString(mvDOBEstimated))
      If mvChangedDateOfBirth Then vUpdateFields.Add("date_of_birth", CDBField.FieldTypes.cftDate, mvDateOfBirth)
      If mvChangedOwnershipGroup Then vUpdateFields.Add("ownership_group", CDBField.FieldTypes.cftCharacter, mvOwnershipGroup)
      If mvChangedDOBEstimated Or mvChangedDateOfBirth Or mvChangedOwnershipGroup Then
        vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(MemberFields.mfContactNumber).Value)
        mvEnv.Connection.UpdateRecords("contacts", vUpdateFields, vWhereFields)
        If mvChangedOwnershipGroup And mvContactType = Contact.ContactTypes.ctcOrganisation Then
          If mvChangedDOBEstimated Then vUpdateFields.Remove("dob_estimated")
          If mvChangedDateOfBirth Then vUpdateFields.Remove("date_of_birth")
          vWhereFields.Item("contact_number").Name = "organisation_number"
          mvEnv.Connection.UpdateRecords("organisations", vUpdateFields, vWhereFields)
        End If
      End If
      'UPDATE/INSERT MEMBER
      If mvExisting Then
        If mvClassFields.Item(MemberFields.mfCancellationReason).SetValue <> mvClassFields.Item(MemberFields.mfCancellationReason).Value Then vCancellation = True
        mvClassFields.Save(mvEnv, mvExisting, mvEnv.User.UserID, mvEnv.AuditStyle = CDBEnvironment.AuditStyleTypes.ausExtended)
        mvEnv.AddJournalRecord(JournalTypes.jnlMember, CType(IIf(vCancellation, JournalOperations.jnlCancel, JournalOperations.jnlUpdate), JournalOperations), mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue, (mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)
        If vCancellation Then CancelOtherMembershipFlags()
      Else
        If Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) > 0 Then vCancellation = True
        If CDbl(mvClassFields.Item(MemberFields.mfMembershipNumber).Value) = INVALID_NUMBER Then mvClassFields.Item(MemberFields.mfMembershipNumber).Value = CStr(mvEnv.GetControlNumber("M"))
        If Len(mvClassFields.Item(MemberFields.mfMemberNumber).Value) = 0 Then mvClassFields.Item(MemberFields.mfMemberNumber).Value = mvClassFields.Item(MemberFields.mfMembershipNumber).Value
        mvEnv.InsertWithExtendedAmendmentHistory(mvEnv.Connection, "members", mvClassFields, mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue)
        mvEnv.AddJournalRecord(JournalTypes.jnlMember, JournalOperations.jnlInsert, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue, (mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)

        'ADD THE MEMBERSHIP ACTIVITY AND SUPPRESSION
        If mvChangeMembershipType Then
          vDate = mvCMTDate
          If IsDate(mvCMTDate) = False Then vDate = TodaysDate()
        Else
          vDate = mvClassFields.Item(MemberFields.mfJoined).Value
        End If
        If vCancellation Then
          AddCancelledActivityAndSuppression(vDate)
        Else
          AddActivityAndSuppression(vDate)
          'ADD THE MEMBER FUTURE TYPE
          AddMemberFutureType(mvRenewalDate, mvPayPlanTerm, mvDateOfBirth, mvRenewalPending, mvPayPlanTermUnits)
        End If
        'Handle MembershipGroups
        If mvChangeMembershipType = False And vCancellation = False Then
          'Add MembershipGroup for new Member
          AddDefaultMembershipGroup(True)
        End If
        'J1447: Create registered users record on member creation, after member record (and associated records) inserted
        If Not vCancellation AndAlso mvEnv.GetConfigOption("me_create_registered_user") Then
          AddRegisteredUser(pWarningMessage)
        End If
      End If
      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub
    Public Function FutureTypeChangeDate(ByRef pRenewalDate As String, ByRef pTerm As Integer, ByRef pDateOfBirth As String, ByRef pRenewalPending As Boolean, ByRef pTermUnits As PaymentPlan.OrderTermUnits) As String
      Dim vDate As String
      Dim vToDate As String = ""
      Dim vActivityDate As String
      Dim vAgeAtRenewal As Integer

      Select Case mvMembershipType.SubsequentTrigger
        Case "F" 'Fixed progression, will change after 1 year
          If Len(pRenewalDate) > 0 And pTerm <> 0 Then
            vDate = pRenewalDate
            If pRenewalPending Then
              If pTerm >= 0 Then
                Select Case pTermUnits
                  Case PaymentPlan.OrderTermUnits.otuMonthly
                    vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, pTerm, CDate(vDate)))
                  Case PaymentPlan.OrderTermUnits.otuWeekly
                    vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.WeekOfYear, pTerm, CDate(vDate)))
                  Case Else 'otuNone
                    vToDate = CDate(vDate).AddYears(pTerm).ToString(CAREDateFormat)
                End Select
              Else
                vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(pTerm), CDate(vDate)))
              End If
            Else
              vToDate = pRenewalDate
            End If
          End If
        Case "A" 'Age progression, will change after max junior age reached
          If Len(pRenewalDate) > 0 And pTerm <> 0 And Len(pDateOfBirth) > 0 And mvMembershipType.MaxJuniorAge <> 0 Then
            vDate = pRenewalDate
            If pTerm < 0 Then vDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(pTerm), CDate(vDate)))
            If DateSerial(Year(CDate(vDate)), Month(CDate(pDateOfBirth)), Day(CDate(pDateOfBirth))) > CDate(vDate) Then
              vAgeAtRenewal = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(pDateOfBirth), CDate(vDate)) - 1)
            Else
              vAgeAtRenewal = CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(pDateOfBirth), CDate(vDate)))
            End If
            vToDate = CDate(vDate).AddYears(mvMembershipType.MaxJuniorAge - vAgeAtRenewal).ToString(CAREDateFormat)
          End If
        Case "C" 'Category progression, will change after valid_to on specified category reached
          '*** TA Ref BR 10837: If Category Progression ever reinstated, will need
          '*** modification to DateDiff as per Age Progression above.
          If Len(mvMembershipType.SubsequentTriggerActivity) > 0 And Len(mvMembershipType.SubsequentTriggerActValue) > 0 And Len(pRenewalDate) > 0 And pTerm <> 0 Then
            vActivityDate = mvEnv.Connection.GetValue("SELECT valid_to FROM contact_categories WHERE contact_number = " & mvClassFields.Item(MemberFields.mfContactNumber).Value & " AND activity = '" & mvMembershipType.SubsequentTriggerActivity & "' AND activity_value ='" & mvMembershipType.SubsequentTriggerActValue & "' AND valid_from " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())))
            If Len(vActivityDate) > 0 Then
              vDate = pRenewalDate
              If pTerm < 0 Then vDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(pTerm), CDate(vDate)))
              vToDate = CDate(vDate).AddYears(CInt(DateDiff(Microsoft.VisualBasic.DateInterval.Year, CDate(vDate), CDate(vActivityDate)))).ToString(CAREDateFormat)
            End If
          End If
      End Select
      FutureTypeChangeDate = vToDate
    End Function
    Public Sub AddMemberFutureType(ByRef pRenewalDate As String, ByRef pTerm As Integer, ByRef pDateOfBirth As String, ByRef pRenewalPending As Boolean, ByRef pTermUnits As PaymentPlan.OrderTermUnits)
      Dim vUpdateFields As New CDBFields
      Dim vToDate As String

      'ADD THE MEMBER FUTURE TYPE
      If Len(mvMembershipType.SubsequentMembershipType) > 0 Then
        With vUpdateFields
          .Clear()
          .Add("membership_number", CDBField.FieldTypes.cftLong, mvClassFields.Item(MemberFields.mfMembershipNumber).Value)
          .Add("future_membership_type", CDBField.FieldTypes.cftCharacter, mvMembershipType.SubsequentMembershipType)
          .Add("amended_by", CDBField.FieldTypes.cftCharacter, "automatic")
          .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)

          vToDate = FutureTypeChangeDate(pRenewalDate, pTerm, pDateOfBirth, pRenewalPending, pTermUnits)

          If Len(vToDate) > 0 Then
            .Add("future_change_date", CDBField.FieldTypes.cftDate, CDate(vToDate).ToString(CAREDateFormat))
            If mvFutureMembershipType Is Nothing Then FutureMembershipTypeCode = mvMembershipType.SubsequentMembershipType
            .Add("product", CDBField.FieldTypes.cftCharacter, mvFutureMembershipType.FirstPeriodsProduct)
            .Add("rate", CDBField.FieldTypes.cftCharacter, mvFutureMembershipType.FirstPeriodsRate)
            .Add("amount", CDBField.FieldTypes.cftNumeric, mvFutureMembershipType.ProductRate.Price(CDate(vToDate), ContactNumber))
            mvEnv.Connection.InsertRecord("member_future_type", vUpdateFields)

            'Now add the future membership category
            If mvMembershipType.SubsequentMembershipTypeActivity <> "" And mvMembershipType.SubsequentMembershipTypeActivityValue <> "" Then
              If (mvMembershipType.Activity <> mvMembershipType.SubsequentMembershipTypeActivity) Or (mvMembershipType.ActivityValue <> mvMembershipType.SubsequentMembershipTypeActivityValue) Then
                Dim vCC As New ContactCategory(mvEnv)
                vCC.ContactTypeSaveActivity(mvContactType, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvMembershipType.SubsequentMembershipTypeActivity, mvMembershipType.SubsequentMembershipTypeActivityValue, mvClassFields.Item(MemberFields.mfSource).Value, vToDate, vToDate, "", ContactCategory.ActivityEntryStyles.aesNormal)
              End If
            End If
          Else
            RaiseError(DataAccessErrors.daeInsertFailed, "member_future_type", "Future Change Date could not be calculated for Membership Type '" & mvMembershipType.SubsequentMembershipType & "'")
          End If
        End With
      End If
    End Sub
    Public Sub SetCancelled(ByRef pCancellationReason As String, Optional ByRef pCancOn As String = "", Optional ByRef pCancBy As String = "", Optional ByRef pCancSource As String = "")
      mvClassFields.Item(MemberFields.mfCancellationReason).Value = pCancellationReason
      If Len(pCancBy) = 0 Then mvClassFields.Item(MemberFields.mfCancelledBy).Value = "automatic" Else mvClassFields.Item(MemberFields.mfCancelledBy).Value = pCancBy
      If Len(pCancOn) = 0 Then mvClassFields.Item(MemberFields.mfCancelledOn).Value = TodaysDate() Else mvClassFields.Item(MemberFields.mfCancelledOn).Value = pCancOn
      If Len(pCancSource) > 0 Then mvClassFields.Item(MemberFields.mfCancellationSource).Value = pCancSource
    End Sub

    Public Sub SaveChangeMembership(ByRef pCancellationReason As String, ByRef pNewMembershipTypeCode As String, ByVal pNewMembershipType As MembershipType, ByVal pPaymentPlanRenewalDate As String, ByVal pPaymentPlanRenewalPending As Boolean, ByVal pPaymentPlanTerm As Integer, ByVal pBatchNumber As Integer, ByVal pTransactionNumber As Integer, ByVal pPaymentPlanTermUnits As PaymentPlan.OrderTermUnits, ByVal pWarningMessage As String, ByVal pCMTDate As String)
      Dim vOldVotingRights As Boolean
      Dim vTransaction As Boolean
      Dim vOldMemTypeCode As String
      Dim vOldMShipNumber As Integer
      Dim vOldBranchCode As String

      mvChangeMembershipType = True
      mvCMTDate = pCMTDate
      If IsDate(mvCMTDate) = False Then mvCMTDate = TodaysDate()

      If Len(pPaymentPlanRenewalDate) > 0 Then
        'Update the order flags on this members class as we are CMTing
        mvRenewalDate = pPaymentPlanRenewalDate
        mvRenewalPending = pPaymentPlanRenewalPending
        mvPayPlanTerm = pPaymentPlanTerm
        mvPayPlanTermUnits = pPaymentPlanTermUnits
      End If

      vOldMemTypeCode = mvClassFields.Item(MemberFields.mfMembershipType).SetValue
      vOldMShipNumber = MembershipNumber
      vOldBranchCode = mvClassFields.Item(MemberFields.mfBranch).SetValue

      If Not mvEnv.Connection.InTransaction Then
        vTransaction = True
        mvEnv.Connection.StartTransaction()
      End If

      'Class may have been updated already with new Member details so ensure cancellation of old member does not update other info
      'Mostly just Joined date & M/Ship card details
      Dim vCMTParams As New CDBParameters
      For Each vClassField As ClassField In mvClassFields
        If vClassField.ValueChanged Then
          vCMTParams.Add(vClassField.ProperName, vClassField.Value)
          vClassField.Value = vClassField.SetValue
        End If
      Next
      If vCMTParams.ContainsKey("ReprintMshipCard") Then vCMTParams.Add("ReprintMembershipCard", vCMTParams("ReprintMshipCard").Value)

      With mvClassFields
        .Item(MemberFields.mfCancellationReason).Value = pCancellationReason
        .Item(MemberFields.mfCancelledBy).Value = mvEnv.User.UserID
        .Item(MemberFields.mfCancelledOn).Value = mvCMTDate
      End With

      'Cancel Old Membership Record
      mvClassFields.Save(mvEnv, mvExisting, mvEnv.User.UserID, mvEnv.AuditStyle = CDBEnvironment.AuditStyleTypes.ausExtended)
      mvEnv.AddJournalRecord(JournalTypes.jnlMember, JournalOperations.jnlCancel, mvClassFields.Item(MemberFields.mfContactNumber).IntegerValue, mvClassFields.Item(MemberFields.mfAddressNumber).IntegerValue, (mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue), 0, 0, pBatchNumber, pTransactionNumber)

      'Cancel Membership Activity & Suppression - Remove Future record
      CancelOtherMembershipFlags(mvCMTDate)

      'Insert New Membership record
      If vCMTParams.Count > 0 Then Update(vCMTParams)
      With mvClassFields
        If vCMTParams.ContainsKey("Joined") AndAlso IsDate(vCMTParams("Joined").Value) Then .Item(MemberFields.mfJoined).Value = vCMTParams("Joined").Value
        .Item(MemberFields.mfCancellationReason).Value = ""
        .Item(MemberFields.mfCancelledBy).Value = ""
        .Item(MemberFields.mfCancelledOn).Value = ""
        .Item(MemberFields.mfMemberNumber).Value = GetMemberNumber(mvClassFields.Item(MemberFields.mfMemberNumber).Value)
      End With
      SetMembershipCardIssueNumber(SetCardIssueNumberTypes.scintReinitialise)
      vOldVotingRights = mvMembershipType.VotingRights
      SetMembershipType(pNewMembershipTypeCode, pNewMembershipType)
      If vOldVotingRights = False And mvMembershipType.VotingRights = True Then
        'Reset Joined and branch applied dates
        mvClassFields.Item(MemberFields.mfBranchMember).Value = "N"
        mvClassFields.Item(MemberFields.mfApplied).Value = mvClassFields.Item(MemberFields.mfJoined).Value
        mvClassFields.Item(MemberFields.mfAccepted).Value = ""
      End If
      mvClassFields.Item(MemberFields.mfMembershipNumber).Value = CStr(INVALID_NUMBER)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdvanceCMT) Then mvClassFields.Item(MemberFields.mfCmtDate).Value = mvCMTDate
      mvExisting = False
      mvClassFields.ClearSetValues()

      Save("", False, pBatchNumber, pTransactionNumber, pWarningMessage)

      'Set MembershipGroups (this has to be done here as we need the new MembershipNumber)
      SetMembershipGroupsForCMT(vOldMemTypeCode, vOldMShipNumber, vOldBranchCode)

      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      If Len(mvClassFields.Item(MemberFields.mfCancelledOn).SetValue) > 0 And Len(mvClassFields.Item(MemberFields.mfCancelledOn).Value) = 0 Then
        'Member has been reinstated
        ReinstateMembershipGroups()
      End If
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub SetJoinedDate(Optional ByRef pDefault As String = "")
      mvClassFields.Item(MemberFields.mfJoined).Value = GetNewJoinedDate(pDefault)
    End Sub

    Public Sub SetMembershipType(ByVal pMembershipTypeCode As String, Optional ByVal pMembershipType As MembershipType = Nothing)
      mvClassFields.Item(MemberFields.mfMembershipType).Value = pMembershipTypeCode
      mvMembershipType = Nothing
      If pMembershipTypeCode.Length > 0 Then
        mvMembershipType = New MembershipType(mvEnv)
        If pMembershipType Is Nothing Then
          mvMembershipType.Init(pMembershipTypeCode)
        Else
          mvMembershipType = pMembershipType
        End If
      End If
    End Sub

    Public Sub SetUnCancelled()
      mvClassFields.Item(MemberFields.mfCancellationReason).Value = ""
      mvClassFields.Item(MemberFields.mfCancelledOn).Value = ""
      mvClassFields.Item(MemberFields.mfCancelledBy).Value = ""
      mvClassFields.Item(MemberFields.mfCancellationSource).Value = ""
    End Sub

    Public Function GetNewJoinedDate(Optional ByRef pDefault As String = "") As String
      Dim vDate As String

      vDate = mvEnv.GetStartDate(CDBEnvironment.ppType.pptMember)
      If vDate = "" Then
        If IsDate(pDefault) Then
          vDate = pDefault
        Else
          vDate = TodaysDate()
        End If
      End If
      GetNewJoinedDate = vDate
    End Function

    Public Sub Create(ByVal pParams As CDBParameters)
      'Used by Web Services and Smart Client only
      Dim vMembershipType As MembershipType

      If pParams.ParameterExists("TransactionType").Value = "MEMC" Then
        'CMT - mvMembershipType must stay as the old MembershipType as it is required in SaveChangeMembership
        vMembershipType = mvEnv.MembershipType((pParams("MembershipType").Value))
      Else
        vMembershipType = mvMembershipType
      End If

      With mvClassFields
        If vMembershipType.MembersPerOrder > 0 And vMembershipType.SetNumberOfMembers = True Then
          .Item(MemberFields.mfNumberOfMembers).Value = pParams.OptionalValue("NumberOfMembers", "1")
        Else
          .Item(MemberFields.mfNumberOfMembers).Value = "1"
        End If
        .Item(MemberFields.mfAgeOverride).Value = pParams.ParameterExists("AgeOverride").Value
        .Item(MemberFields.mfBranch).Value = pParams("Branch").Value
        .Item(MemberFields.mfJoined).Value = pParams("Joined").Value
        If pParams.Exists("BranchMember") Then
          .Item(MemberFields.mfBranchMember).Bool = pParams("BranchMember").Bool
        Else
          .Item(MemberFields.mfBranchMember).Bool = vMembershipType.BranchMembership
        End If
        .Item(MemberFields.mfApplied).Value = pParams.ParameterExists("Applied").Value
        .Item(MemberFields.mfAccepted).Value = pParams.ParameterExists("Accepted").Value
        .Item(MemberFields.mfVotingRights).Bool = vMembershipType.VotingRights
        .Item(MemberFields.mfSource).Value = pParams.OptionalValue("MemberSource", (pParams("Source").Value))
        If (mvEnv.GetConfigOption("enter_member_number") = True And mvEnv.GetConfig("member_number_format") = "char_seq_integer") Then
          'User must supply MemberNumber
          .Item(MemberFields.mfMemberNumber).Value = pParams("MemberNumber").Value
        End If
        If pParams.ParameterExists("TransactionType").Value = "MEMC" Then
          .Item(MemberFields.mfMembershipCardExpires).Value = pParams.ParameterExists("MembershipCardExpires").Value
          .Item(MemberFields.mfReprintMshipCard).Bool = pParams.ParameterExists("ReprintMshipCard").Bool
          .Item(MemberFields.mfMembershipCardIssueNumber).Value = pParams.OptionalValue("MembershipCardIssueNumber", "1")
        Else
          .Item(MemberFields.mfReprintMshipCard).Bool = False
          .Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
        End If
        If pParams.ParameterExists("MemberDOB").Value.Length > 0 Then ContactDateOfBirth = pParams("MemberDOB").Value
        If pParams.ParameterExists("MemberDobEstimated").Value.Length > 0 Then ContactDOBEstimated = pParams("MemberDobEstimated").Bool
        If pParams.ParameterExists("BranchOwnershipGroup").Value.Length > 0 Then ContactOwnershipGroup = pParams("BranchOwnershipGroup").Value
        If pParams.Exists("MembershipStatus") Then .Item(MemberFields.mfMembershipStatus).Value = pParams("MembershipStatus").Value
        If pParams.Exists("LockBranch") Then .Item(MemberFields.mfLockBranch).Value = pParams("LockBranch").Value
        If pParams.ParameterExists("TransactionType").Value <> "MEMC" Then
          'Do not do this for CMT
          If pParams.Exists("MemberCancellationReason") = True Or pParams.Exists("CancellationReason") = True Then
            .Item(MemberFields.mfCancellationReason).Value = pParams.OptionalValue("MemberCancellationReason", (pParams("CancellationReason").Value))
            .Item(MemberFields.mfCancelledBy).Value = pParams.OptionalValue("MemberCancelledBy", (pParams("CancelledBy").Value))
            .Item(MemberFields.mfCancelledOn).Value = pParams.OptionalValue("MemberCancelledOn", (pParams("CancelledOn").Value))
            .Item(MemberFields.mfCancellationSource).Value = pParams.OptionalValue("MemberCancellationSource", (pParams("CancellationSource").Value))
          End If
        End If
      End With
      'mfContactNumber, mfAddressNumber, mfMembershipType =  Set by PaymentPlan.AddMember

    End Sub
    Public Function FutureMembershipTypeValidation() As FutureMembershipTypeErrors
      Dim vError As FutureMembershipTypeErrors
      Dim vPP As PaymentPlan
      Dim vFMT As FutureMembershipType

      vError = FutureMembershipTypeErrors.fmteNone
      If Not MembershipType.FutureTypeValidation Then vError = FutureMembershipTypeErrors.fmteMembershipType
      vPP = New PaymentPlan
      vPP.Init(mvEnv, PaymentPlanNumber)
      If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(TodaysDate()), CDate(vPP.RenewalDate)) < 0 Then
        vFMT = New FutureMembershipType(mvEnv)
        vFMT.Init(MembershipNumber)
        If Not vFMT.Existing Then vError = FutureMembershipTypeErrors.fmteHistoricRenewalDate
      End If
      FutureMembershipTypeValidation = vError
    End Function

    Public Sub SetAmended(ByRef pAmendedOn As String, ByRef pAmendedBy As String)
      mvClassFields.Item(MemberFields.mfAmendedOn).Value = pAmendedOn
      mvClassFields.Item(MemberFields.mfAmendedBy).Value = pAmendedBy
      mvAmendedValid = True
    End Sub

    Public Function CheckMemberNumber(ByVal pMemberNumber As String, Optional ByRef pMessage As String = "") As Boolean
      Dim vChar As String
      Dim vCharPos As Integer
      Dim vError As Boolean
      Dim vIntPos As Integer
      Dim vLen As Integer
      Dim vPos As Integer

      If mvEnv.GetConfigOption("check_member_number") Then
        vLen = Len(pMemberNumber)
        If vLen > 0 Then
          'Check member number starts with a character between A and Z
          vChar = Left(pMemberNumber, 1)
          If IsNumeric(vChar) Then vError = True
          If Not vError Then
            If vChar < "A" Or vChar > "Z" Then vError = True
          End If

          If Not vError Then
            vPos = 2 'Start at 2nd character
            'Find position of first number
            While vPos <= vLen And vIntPos = 0
              vChar = Mid(pMemberNumber, vPos, 1)
              If IsNumeric(vChar) Then vIntPos = vPos
              vPos = vPos + 1
            End While
            If vIntPos = 0 And (vPos - 1 = 8) Then vError = True

            If Not vError Then
              'Ensure that no charactrs appear after the first number
              While vPos <= vLen And vCharPos = 0
                vChar = Mid(pMemberNumber, vPos, 1)
                If IsNumeric(vChar) = False Then vCharPos = vPos
                vPos = vPos + 1
              End While
              If vCharPos > 0 Then vError = True
            End If
          End If
        End If
      End If

      If vError Then pMessage = "Invalid Member Number format"
      CheckMemberNumber = Not vError
    End Function

    Public Sub AddSponsorActivity(ByVal pPaymentPlan As PaymentPlan, ByVal pPayerContactNo As Integer, ByVal pContactType As Contact.ContactTypes, Optional ByVal pCMT As Boolean = False, Optional ByVal pChangePayer As Boolean = False)
      Dim vFromDate As String = ""
      Dim vToDate As String
      Dim vSponsorAct As String
      Dim vSponsorActValue As String

      vSponsorAct = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivity)
      vSponsorActValue = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlSponsorActivityValue)

      If Len(vSponsorAct) > 0 And pPaymentPlan.GiftMembership Then
        If pCMT = True Or pChangePayer = True Then
          If pCMT Then vFromDate = TodaysDate() 'Change membership type
          If pChangePayer Then vFromDate = pPaymentPlan.NextPaymentDue 'Change Payer
          If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(vFromDate), CDate(pPaymentPlan.RenewalDate)) > 0 Then
            'Renewal date is today or in the future
            vToDate = pPaymentPlan.RenewalDate
          Else
            vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, MembershipType.SuspensionGrace, CDate(vFromDate)))
          End If
        Else
          'New member
          vFromDate = Joined
          If pPaymentPlan.FixedRenewalCycle Then
            vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, MembershipType.SuspensionGrace, CDate(pPaymentPlan.RenewalDate)))
          Else
            vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, MembershipType.SuspensionGrace, CDate(vFromDate)))
          End If
        End If
        'Only deduct 1 day if the dates are not the same (e.g. if SuspensionGrace = 0 then both dates are the same)
        If CDate(vFromDate) <> CDate(vToDate) Then vToDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.DayOfYear, -1, CDate(vToDate))) 'Minus one day

        Dim vCC As New ContactCategory(mvEnv)
        vCC.ContactTypeSaveActivity(pContactType, pPayerContactNo, vSponsorAct, vSponsorActValue, Source, vFromDate, vToDate, "", ContactCategory.ActivityEntryStyles.aesNormal)
      End If
    End Sub

    Public Sub SetMembershipCardIssueNumber(ByVal pType As SetCardIssueNumberTypes)
      Select Case pType
        Case SetCardIssueNumberTypes.scintDecrement
          mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = CStr(MembershipCardIssueNumber - 1)
          If mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).IntegerValue = 0 Then mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
        Case SetCardIssueNumberTypes.scintIncrement
          mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = CStr(MembershipCardIssueNumber + 1)
        Case SetCardIssueNumberTypes.scintReinitialise
          mvClassFields.Item(MemberFields.mfMembershipCardIssueNumber).Value = "1"
      End Select
    End Sub

    Public Function CanCancel() As Boolean
      CanCancel = mvExisting And Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) = 0
    End Function
    Public Function CanCMT() As Boolean
      Dim vCan As Boolean

      vCan = mvExisting And Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) = 0
      If vCan Then vCan = mvEnv.Connection.GetCount("fp_applications", Nothing, "fp_application = '" & mvEnv.GetConfig("trader_application_cmt") & "'") > 0
      If vCan Then vCan = MembershipType.MembersPerOrder <> 0 Or mvEnv.GetConfigOption("opt_me_allow_group_change")
      If vCan Then
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan
          mvPaymentPlan.Init(mvEnv, (mvClassFields.Item(MemberFields.mfPaymentPlanNumber).IntegerValue))
        End If
        vCan = mvPaymentPlan.Existing And Not mvPaymentPlan.Provisional
      End If
      CanCMT = vCan
    End Function
    Public Function CanFMT() As Boolean
      Dim vCan As Boolean

      vCan = mvExisting And Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) = 0
      If vCan Then
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan
          mvPaymentPlan.Init(mvEnv, (mvClassFields.Item(MemberFields.mfPaymentPlanNumber).IntegerValue))
        End If
        vCan = mvPaymentPlan.Existing And mvPaymentPlan.CanEditFutureMembership
      End If
      CanFMT = vCan
    End Function
    Public Function CanMaintain() As Boolean
      Dim vCan As Boolean
      vCan = mvExisting = True And Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) = 0
      CanMaintain = vCan
    End Function
    Public Function CanReinstate() As Boolean
      Dim vCan As Boolean

      vCan = mvExisting And Len(mvClassFields.Item(MemberFields.mfCancellationReason).Value) > 0 And mvEnv.GetConfigOption("me_membership_reinstatement")
      If vCan Then
        If mvPaymentPlan Is Nothing Then
          mvPaymentPlan = New PaymentPlan
          mvPaymentPlan.Init(mvEnv, (mvClassFields.Item(MemberFields.mfPaymentPlanNumber).IntegerValue))
        End If
        vCan = mvPaymentPlan.Existing And mvPaymentPlan.TermUnits = PaymentPlan.OrderTermUnits.otuMonthly
      End If
      CanReinstate = vCan
    End Function
    Public Function CanReprintCard() As Boolean
      Dim vCan As Boolean
      vCan = mvExisting And mvClassFields.Item(MemberFields.mfCancellationReason).Value.Length = 0
      If vCan Then vCan = MembershipType.MembershipCard
      Return vCan
    End Function

    Public Sub SCAddMemberSummary(ByVal pContact As Contact, ByVal pJoined As String, ByVal pBranchCode As String, ByVal pBranchMember As Boolean, ByVal pAppliedDate As String, ByVal pDistributionCode As String, ByVal pAgeOverride As String, ByVal pContactDOB As String, Optional ByVal pCMTMembershipTypeCode As String = "", Optional ByVal pCMTMembershipNumber As Integer = 0)
      'Used by Smart Client & Web Services to set the members for the MembershipMembersSummary grid
      'pCMTMembershipTypeCode used by CMT to set the new MembershipTypeCode (this is not expected to be updated)
      mvContact = pContact
      With mvClassFields
        .Item(MemberFields.mfContactNumber).Value = CStr(mvContact.ContactNumber)
        .Item(MemberFields.mfAddressNumber).Value = CStr(mvContact.Address.AddressNumber)
        .Item(MemberFields.mfJoined).Value = pJoined
        .Item(MemberFields.mfBranch).Value = pBranchCode
        .Item(MemberFields.mfBranchMember).Bool = pBranchMember
        .Item(MemberFields.mfApplied).Value = pAppliedDate
        .Item(MemberFields.mfAgeOverride).Value = pAgeOverride
        If pCMTMembershipTypeCode.Length > 0 Then .Item(MemberFields.mfMembershipType).Value = pCMTMembershipTypeCode
        If pCMTMembershipNumber > 0 Then .Item(MemberFields.mfMembershipNumber).Value = CStr(pCMTMembershipNumber)
      End With
      mvDateOfBirth = pContactDOB
      mvDOBEstimated = mvContact.DobEstimated
      mvDistributionCode = pDistributionCode
    End Sub

    Public Function GetSummaryDataAsParameters() As CDBParameters
      'Used by Smart Client to set the members for the MembershipMembersSummary grid
      Dim vParams As New CDBParameters

      With vParams
        .Add("AddressNumber", Contact.Address.AddressNumber)
        .Add("MembershipNumber", CDBField.FieldTypes.cftCharacter, If(mvClassFields.Item(MemberFields.mfMembershipNumber).IntegerValue > 0, mvClassFields.Item(MemberFields.mfMembershipNumber).Value, ""))
        .Add("DateOfBirth", CDBField.FieldTypes.cftCharacter, mvDateOfBirth)
        .Add("DOBEstimated", CDBField.FieldTypes.cftCharacter, If(mvDOBEstimated = True, "Y", "N"))
        .Add("AgeOverride", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfAgeOverride).Value)
        .Add("ContactNumber", Contact.ContactNumber)
        .Add("MembershipType", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfMembershipType).Value)
        .Add("ContactName", CDBField.FieldTypes.cftCharacter, Contact.Name)
        .Add("Joined", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfJoined).Value)
        .Add("Branch", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfBranch).Value)
        .Add("BranchMember", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfBranchMember).Value)
        .Add("Applied", CDBField.FieldTypes.cftCharacter, mvClassFields.Item(MemberFields.mfApplied).Value)
        .Add("DistributionCode", CDBField.FieldTypes.cftCharacter, mvDistributionCode)
        .Add("AddressLine", CDBField.FieldTypes.cftCharacter, Contact.AccessCheckAddressLine)
      End With

      GetSummaryDataAsParameters = vParams

    End Function

    Public Function LineDataType(ByRef pAttributeName As String) As CDBField.FieldTypes
      Select Case pAttributeName
        Case "DateOfBirth"
          LineDataType = CDBField.FieldTypes.cftDate
        Case "DistributionCode", "DOBEstimated", "ContactName", "AddressLine"
          LineDataType = CDBField.FieldTypes.cftCharacter
        Case "LineNumber"
          LineDataType = CDBField.FieldTypes.cftLong
        Case Else
          LineDataType = mvClassFields.ItemDataType(pAttributeName)
      End Select
    End Function

    Public Sub Cancel(ByVal pCancellationReason As String, Optional ByVal pCancellationSource As String = "")
      With mvClassFields
        .Item(MemberFields.mfCancellationReason).Value = pCancellationReason
        .Item(MemberFields.mfCancelledBy).Value = mvEnv.User.UserID
        .Item(MemberFields.mfCancelledOn).Value = TodaysDate()
        If pCancellationSource.Length > 0 Then .Item(MemberFields.mfCancellationSource).Value = pCancellationSource
      End With
    End Sub

    Sub SynchroniseBranch(ByVal pAddressBranch As String, ByVal pNewBranch As String, ByVal pOrigMemBranch As String, ByRef pInformationMessage As String, ByRef pConfirmChangeBranch As Boolean, ByRef pChangeBranch As Boolean)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vAddressChangeWithBranch As String
      Dim vMemberCount As Integer
      Dim vPayPlanCount As Integer
      Dim vAddress As New Address(mvEnv)

      If mvEnv.GetConfigOption("me_synchronise_branch", False) Then
        If pAddressBranch <> pNewBranch Then
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
          vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
          vUpdateFields.Add("branch", CDBField.FieldTypes.cftCharacter, pNewBranch)
          mvEnv.Connection.UpdateRecords("addresses", vUpdateFields, vWhereFields)
          UpdateMembersBranch(mvEnv, AddressNumber, AddressNumber, pAddressBranch, pNewBranch)
        End If
      Else
        vAddressChangeWithBranch = mvEnv.GetConfig("cd_address_change_with_branch")
        If Len(vAddressChangeWithBranch) = 0 Then vAddressChangeWithBranch = "N"
        If vAddressChangeWithBranch <> "N" And Len(pOrigMemBranch) > 0 Then
          If pNewBranch <> pOrigMemBranch Then
            vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
            vWhereFields.Add("branch", CDBField.FieldTypes.cftCharacter, pOrigMemBranch)
            vWhereFields.Add("cancellation_reason", CDBField.FieldTypes.cftCharacter)
            vMemberCount = mvEnv.Connection.GetCount("members", vWhereFields)
            vPayPlanCount = mvEnv.Connection.GetCount("orders", vWhereFields)
            If (vAddressChangeWithBranch <> "P" And vMemberCount > 0) Or (vAddressChangeWithBranch = "P" And (vMemberCount > 0 Or vPayPlanCount > 0)) Then
              If pConfirmChangeBranch Then
                RaiseError(DataAccessErrors.daeAddressChangePromptSynchBranch, vAddressChangeWithBranch)
              Else
                If pChangeBranch Then
                  UpdateMembersBranch(mvEnv, AddressNumber, AddressNumber, pAddressBranch, pNewBranch)
                  vUpdateFields.AddAmendedOnBy(mvEnv.User.UserID)
                  vUpdateFields.Add("branch", CDBField.FieldTypes.cftCharacter, pNewBranch)
                  If vAddressChangeWithBranch = "P" And vPayPlanCount > 0 Then mvEnv.Connection.UpdateRecords("orders", vUpdateFields, vWhereFields, False)
                End If
              End If
            End If
          End If
        End If
      End If
    End Sub

    Public Shared Sub UpdateMembersBranch(pEnv As CDBEnvironment, pOldAddressNumber As Integer, pNewAddressNumber As Integer, pOldBranch As String, pNewBranch As String)
      If pOldAddressNumber <> pNewAddressNumber OrElse pNewBranch <> pOldBranch Then
        Dim vMember As Member = New Member
        vMember.Init(pEnv)
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("m.address_number", CDBField.FieldTypes.cftLong, pOldAddressNumber)
        vWhereFields.Add("cancellation_reason") 'Cancellation Reason is null
        Dim vAnsiJoins As New AnsiJoins
        vAnsiJoins.Add("membership_types mt", "m.membership_type", "mt.membership_type")
        vAnsiJoins.Add("contacts c", "m.contact_number", "c.contact_number")
        vAnsiJoins.AddLeftOuterJoin("contact_addresses ca", "ca.contact_number", "m.contact_number", "ca.address_number", "m.address_number")
        vAnsiJoins.AddLeftOuterJoin("organisation_addresses oa", "oa.organisation_number", "m.contact_number", "oa.address_number", "m.address_number")
        Dim vFields As String = vMember.GetRecordSetFields(Member.MemberRecordSetTypes.mrtAll) & ",mt.branch_membership,c.contact_type,ca.valid_from AS ca_valid_from,oa.valid_from AS oa_valid_from"
        Dim vSQL As New SQLStatement(pEnv.Connection, vFields, "members m", vWhereFields, "", vAnsiJoins)
        Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
        While vRecordSet.Fetch() = True
          vMember = New Member
          vMember.InitFromRecordSet(pEnv, vRecordSet, Member.MemberRecordSetTypes.mrtAll)
          Dim vUpdateMember As Boolean = True
          If pOldAddressNumber <> pNewAddressNumber Then
            'BR20708: only update the member address where the address exists against the contact/organisation
            Dim vUpdateAddress As Boolean = False
            If vMember.ContactType = Access.Contact.ContactTypes.ctcOrganisation Then
              Dim vOrgAddress As New OrganisationAddress(pEnv)
              vOrgAddress.Init(vMember.ContactNumber, pNewAddressNumber)
              If vOrgAddress.Existing Then
                vUpdateAddress = True
              End If
            Else
              Dim vContactAddress As New ContactAddress(pEnv)
              vContactAddress.Init(vMember.ContactNumber, pNewAddressNumber)
              If vContactAddress.Existing Then
                vUpdateAddress = True
              End If
            End If
            If vUpdateAddress Then
              vMember.AddressNumber = pNewAddressNumber
            Else
              vUpdateMember = False
            End If
          End If
          If vUpdateMember AndAlso Not pOldBranch.Equals(pNewBranch) Then
            ' BR19600 if the new Branch is different to the old branch then change it in the membership
            If vMember.LockBranch = False AndAlso pNewBranch <> pOldBranch Then
              If vRecordSet.Fields("branch_membership").Bool Then
                vMember.ChangeMembershipGroupsBranch(vMember.Branch, pNewBranch)
              End If
              Dim vAppliedDate As String = ""
              If vRecordSet.Fields("contact_type").Value = "O" Then
                vAppliedDate = vRecordSet.Fields("oa_valid_from").Value
              Else
                vAppliedDate = vRecordSet.Fields("ca_valid_from").Value
              End If
              vMember.Branch = pNewBranch
              If pEnv.GetConfigOption("me_update_applied_with_branch", False) Then
                If vAppliedDate.Length = 0 Then vAppliedDate = TodaysDate()
                vMember.Applied = vAppliedDate
              End If
            End If
          End If
          If vUpdateMember Then
            vMember.Save()
          End If
        End While
        vRecordSet.CloseRecordSet()
        Dim vAddress As New Address(pEnv)
        Dim vMsg As String = ""
        vAddress.SaveBranchHistory(pNewAddressNumber, pOldBranch, pNewBranch, vMsg)
      End If
    End Sub

    Public Function CalculateMembershipCardExpiryDate(ByVal pReprintMShipCard As Boolean, ByVal pMembershipCardExpires As String, ByVal pRenewalDate As String, ByVal pJoined As String, ByVal pPlanTerm As Integer, ByVal pMembershipCardDuration As Integer, ByVal pPlanStartDate As String, ByVal pRenewalPending As Boolean, ByVal pResetOrderTerm As String, ByVal pMembershipTypeCode As String, ByVal pMembersPerOrder As Integer, ByVal pFixedRenewalM As String) As String
      'Using the given parameters, calculate the card expiry date and return it to the calling process to deal with
      'Used by MembershipCards and Reports
      Dim vMembershipType As MembershipType
      Dim vCardExpiryDate As String

      'Summary:
      '1) If reprint_mship_card = 'Y' And membership_card_expires = Null Then membership_card_expires = renewal_date
      '2) If membership_card_expires = Null And members_per_order = 0 and fixed_renewal_M config not set Then membership_card_expires = renewal_date
      '3) If reprint_mship_card = 'N' And membership_card_expires = Null Then
      ' (a) If order_term < 0 And reset_order_term <> 'N' Then membership_card_expires = order_term + order/joined date (dependent upon fixed_renewal_M config)
      ' (b) If order_term >= 0 Or reset_order_term = 'N' Then membership_card_expires = membership_card_duration + order/joined date (dependent upon fixed_renewal_M config)
      '4) If reprint_mship_card = 'N' And membership_card_expires Is Not Null Then membership_card_expires = membership_card_duration + membership_card_expires

      vMembershipType = mvEnv.MembershipType(pMembershipTypeCode)
      vCardExpiryDate = ""
      If pRenewalPending Then
        If pPlanTerm > 0 Then
          pRenewalDate = AddTerm(vMembershipType, pRenewalDate, pPlanTerm)
        Else
          pRenewalDate = AddMonths(pPlanStartDate, pRenewalDate, System.Math.Abs(pPlanTerm))
        End If
      End If

      If pReprintMShipCard Then
        vCardExpiryDate = pMembershipCardExpires
        If String.IsNullOrEmpty(pMembershipCardExpires) Then
          vCardExpiryDate = pRenewalDate
        End If
      Else
        If String.IsNullOrEmpty(pMembershipCardExpires) Then
          If (String.IsNullOrEmpty(pFixedRenewalM) And pMembersPerOrder = 0) Then
            'Group Membership so use RenewalDate instead of Joined
            vCardExpiryDate = pRenewalDate
          ElseIf (pPlanTerm < 1 And pResetOrderTerm <> "N") Then
            If Not String.IsNullOrEmpty(pFixedRenewalM) Then
              vCardExpiryDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(pPlanTerm), CDate(pPlanStartDate)))
            Else
              vCardExpiryDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, System.Math.Abs(pPlanTerm), CDate(pJoined)))
            End If
          Else
            If Not String.IsNullOrEmpty(pFixedRenewalM) Then
              vCardExpiryDate = CDate(pPlanStartDate).AddYears(pMembershipCardDuration).ToString(CAREDateFormat)
            Else
              'There is no ExpiryDate, so create it using the Joined date and the Membership Card Duration. Not Fixed Renewal
              vCardExpiryDate = AddTerm(vMembershipType, pJoined, pMembershipCardDuration)
            End If
          End If
        Else
          'There is a Memebership Card Exipires, so it can be used. 
          vCardExpiryDate = AddTerm(vMembershipType, pMembershipCardExpires, pMembershipCardDuration)
        End If
      End If
      Return vCardExpiryDate
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pMemberShipType"></param>
    ''' <param name="pOldDate">The date that we want ot move on.</param>
    ''' <param name="pTermDuration">Number of Months or Years dependent on the value of pMemebershipType.PaymentTerm.</param>
    ''' <returns></returns>
    ''' <remarks>Suitably bland name as this function is intended to to be used for CardExpiryDate and RenewalDate</remarks>
    Private Function AddTerm(ByVal pMemberShipType As MembershipType, ByVal pOldDate As String, ByVal pTermDuration As Integer) As String
      Dim vNewDate As String
      Select Case pMemberShipType.PaymentTerm
        Case MembershipType.MembershipTypeTerms.mtfMonthlyTerm
          vNewDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, pTermDuration, CDate(pOldDate)))
        Case MembershipType.MembershipTypeTerms.mtfWeeklyTerm
          vNewDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.WeekOfYear, pTermDuration, CDate(pOldDate)))
        Case Else
          vNewDate = CDate(pOldDate).AddYears(pTermDuration).ToString(CAREDateFormat)
      End Select
      Return vNewDate
    End Function
    Public Sub SetMembershipGroups(ByVal pEnv As CDBEnvironment, Optional ByVal pMembershipNumber As Integer = 0, Optional ByVal pMembershipTypeCode As String = "", Optional ByVal pNewGroupCode As String = "")
      'Create MembershipGroups data for:
      '(1) All Memberships, or
      '(2) An individual Membership
      '(3) An individual MembershipType
      'pNewGroupCode is used by TableMaintenance when the GroupCode is set/changed
      Dim vRS As CDBRecordSet
      Dim vInsertFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vMaxMGNumber As Integer
      Dim vMGNumber As Integer
      Dim vOrgGroup As String
      Dim vSQL As String
      Dim vValidFrom As String
      Dim vValidTo As String

      vOrgGroup = pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMemberOrganisationGroup)
      If pNewGroupCode.Length > 0 Then vOrgGroup = pNewGroupCode

      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) = True And Len(vOrgGroup) > 0 Then
        vSQL = "SELECT m.membership_number, org.organisation_number, m.applied, mg.valid_from, mg.valid_to, " & pEnv.Connection.DBIsNull("mg.membership_group_number", "0") & " AS membership_group_number"
        vSQL = vSQL & " FROM membership_types mt"
        vSQL = vSQL & " INNER JOIN members m ON mt.membership_type = m.membership_type"
        vSQL = vSQL & " INNER JOIN branches b ON m.branch = b.branch"
        vSQL = vSQL & " INNER JOIN organisations org ON b.organisation_number = org.organisation_number"
        vSQL = vSQL & " LEFT OUTER JOIN membership_groups mg ON m.membership_number = mg.membership_number" 'AND b.organisation_number = mg.organisation_number"
        vSQL = vSQL & " WHERE"
        If pMembershipTypeCode.Length > 0 Then vSQL = vSQL & " mt.membership_type = '" & pMembershipTypeCode & "' AND"
        vSQL = vSQL & " mt.branch_membership = 'Y'"
        If pMembershipNumber > 0 Then vSQL = vSQL & " AND m.membership_number = " & pMembershipNumber
        vSQL = vSQL & " AND m.cancellation_reason IS NULL"
        vSQL = vSQL & " AND org.organisation_group = '" & vOrgGroup & "'"
        vSQL = vSQL & " AND (mg.organisation_number IS NULL OR (mg.organisation_number IS NOT NULL AND b.organisation_number = mg.organisation_number))"
        vSQL = vSQL & " ORDER BY mt.membership_type, m.membership_number"

        With vInsertFields
          .Add("membership_group_number", CDBField.FieldTypes.cftLong)
          .Add("membership_number", CDBField.FieldTypes.cftLong)
          .Add("organisation_number", CDBField.FieldTypes.cftLong)
          .Add("default_group", CDBField.FieldTypes.cftCharacter, "Y")
          .Add("valid_from", CDBField.FieldTypes.cftDate)
          .Add("is_current", CDBField.FieldTypes.cftCharacter, "Y")
          .AddAmendedOnBy(pEnv.User.UserID)
        End With

        With vUpdateFields
          .Add("valid_to", CDBField.FieldTypes.cftDate)
          .Add("is_current", CDBField.FieldTypes.cftCharacter, "Y")
          .AddAmendedOnBy(pEnv.User.UserID)
        End With

        With vWhereFields
          .Add("membership_group_number", CDBField.FieldTypes.cftLong)
        End With

        vRS = pEnv.Connection.GetRecordSetAnsiJoins(vSQL)
        While vRS.Fetch() = True
          If vRS.Fields(6).IntegerValue > 0 Then ' IsDate(vRS.Fields(4).Value) Then
            'Update
            vValidTo = vRS.Fields(5).Value
            If Not (IsDate(vValidTo)) Then vValidTo = CStr(DateSerial(9999, 12, 31))
            If CDate(vValidTo) <= CDate(TodaysDate()) Then
              vWhereFields(1).Value = CStr(vRS.Fields(6).IntegerValue)
              pEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields)
            End If
          Else
            'Insert
            If vMGNumber >= vMaxMGNumber Then
              vMGNumber = pEnv.GetControlNumber("+MG", 100)
              vMaxMGNumber = vMGNumber + 100
            End If
            vMGNumber = pEnv.GetControlNumber("?MG", vMGNumber)
            vMGNumber = vMGNumber + 1

            vValidFrom = vRS.Fields(3).Value
            If Not (IsDate(vValidFrom)) Then vValidFrom = TodaysDate()
            With vInsertFields
              .Item(1).Value = CStr(vMGNumber)
              .Item(2).Value = CStr(vRS.Fields(1).IntegerValue)
              .Item(3).Value = CStr(vRS.Fields(2).IntegerValue)
              .Item(5).Value = vValidFrom
            End With
            pEnv.Connection.InsertRecord("membership_groups", vInsertFields)
          End If
        End While
        vRS.CloseRecordSet()
      End If

    End Sub

    Public Sub SetMembershipGroupsHistoric(ByVal pEnv As CDBEnvironment, Optional ByVal pMembershipNumber As Integer = 0, Optional ByVal pMembershipTypeCode As String = "")
      'Set MembershipGroups data as historic for:
      '(1) All Memberships
      '(2) An individual Membership
      '(3) An individual MembershipType
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vSQL As String

      With vUpdateFields
        .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
        .Add("is_current", CDBField.FieldTypes.cftCharacter, "N")
        .AddAmendedOnBy(pEnv.User.UserID)
      End With

      With vWhereFields
        If pMembershipNumber > 0 Then vWhereFields.Add("membership_number", CDBField.FieldTypes.cftLong, pMembershipNumber)
        .Add("valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        .Add("valid_to#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End With

      If pMembershipTypeCode.Length > 0 Then
        vSQL = "(SELECT m.membership_number FROM membership_types mt, members m WHERE m.membership_type = '" & pMembershipTypeCode & "'"
        vSQL = vSQL & " AND mt.membership_type = m.membership_type AND m.cancellation_reason IS NULL)"
        vWhereFields.Add("membership_number#2", CDBField.FieldTypes.cftLong, vSQL, CDBField.FieldWhereOperators.fwoIn)
      End If

      pEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields, False)

    End Sub

    Public Sub Update(pFields As CDBFields)
      For Each vField As CDBField In pFields
        If mvClassFields.ContainsKey(vField.Name) Then
          mvClassFields.Item(vField.Name).Value = vField.Value
        End If
      Next
    End Sub

    Public Sub Update(ByVal pParams As CDBParameters)
      'Used by Smart Client / Web Services only
      'Dim vMemberGroup    As MembershipGroup

      With mvClassFields
        .Item(MemberFields.mfSource).Value = pParams.OptionalValue("Source", Source)
        .Item(MemberFields.mfMembershipCardExpires).Value = pParams.OptionalValue("MembershipCardExpires", MembershipCardExpires)
        .Item(MemberFields.mfAgeOverride).Value = pParams.OptionalValue("AgeOverride", AgeOverride)
        .Item(MemberFields.mfAddressNumber).Value = pParams.OptionalValue("AddressNumber", CStr(AddressNumber))
        .Item(MemberFields.mfBranch).Value = pParams.OptionalValue("Branch", Branch)
        .Item(MemberFields.mfBranchMember).Value = pParams.OptionalValue("BranchMember", BranchMember)
        .Item(MemberFields.mfApplied).Value = pParams.OptionalValue("Applied", Applied)
        .Item(MemberFields.mfAccepted).Value = pParams.OptionalValue("Accepted", Accepted)
        If pParams.ContainsKey("ReprintMembershipCard") Then
          .Item(MemberFields.mfReprintMshipCard).Bool = pParams("ReprintMembershipCard").Bool
          If .Item(MemberFields.mfReprintMshipCard).ValueChanged Then SetMembershipCardIssueNumber(CType(IIf(pParams("ReprintMembershipCard").Bool = True, SetCardIssueNumberTypes.scintIncrement, SetCardIssueNumberTypes.scintDecrement), SetCardIssueNumberTypes))
        End If
        If pParams.Exists("MembershipStatus") Then .Item(MemberFields.mfMembershipStatus).Value = pParams.OptionalValue("MembershipStatus", MembershipStatus)
        If pParams.Exists("LockBranch") Then .Item(MemberFields.mfLockBranch).Value = pParams.OptionalValue("LockBranch", .Item(MemberFields.mfLockBranch).Value)
        If pParams.ContainsKey("VotingRights") = True AndAlso IsDate(mvCMTDate) = True Then
          .Item(MemberFields.mfVotingRights).Value = pParams("VotingRights").Value
        End If
      End With

      If mvClassFields(MemberFields.mfBranch).ValueChanged Then ChangeMembershipGroupsBranch(mvClassFields.Item(MemberFields.mfBranch).SetValue, Branch)

    End Sub

    Private Sub AddDefaultMembershipGroup(ByVal pDefaultGroup As Boolean)
      Dim vBranch As New Branch
      Dim vMembershipGroup As New MembershipGroup(mvEnv)
      Dim vParams As New CDBParameters
      Dim vValidFrom As String

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) Then
        If mvMembershipType.UseMembershipGroups Then
          vBranch.Init(mvEnv, Branch)
          If vBranch.Existing Then
            vValidFrom = If(mvChangeMembershipType = True, TodaysDate, Applied)
            If Not (IsDate(vValidFrom)) Then vValidFrom = TodaysDate()
            With vParams
              .Add("MembershipNumber", MembershipNumber)
              .Add("OrganisationNumber", vBranch.OrganisationNumber)
              .Add("DefaultGroup", CDBField.FieldTypes.cftCharacter, BooleanString(pDefaultGroup))
              .Add("ValidFrom", CDBField.FieldTypes.cftDate, vValidFrom)
            End With
            vMembershipGroup.Init()
            vMembershipGroup.Create(vParams)
            vMembershipGroup.Save()
          End If
        End If
      End If

    End Sub

    Private Sub SetMembershipGroupsForCMT(ByVal pOldMembershipTypeCode As String, ByVal pOldMembershipNumber As Integer, ByVal pOldBranchCode As String)
      'Cancel MembershipGroups for old Memberships and create for new Membership
      Dim vMembershipGroup As New MembershipGroup(mvEnv)
      Dim vOldMembershipType As MembershipType
      Dim vRS As CDBRecordSet
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vSQL As String

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) = True Then
        vOldMembershipType = mvEnv.MembershipType(pOldMembershipTypeCode)

        vWhereFields = New CDBFields
        With vWhereFields
          .Add("membership_number", CDBField.FieldTypes.cftLong, pOldMembershipNumber)
          .Add("valid_to", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
          .Add("valid_to#2", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
          .Add("organisation_group", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMemberOrganisationGroup))
        End With

        If mvMembershipType.UseMembershipGroups Then
          If vOldMembershipType.UseMembershipGroups Then
            'Both old & new MembershipTypes are using MembershipGroups
            vMembershipGroup = New MembershipGroup(mvEnv)
            vMembershipGroup.Init()
            vSQL = "SELECT " & vMembershipGroup.GetRecordSetFields(MembershipGroup.MembershipGroupRecordSetTypes.mgrtAll Or MembershipGroup.MembershipGroupRecordSetTypes.mgrtBranch) & " FROM membership_groups mg"
            vSQL = vSQL & " INNER JOIN organisations o ON mg.organisation_number = o.organisation_number"
            vSQL = vSQL & " LEFT OUTER JOIN branches b ON o.organisation_number = b.organisation_number"
            vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields)
            vRS = mvEnv.Connection.GetRecordSet(vSQL)
            While vRS.Fetch() = True
              vMembershipGroup = New MembershipGroup(mvEnv)
              With vMembershipGroup
                .InitFromRecordSet(mvEnv, vRS, MembershipGroup.MembershipGroupRecordSetTypes.mgrtAll Or MembershipGroup.MembershipGroupRecordSetTypes.mgrtBranch)
                .CloneForCMT(MembershipNumber, pOldBranchCode, Branch)
                .Save()
              End With
            End While
            vRS.CloseRecordSet()
            'Now check to ensure the new Branch has been included
            vMembershipGroup = New MembershipGroup(mvEnv)
            vMembershipGroup.InitFromMemberAndBranch(mvEnv, MembershipNumber, Branch)
            If vMembershipGroup.Existing = False Then AddDefaultMembershipGroup(False)
          Else
            'Just add a new MembershipGoup for the Members Branch
            AddDefaultMembershipGroup(True)
          End If
        End If

        If vOldMembershipType.UseMembershipGroups Then
          'Update old MembershipGroups to be historic
          vUpdateFields = New CDBFields
          With vUpdateFields
            .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
            .Add("is_current", CDBField.FieldTypes.cftCharacter, "N")
            .AddAmendedOnBy(mvEnv.User.UserID)
          End With
          vWhereFields.Remove("organisation_group")
          mvEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields, False)
        End If
      End If

    End Sub

    Private Sub ReinstateMembershipGroups()
      'Reinstating a Member needs to reinstate any MembershipGroups
      'Dim vMembershipGroup  As New MembershipGroup(mvEnv)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vRS As CDBRecordSet
      Dim vCancelledOn As String
      Dim vDefaultCount As Integer
      'Dim vGotDefault       As Boolean
      Dim vSQL As String

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataMembershipGroups) Then
        If mvMembershipType.UseMembershipGroups Then
          '(1) Reinstate existing MembershipGroups
          vCancelledOn = mvClassFields.Item(MemberFields.mfCancelledOn).SetValue
          With vUpdateFields
            .Add("valid_to", CDBField.FieldTypes.cftDate, "")
            .Add("is_current", CDBField.FieldTypes.cftCharacter, "Y")
            .AddAmendedOnBy(mvEnv.User.UserID)
          End With
          With vWhereFields
            .Add("membership_number", CDBField.FieldTypes.cftLong, MembershipNumber)
            .Add("valid_to", CDBField.FieldTypes.cftDate, vCancelledOn)
            .Add("default_group", CDBField.FieldTypes.cftCharacter, "Y")
          End With

          '(a) Update defaults first
          vDefaultCount = 0
          'vMembershipGroup.Init mvEnv
          vSQL = "SELECT mg.membership_group_number, mgh.change_date"
          vSQL = vSQL & " FROM membership_groups mg"
          vSQL = vSQL & " LEFT OUTER JOIN membership_group_history mgh ON mg.membership_number = mgh.membership_number AND mg.organisation_number = mgh.old_organisation_number"
          vSQL = vSQL & " WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY valid_from DESC"
          vSQL = Replace(vSQL, "WHERE ", "WHERE  mg.")
          vRS = mvEnv.Connection.GetRecordSetAnsiJoins(vSQL)
          vWhereFields.Add("membership_group_number", CDBField.FieldTypes.cftLong)
          While vRS.Fetch() = True
            If Not (IsDate(vRS.Fields(2).Value)) Then
              vDefaultCount = vDefaultCount + 1
              vWhereFields.Item("membership_group_number").Value = CStr(vRS.Fields(1).IntegerValue)
              mvEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields)
            End If
          End While
          vRS.CloseRecordSet()

          '(b) Update everything else
          vWhereFields.Remove("membership_group_number")
          vWhereFields.Item("default_group").Value = "N"
          mvEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields, False)

          '(2) If we were originally left with multiple default Groups then we now need to fix them
          If vDefaultCount > 1 Then
            With vWhereFields
              .Clear()
              .Add("membership_number", CDBField.FieldTypes.cftLong, MembershipNumber)
              .Add("default_group", CDBField.FieldTypes.cftCharacter, "Y")
              .Add("valid_to", CDBField.FieldTypes.cftDate, "")
            End With
            With vUpdateFields
              .Item("valid_to").Value = vCancelledOn
              .Item("is_current").Value = "N"
            End With

            vSQL = "SELECT membership_group_number FROM membership_groups WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " ORDER BY valid_from"
            vRS = mvEnv.Connection.GetRecordSet(vSQL)
            vWhereFields.Add("membership_group_number", CDBField.FieldTypes.cftLong)
            While vRS.Fetch() = True And vDefaultCount > 1
              vWhereFields.Item("membership_group_number").Value = CStr(vRS.Fields(1).IntegerValue)
              mvEnv.Connection.UpdateRecords("membership_groups", vUpdateFields, vWhereFields)
              vDefaultCount = vDefaultCount - 1
            End While
            vRS.CloseRecordSet()
          End If
        End If
      End If

    End Sub

    Public Sub ChangeMembershipGroupsBranch(ByVal pOldBranchCode As String, ByVal pNewBranchCode As String)
      Dim vMembershipGroup As New MembershipGroup(mvEnv)
      Dim vNewBranch As New Branch
      Dim vParams As New CDBParameters
      Dim vTrans As Boolean

      If MembershipType.UseMembershipGroups And (pOldBranchCode <> pNewBranchCode) Then
        'Processing required
        '(1) Set original MembershipGroup as historic
        '(2) Add new MembershipGroup for the new Branch
        '(3) Dedup this new MembershipGroup to ensure that it does not overlap with an existing record for this Branch

        '(1a) Set up parameters collection for the new MembershipGroup
        vNewBranch.Init(mvEnv, pNewBranchCode)
        With vParams
          .Add("MembershipNumber", MembershipNumber)
          .Add("OrganisationNumber", vNewBranch.OrganisationNumber)
          .Add("ValidFrom", CDBField.FieldTypes.cftDate, TodaysDate)
          .Add("ValidTo", CDBField.FieldTypes.cftDate)
          .Add("DefaultGroup", CDBField.FieldTypes.cftCharacter, "N")
        End With

        '(1b) Find existing MembershipGroup for old Branch (if any)
        vMembershipGroup.InitFromMemberAndBranch(mvEnv, MembershipNumber, pOldBranchCode)

        If mvEnv.Connection.InTransaction = False Then
          mvEnv.Connection.StartTransaction()
          vTrans = True
        End If

        If vMembershipGroup.Existing Then
          'If this is the DefaultGroup then set as historic and create a new record for the new Branch
          With vMembershipGroup
            If IsDate(.ValidTo) Then
              If CDate(.ValidTo) > CDate(TodaysDate()) Then vParams("ValidTo").Value = .ValidTo
            End If
            vParams("DefaultGroup").Value = BooleanString(.DefaultGroup)
            If .DefaultGroup Then .SetHistoric(vNewBranch.OrganisationNumber)
          End With
        End If

        '(2) Add new MembershipGroup
        vMembershipGroup = New MembershipGroup(mvEnv)
        vMembershipGroup.Init()
        vMembershipGroup.Create(vParams)

        '(3) Dedup existing Groups to ensure that there is no overlap of the dates
        vMembershipGroup.DedupNewGroup()

        If vTrans Then mvEnv.Connection.CommitTransaction()
      End If

    End Sub

    Private Sub AddRegisteredUser(ByRef pWarningMessage As String)
      Dim vRegisteredUser As New RegisteredUser
      vRegisteredUser.Init(mvEnv, "", "", Contact.ContactNumber.ToString)
      'If registered user record already exists with matching contact number then do not create
      If Not vRegisteredUser.Existing Then
        Dim vCreate As Boolean = True
        'If member contact d.o.b. not set then set warning message and do not create
        If Not IsDate(Contact.DateOfBirth) Then
          If pWarningMessage.Length > 0 Then pWarningMessage &= vbNewLine & vbNewLine
          pWarningMessage &= String.Format(ProjectText.String33023, MemberNumber, "Date of Birth") 'The creation of Member Number {0} Registered User record failed because the Member Contact {1} was not set
          vCreate = False
        End If
        If vCreate Then
          Dim vCommsDataTable As CDBDataTable = Contact.CommsDataTable
          Dim vEmailAddress As String = ""
          'If no active member contact communications email address record then set warning message and do not create
          If vCommsDataTable.Rows.Count > 0 Then
            For Each vRow As CDBDataRow In vCommsDataTable.Rows
              vEmailAddress = vRow.Item("EMailAddress")
            Next
          End If
          If vEmailAddress.Length = 0 Then
            If pWarningMessage.Length > 0 Then pWarningMessage &= vbNewLine & vbNewLine
            pWarningMessage &= String.Format(ProjectText.String33023, MemberNumber, "active Communication Email Address") 'The creation of Member Number {0} Registered User record failed because the Member Contact {1} was not set
            vCreate = False
          End If
          If vCreate Then
            'Create registered user record
            Dim vParams As New CDBParameters
            vParams.Add("UserName", MemberNumber)
            vParams.Add("Password", CDate(Contact.DateOfBirth).ToString("dd/MM/yyyy"))
            vParams.Add("EmailAddress", vEmailAddress)
            vRegisteredUser.Create(vParams, Contact)
          End If
        End If
      Else
        'If registered user record exists with matching contact number but user name is not set to member number
        'then update record to set user name to member number
        If Not vRegisteredUser.UserName = MemberNumber Then
          Dim vParams As New CDBParameters
          vParams.Add("UserName", MemberNumber)
          vRegisteredUser.Update(mvEnv, vParams)
          vRegisteredUser.Save(True)
        End If
      End If
    End Sub

    ''' <summary>Reset the NumberOfMembers to the specified value.</summary>
    ''' <param name="pNumberOfMembers">The total number of members for the membership type in this membership</param>
    ''' <remarks>The NumberOfMembers is only updated when the MembershipType.SetNumberOfMembers = True and the MembershipType.MembersPerOrder is greater than zero.</remarks>
    Public Sub ResetNumberOfMembers(ByVal pNumberOfMembers As Integer)
      If pNumberOfMembers > 0 Then
        If MembershipType.SetNumberOfMembers = True AndAlso MembershipType.MembersPerOrder > 0 Then mvClassFields.Item(MemberFields.mfNumberOfMembers).IntegerValue = pNumberOfMembers
      End If
    End Sub

    ''' <summary>Get the earliest date that can be set for a CMT.</summary>
    ''' <param name="pTermStartDate">The start date of the current renewal period.</param>
    Friend Function GetEarliestCMTDate(ByVal pTermStartDate As Date) As String
      Dim vEarliestDate As String = ""
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbAdvanceCMT) Then
        Dim vWhereFields As New CDBFields(New CDBField("order_number", PaymentPlanNumber))
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, "MAX(cmt_date) AS max_cmt_date", "members m", vWhereFields)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        If vRS.Fetch Then vEarliestDate = vRS.Fields(1).Value
        vRS.CloseRecordSet()
      End If
      If IsDate(vEarliestDate) = False Then vEarliestDate = pTermStartDate.ToString(CAREDateFormat)
      If CDate(vEarliestDate) < pTermStartDate Then vEarliestDate = pTermStartDate.ToString(CAREDateFormat)
      Return vEarliestDate
    End Function

  End Class
End Namespace
