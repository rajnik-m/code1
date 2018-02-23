Namespace Access

  Public Class GaAppropriateCertificate
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum GaAppropriateCertificateFields
      AllFields = 0
      CertificateNumber
      ContactNumber
      StartDate
      EndDate
      CertificateAmount
      TaxStatus
      SignatureDate
      ClaimNumber
      AmountClaimed
      AmountPaid
      CreatedBy
      CreatedOn
      CancelledBy
      CancelledOn
      CancellationReason
      CancellationSource
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("certificate_number", CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("start_date", CDBField.FieldTypes.cftDate)
        .Add("end_date", CDBField.FieldTypes.cftDate)
        .Add("certificate_amount", CDBField.FieldTypes.cftNumeric)
        .Add("tax_status")
        .Add("signature_date", CDBField.FieldTypes.cftDate)
        .Add("claim_number", CDBField.FieldTypes.cftLong)
        .Add("amount_claimed", CDBField.FieldTypes.cftNumeric)
        .Add("amount_paid", CDBField.FieldTypes.cftNumeric)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)
        .Add("cancelled_by")
        .Add("cancelled_on", CDBField.FieldTypes.cftDate)
        .Add("cancellation_reason")
        .Add("cancellation_source")

        .Item(GaAppropriateCertificateFields.ContactNumber).PrefixRequired = True
        .Item(GaAppropriateCertificateFields.CertificateNumber).PrimaryKey = True

        .SetControlNumberField(GaAppropriateCertificateFields.CertificateNumber, "AP")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "gac"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "ga_appropriate_certificates"
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
    Public ReadOnly Property CertificateNumber() As Integer
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CertificateNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(GaAppropriateCertificateFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property StartDate() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.StartDate).Value
      End Get
    End Property
    Public ReadOnly Property EndDate() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.EndDate).Value
      End Get
    End Property
    Public ReadOnly Property CertificateAmount() As Double
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CertificateAmount).DoubleValue
      End Get
    End Property
    Public ReadOnly Property TaxStatus() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.TaxStatus).Value
      End Get
    End Property
    Public ReadOnly Property SignatureDate() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.SignatureDate).Value
      End Get
    End Property
    Public ReadOnly Property ClaimNumber() As Integer
      Get
        Return mvClassFields(GaAppropriateCertificateFields.ClaimNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property AmountClaimed() As Double
      Get
        Return mvClassFields(GaAppropriateCertificateFields.AmountClaimed).DoubleValue
      End Get
    End Property
    Public ReadOnly Property AmountPaid() As Double
      Get
        Return mvClassFields(GaAppropriateCertificateFields.AmountPaid).DoubleValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property CancelledBy() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CancelledBy).Value
      End Get
    End Property
    Public ReadOnly Property CancelledOn() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CancelledOn).Value
      End Get
    End Property
    Public ReadOnly Property CancellationReason() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CancellationReason).Value
      End Get
    End Property
    Public ReadOnly Property CancellationSource() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.CancellationSource).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(GaAppropriateCertificateFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Public Function CanCancel(Optional ByRef pCancelMessage As String = "") As Boolean
      'Check the tax claims where the amount paid is zero
      Dim vCanCancel As Boolean = True
      If CancellationReason.Length > 0 Then vCanCancel = False 'Already Cancelled
      If ClaimNumber > 0 And AmountPaid > 0 Then vCanCancel = False 'Got a valid tax claim
      If Not vCanCancel Then pCancelMessage = "This Certificate has been successfully claimed. It can not be cancelled."
      Return vCanCancel
    End Function

    Public Sub Cancel(ByVal pCancelReason As String, ByVal pCancellationSource As String, Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vCancellationRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT status,cancellation_reason_desc FROM cancellation_reasons WHERE cancellation_reason = '" & pCancelReason & "'")
      If vCancellationRS.Fetch Then
        If vCancellationRS.Fields(1).Value.Length > 0 Then
          Dim vContact As New Contact(mvEnv)
          vContact.Init(ContactNumber)
          If vContact.Existing Then
            vContact.SetStatus(ContactNumber,
                               vCancellationRS.Fields(1).Value,
                               TodaysDate,
                               If(String.IsNullOrWhiteSpace(vContact.StatusReason),
                                  vCancellationRS.Fields(2).Value,
                                  String.Empty))
          End If
        End If
      End If
      vCancellationRS.CloseRecordSet()
      mvClassFields.Item(GaAppropriateCertificateFields.CancellationReason).Value = pCancelReason
      mvClassFields.Item(GaAppropriateCertificateFields.CancelledBy).Value = mvEnv.User.UserID
      mvClassFields.Item(GaAppropriateCertificateFields.CancelledOn).Value = TodaysDate()
      If Len(pCancellationSource) > 0 Then mvClassFields.Item(GaAppropriateCertificateFields.CancellationSource).Value = pCancellationSource
      Save(pAmendedBy, pAudit)
    End Sub
#End Region

  End Class
End Namespace
