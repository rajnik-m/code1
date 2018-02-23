Namespace Access

  Partial Public Class EventBooking

    'Public Enum EventBookingStatuses
    '  ebsBooked = 1 'F
    '  ebsWaiting 'W
    '  ebsBookedTransfer 'X
    '  ebsBookedAndPaid 'B
    '  ebsWaitingPaid 'P
    '  ebsBookedAndPaidTransfer 'Y
    '  ebsBookedCreditSale 'S
    '  ebsWaitingCreditSale 'A
    '  ebsBookedCreditSaleTransfer 'R
    '  ebsBookedInvoiced 'V
    '  ebsWaitingInvoiced 'O
    '  ebsBookedInvoicedTransfer 'D
    '  ebsExternal 'E
    '  ebsCancelled 'C
    '  ebsInterested 'I
    '  ebsAwaitingAcceptance 'T
    '  ebsAmended 'U
    'End Enum

    Public Shared Function GetBookingStatusCode(ByVal pBookingStatus As EventBooking.EventBookingStatuses) As String
      Select Case pBookingStatus
        Case EventBooking.EventBookingStatuses.ebsAmended
          Return "U"
        Case EventBooking.EventBookingStatuses.ebsBooked
          Return "F"
        Case EventBooking.EventBookingStatuses.ebsWaiting
          Return "W"
        Case EventBooking.EventBookingStatuses.ebsBookedTransfer
          Return "X"
        Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
          Return "B"
        Case EventBooking.EventBookingStatuses.ebsWaitingPaid
          Return "P"
        Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
          Return "Y"
        Case EventBooking.EventBookingStatuses.ebsBookedCreditSale
          Return "S"
        Case EventBooking.EventBookingStatuses.ebsWaitingCreditSale
          Return "A"
        Case EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer
          Return "R"
        Case EventBooking.EventBookingStatuses.ebsBookedInvoiced
          Return "V"
        Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced
          Return "O"
        Case EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer
          Return "D"
        Case EventBooking.EventBookingStatuses.ebsExternal
          Return "E"
        Case EventBooking.EventBookingStatuses.ebsCancelled
          Return "C"
        Case EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
          Return "T"
        Case Else ' EventBooking.EventBookingStatuses.ebsInterested
          Return "I"
      End Select
    End Function

    Public Shared Function GetBookingStatus(ByRef pStatusCode As String) As EventBooking.EventBookingStatuses
      Select Case pStatusCode
        Case "F"
          Return EventBooking.EventBookingStatuses.ebsBooked
        Case "W"
          Return EventBooking.EventBookingStatuses.ebsWaiting
        Case "X"
          Return EventBooking.EventBookingStatuses.ebsBookedTransfer
        Case "B"
          Return EventBooking.EventBookingStatuses.ebsBookedAndPaid
        Case "P"
          Return EventBooking.EventBookingStatuses.ebsWaitingPaid
        Case "Y"
          Return EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
        Case "S"
          Return EventBooking.EventBookingStatuses.ebsBookedCreditSale
        Case "A"
          Return EventBooking.EventBookingStatuses.ebsWaitingCreditSale
        Case "R"
          Return EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer
        Case "V"
          Return EventBooking.EventBookingStatuses.ebsBookedInvoiced
        Case "O"
          Return EventBooking.EventBookingStatuses.ebsWaitingInvoiced
        Case "D"
          Return EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer
        Case "E"
          Return EventBooking.EventBookingStatuses.ebsExternal
        Case "C"
          Return EventBooking.EventBookingStatuses.ebsCancelled
        Case "I"
          Return EventBooking.EventBookingStatuses.ebsInterested
        Case "T"
          Return EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
        Case "U"
          Return EventBooking.EventBookingStatuses.ebsAmended
      End Select
    End Function

    Public Shared Function GetBookingStatusDescription(ByVal pStatusCode As String) As String
      Dim vStatus As EventBooking.EventBookingStatuses

      vStatus = GetBookingStatus(pStatusCode)
      Select Case vStatus
        Case EventBooking.EventBookingStatuses.ebsAmended
          Return ProjectText.String16491  'Amended
        Case EventBooking.EventBookingStatuses.ebsBooked
          Return ProjectText.String17189  'Booked
        Case EventBooking.EventBookingStatuses.ebsWaiting
          Return ProjectText.String17193  'Waiting
        Case EventBooking.EventBookingStatuses.ebsBookedTransfer
          Return ProjectText.String17191  'Booked Transfer
        Case EventBooking.EventBookingStatuses.ebsBookedAndPaid
          Return ProjectText.String17190  'Booked (Paid)
        Case EventBooking.EventBookingStatuses.ebsWaitingPaid
          Return ProjectText.String17194  'Waiting (Paid)
        Case EventBooking.EventBookingStatuses.ebsBookedAndPaidTransfer
          Return ProjectText.String17192  'Booked (Paid) Transfer
        Case EventBooking.EventBookingStatuses.ebsBookedCreditSale
          Return ProjectText.String16480  'Booked (Credit Sale)
        Case EventBooking.EventBookingStatuses.ebsWaitingCreditSale
          Return ProjectText.String16481  'Waiting (Credit Sale)
        Case EventBooking.EventBookingStatuses.ebsBookedCreditSaleTransfer
          Return ProjectText.String16482  'Booked (Credit Sale) Transfer
        Case EventBooking.EventBookingStatuses.ebsBookedInvoiced
          Return ProjectText.String16483  'Booked (Invoiced)
        Case EventBooking.EventBookingStatuses.ebsWaitingInvoiced
          Return ProjectText.String16484  'Waiting (Invoiced)
        Case EventBooking.EventBookingStatuses.ebsBookedInvoicedTransfer
          Return ProjectText.String16485  'Booked (Invoiced) Transfer
        Case EventBooking.EventBookingStatuses.ebsExternal
          Return ProjectText.String17195  'External Event
        Case EventBooking.EventBookingStatuses.ebsCancelled
          Return ProjectText.String17196  'Cancelled
        Case EventBooking.EventBookingStatuses.ebsInterested
          Return ProjectText.String17202  'Interested
        Case EventBooking.EventBookingStatuses.ebsAwaitingAcceptance
          Return ProjectText.String16486  'Awaiting Acceptance
        Case Else
          Return ProjectText.String17106  'Unknown Booking Status Code
      End Select
    End Function

  End Class

End Namespace