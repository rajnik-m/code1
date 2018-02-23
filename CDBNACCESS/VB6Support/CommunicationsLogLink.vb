Namespace Access

  Partial Public Class CommunicationsLogLink

    Public Shared Function GetLinkTypeDescription(ByVal pLinkType As String) As String
      Dim vLinkType As CommunicationsLogLink.CommunicationLogLinkTypes

      vLinkType = CommunicationsLogLink.GetLinkType(pLinkType)
      Select Case vLinkType
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee 'A
          Return ProjectText.String15803 'Addressee
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied 'C
          Return ProjectText.String15805 'Copied To
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed 'D
          Return ProjectText.String15806 'Distributed To
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated 'R
          Return ProjectText.String15807 'Related To
        Case Else   'CommunicationsLogLink.CommunicationLogLinkTypes.clltSender 'S
          Return ProjectText.String15804 'Sender
      End Select
    End Function

    Public Shared Function GetLinkType(ByRef pLinkType As String) As CommunicationsLogLink.CommunicationLogLinkTypes
      Select Case pLinkType
        Case "A"
          Return CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee
        Case "C"
          Return CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied
        Case "D"
          Return CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed
        Case "R"
          Return CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated
        Case "S"
          Return CommunicationsLogLink.CommunicationLogLinkTypes.clltSender
      End Select
    End Function

  End Class

End Namespace
