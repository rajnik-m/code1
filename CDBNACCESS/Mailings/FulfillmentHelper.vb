Namespace Access

  Public Class FulfillmentHelper

    'Public Shared Function SetMailingDocumentFulfilled(pEnv As CDBEnvironment, pParams As ParameterList) As Integer
    '  Dim vParams As New CDBParameters
    '  For Each vParm As DictionaryEntry In pParams
    '    vParams.Add(vParm.Key.ToString, New CDBParameter(vParm.Value.ToString))
    '  Next vParm
    '  SetMailingDocumentFulfilled(pEnv, vParams)
    'End Function

    Public Shared Function SetMailingDocumentFulfilled(pEnv As CDBEnvironment, pParams As CDBParameters) As Integer

      Dim vMailingDocumentNumber As Integer = pParams.ParameterExists("MailingDocumentNumber").LongValue

      Dim vFulfillmentHistory As New FulfillmentHistory(pEnv)
      vFulfillmentHistory.Init(pParams("FulfillmentNumber").LongValue)
      If pParams.ParameterExists("Fulfilled").Bool Then
        If vFulfillmentHistory.Existing = False AndAlso pParams("FulfillmentNumber").LongValue > 0 Then vFulfillmentHistory.FulfillmentNumber = pParams("FulfillmentNumber").LongValue
        vFulfillmentHistory.DocumentList = pParams("DocumentList").Value
        If vFulfillmentHistory.NumberOfDocuments = 0 Then vFulfillmentHistory.NumberOfDocuments = pParams("DocumentList").Value.Split(","c).Length
        vFulfillmentHistory.SetFulfilled(pParams.ParameterExists("FulfilmentFilename").Value)
        If pParams.ParameterExists("ConfirmGAD").Bool Then
          ConfirmGAD(pEnv, vFulfillmentHistory.FulfillmentNumber)
        End If

        Dim vTable As String = ""
        Dim vAttr1 As String = ""
        Dim vAttr2 As String = ""
        Dim vWhereFields, vFields As New CDBFields
        vWhereFields.Add("date_fulfilled")
        vFields.Add("date_fulfilled", CDBField.FieldTypes.cftDate, TodaysDate)
        While vTable <> "enclosures"
          Select Case vTable
            Case "contact_incentive_responses"
              vTable = "contact_incentives"
            Case "contact_incentives"
              vTable = "new_orders"
              vWhereFields.Remove(vAttr1)
              vAttr1 = "order_number"
              vAttr2 = ""
            Case "new_orders"
              vTable = "enclosures"
            Case Else
              vTable = "contact_incentive_responses"
              vAttr1 = "contact_number"
              vAttr2 = " AND new_contact = 'Y'"
          End Select
          If vWhereFields.Exists(vAttr1) = False Then vWhereFields.Add(vAttr1, "SELECT " & vAttr1 & " FROM contact_mailing_documents cmd WHERE fulfillment_number= " & vFulfillmentHistory.FulfillmentNumber & vAttr2 & " AND (earliest_fulfilment_date IS NULL OR earliest_fulfilment_date" & pEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDate()) & ")", CDBField.FieldWhereOperators.fwoIn)
          pEnv.Connection.UpdateRecords(vTable, vFields, vWhereFields, False)

        End While
      ElseIf vMailingDocumentNumber > 0 Then
        Dim vContactMailingDocument As New ContactMailingDocument(pEnv)
        vContactMailingDocument.Init(vMailingDocumentNumber)
        If Not vContactMailingDocument.Existing Then RaiseError(DataAccessErrors.daeCMDInvalidDeleted)
        If Not vFulfillmentHistory.Existing Then RaiseError(DataAccessErrors.daeCMDFulfillmentHistoryInvalid)
        If vContactMailingDocument.FulfillmentNumber = 0 Then RaiseError(DataAccessErrors.daeCMDAlreadyUnfulfilled)
        If Not vContactMailingDocument.FulfillmentNumber = vFulfillmentHistory.FulfillmentNumber Then RaiseError(DataAccessErrors.daeParameterValueInvalid, "FulfillmentNumber")
        vFulfillmentHistory.SetAsUnfulfilled("", vMailingDocumentNumber)
      Else
        If vFulfillmentHistory.Existing Then vFulfillmentHistory.SetAsUnfulfilled(pParams("CancellationReason").Value)
      End If
      Return vFulfillmentHistory.FulfillmentNumber
    End Function

    Private Shared Sub ConfirmGAD(ByVal pEnv As CDBEnvironment, ByVal pFulfillmentNumber As Integer)
      Dim vUpdateFields As New CDBFields
      vUpdateFields.Add("confirmed_on", CDBField.FieldTypes.cftDate, TodaysDate)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("method", "O")
      vWhereFields.Add("confirmed_on", CDBField.FieldTypes.cftDate)
      vWhereFields.Add("declaration_number", CDBField.FieldTypes.cftLong, "SELECT declaration_number FROM contact_mailing_documents WHERE fulfillment_number = " & pFulfillmentNumber, CDBField.FieldWhereOperators.fwoIn)
      'Exclude cancelled Declarations
      vWhereFields.Add("cancellation_reason")
      pEnv.Connection.UpdateRecords("gift_aid_declarations", vUpdateFields, vWhereFields, False)
    End Sub

  End Class
End Namespace
