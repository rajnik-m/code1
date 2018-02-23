

Namespace Access
  Public Class CollectionBox

    Public Enum CollectionBoxRecordSetTypes 'These are bit values
      cbrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectionBoxFields
      cbfAll = 0
      cbfCollectionBoxNumber
      cbfCollectionNumber
      cbfBoxReference
      cbfCollectorNumber
      cbfAmount
      cbfCollectionPisNumber
      cbfAmendedBy
      cbfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvCollectorContactNumber As Integer
    Private mvCollectorAddressNumber As Integer

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "collection_boxes"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collection_box_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("box_reference")
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("amount", CDBField.FieldTypes.cftNumeric)
          .Add("collection_pis_number", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CollectionBoxFields.cbfCollectionBoxNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(CollectionBoxFields.cbfCollectionNumber).PrefixRequired = True
        mvClassFields.Item(CollectionBoxFields.cbfCollectorNumber).PrefixRequired = True
        mvClassFields.Item(CollectionBoxFields.cbfAmendedBy).PrefixRequired = True
        mvClassFields.Item(CollectionBoxFields.cbfAmendedOn).PrefixRequired = True
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectionBoxFields)
      'Add code here to ensure all values are valid before saving
      With mvClassFields
        If .Item(CollectionBoxFields.cbfCollectionBoxNumber).IntegerValue = 0 Then .Item(CollectionBoxFields.cbfCollectionBoxNumber).IntegerValue = mvEnv.GetControlNumber("CB")
        If .Item(CollectionBoxFields.cbfCollectorNumber).ValueChanged OrElse .Item(CollectionBoxFields.cbfCollectionPisNumber).ValueChanged OrElse .Item(CollectionBoxFields.cbfAmount).ValueChanged Then
          If HasPayments Then RaiseError(DataAccessErrors.daeCannotChangeBoxPaymentsMade)
        End If
        .Item(CollectionBoxFields.cbfAmendedOn).Value = TodaysDate()
        .Item(CollectionBoxFields.cbfAmendedBy).Value = mvEnv.User.Logname
      End With
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectionBoxRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CollectionBoxRecordSetTypes.cbrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cb")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionBoxNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionBoxNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectionBoxRecordSetTypes.cbrtAll) & " FROM collection_boxes cb WHERE collection_box_number = " & pCollectionBoxNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectionBoxRecordSetTypes.cbrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectionBoxRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectionBoxFields.cbfCollectionBoxNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectionBoxRecordSetTypes.cbrtAll) = CollectionBoxRecordSetTypes.cbrtAll Then
          .SetItem(CollectionBoxFields.cbfCollectionNumber, vFields)
          .SetItem(CollectionBoxFields.cbfBoxReference, vFields)
          .SetItem(CollectionBoxFields.cbfCollectorNumber, vFields)
          .SetItem(CollectionBoxFields.cbfAmount, vFields)
          .SetItem(CollectionBoxFields.cbfCollectionPisNumber, vFields)
          .SetItem(CollectionBoxFields.cbfAmendedBy, vFields)
          .SetItem(CollectionBoxFields.cbfAmendedOn, vFields)
        End If
        mvCollectorContactNumber = pRecordSet.Fields.FieldExists("contact_number").IntegerValue
        mvCollectorAddressNumber = pRecordSet.Fields.FieldExists("address_number").IntegerValue
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectionBoxFields.cbfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    'Public Sub Update(ByVal pCollectionNumber As Long, ByVal pBoxReference As String, ByVal pCollectionShiftNumber As Long, ByVal pCollectorNumber As Long, ByVal pAmount As Double)
    '  With mvClassFields
    '    .Item(cbfCollectionNumber).Value = pCollectionNumber
    '    .Item(cbfBoxReference).Value = pBoxReference
    '    .Item(cbfCollectorNumber).IntegerValue = pCollectorNumber
    '    .Item(cbfAmount).DoubleValue = pAmount
    '    .Item(cbfCollectorNumber) = pCollectorNumber
    '  End With
    'End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      If DeleteAllowed() Then
        mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      End If
    End Sub

    Private Function DeleteAllowed() As Boolean
      Dim vDeleteAllowed As Boolean

      vDeleteAllowed = True
      If HasPayments(True) Then
        RaiseError(DataAccessErrors.daeCannotDeleteCollBoxAsPayments)
      End If
      DeleteAllowed = vDeleteAllowed
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------

    Private ReadOnly Property HasPayments(ByVal pDelete As Boolean) As Boolean
      Get
        Dim vDS As New VBDataSelection
        Dim vDT As New CDBDataTable
        Dim vParams As New CDBParameters
        Dim vHasPayments As Boolean = False

        With mvClassFields
          If pDelete Then
            If CollectionPisNumber > 0 Then
              Dim vWhereFields As New CDBFields(New CDBField("collection_pis_number", CollectionPisNumber))
              vHasPayments = mvEnv.Connection.GetCount("collection_payments", New CDBFields(New CDBField("collection_pis_number", CollectionPisNumber))) > 0
            End If
          ElseIf .Item(CollectionBoxFields.cbfCollectorNumber).ValueChanged OrElse .Item(CollectionBoxFields.cbfCollectionPisNumber).ValueChanged OrElse .Item(CollectionBoxFields.cbfAmount).ValueChanged Then
            'BR14157: If the Box Collector is currently not set then always allow it to be set (as far as I know, the Collector must be set in order for payments to have been allocated to the Box).
            'If the Box Collector is set & is changed, allow it to be changed if there are no records in the collection_payments table for the Box.
            'If the PIS is currently not set & is being set, allow it to be changed if there are no records in the collection_payments table for the Box.
            'If the amount has changed, allow it to be changed if there are no records in the collection_payments table for the Box.
            'This will then allow a collector to be allocated to a Box even if that collector has already banked payments for a different Box in the Collection.
            Dim vWhereFields As New CDBFields(New CDBField("collection_box_number", CollectionBoxNumber))
            If .Item(CollectionBoxFields.cbfCollectionPisNumber).SetValue.Length > 0 AndAlso .Item(CollectionBoxFields.cbfCollectionPisNumber).ValueChanged Then
              'If the PIS is set & is changed, allow it to be changed if there are no records in the collection_payments table for the Box and original PIS number.
              vWhereFields.Add("collection_pis_number", CDBField.FieldTypes.cftInteger, .Item(CollectionBoxFields.cbfCollectionPisNumber).SetValue)
            End If
            vHasPayments = mvEnv.Connection.GetCount("collection_payments", vWhereFields) > 0
          End If

          If .Item(CollectionBoxFields.cbfCollectorNumber).SetValue.Length > 0 And Not vHasPayments Then
            vParams.Add("CollectionNumber", CollectionNumber)
            vParams.Add("ContactNumber", IntegerValue(mvEnv.Connection.GetValue("SELECT contact_number FROM manned_collectors WHERE collector_number = " & .Item(CollectionBoxFields.cbfCollectorNumber).SetValue)))
            vDS.Init(mvEnv, DataSelection.DataSelectionTypes.dstContactCollectionPayments, vParams)
            vDT = vDS.DataTable()
            vHasPayments = vDT.Rows.Count() > 0
          End If
        End With
        Return vHasPayments
      End Get
    End Property

    Private ReadOnly Property HasPayments() As Boolean
      Get
        Return HasPayments(False)
      End Get
    End Property

    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(CollectionBoxFields.cbfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectionBoxFields.cbfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Amount() As String
      Get
        Amount = mvClassFields.Item(CollectionBoxFields.cbfAmount).Value
      End Get
    End Property

    Public ReadOnly Property BoxReference() As String
      Get
        BoxReference = mvClassFields.Item(CollectionBoxFields.cbfBoxReference).Value
      End Get
    End Property

    Public ReadOnly Property CollectionBoxNumber() As Integer
      Get
        CollectionBoxNumber = mvClassFields.Item(CollectionBoxFields.cbfCollectionBoxNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(CollectionBoxFields.cbfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        CollectorNumber = mvClassFields.Item(CollectionBoxFields.cbfCollectorNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectionPisNumber() As Integer
      Get
        CollectionPisNumber = mvClassFields.Item(CollectionBoxFields.cbfCollectionPisNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorContactNumber() As Integer
      Get
        CollectorContactNumber = mvCollectorContactNumber
      End Get
    End Property

    Public ReadOnly Property CollectorAddressNumber() As Integer
      Get
        CollectorAddressNumber = mvCollectorAddressNumber
      End Get
    End Property

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(CollectionBoxFields.cbfCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(CollectionBoxFields.cbfBoxReference).Value = pParams("BoxReference").Value
        If pParams.Exists("CollectionPisNumber") Then .Item(CollectionBoxFields.cbfCollectionPisNumber).Value = pParams("CollectionPisNumber").Value
        If pParams.Exists("CollectorNumber") Then .Item(CollectionBoxFields.cbfCollectorNumber).Value = pParams("CollectorNumber").Value
        If pParams.Exists("Amount") Then .Item(CollectionBoxFields.cbfAmount).Value = pParams("Amount").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("CollectorNumber") Then .Item(CollectionBoxFields.cbfCollectorNumber).Value = pParams("CollectorNumber").Value
        If pParams.Exists("Amount") Then .Item(CollectionBoxFields.cbfAmount).Value = pParams("Amount").Value
        If pParams.Exists("CollectionPisNumber") Then .Item(CollectionBoxFields.cbfCollectionPisNumber).Value = pParams("CollectionPisNumber").Value
      End With
    End Sub
  End Class
End Namespace
