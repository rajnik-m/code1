Namespace Access
  Public Class AppealResource

    Public Enum AppealResourceRecordSetTypes 'These are bit values
      arrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum AppealResourceFields
      arfAll = 0
      arfAppealResourceNumber
      arfCampaign
      arfAppeal
      arfProduct
      arfTotalQuantity
      arfQuantityRemaining
      arfAmendedBy
      arfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "appeal_resources"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("appeal_resource_number", CDBField.FieldTypes.cftLong)
          .Add("campaign")
          .Add("appeal")
          .Add("product")
          .Add("total_quantity", CDBField.FieldTypes.cftLong)
          .Add("quantity_remaining", CDBField.FieldTypes.cftLong)
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(AppealResourceFields.arfAppealResourceNumber).SetPrimaryKeyOnly()

          .SetUniqueField(AppealResourceFields.arfCampaign)
          .SetUniqueField(AppealResourceFields.arfAppeal)
          .SetUniqueField(AppealResourceFields.arfProduct)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As AppealResourceFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(AppealResourceFields.arfAppealResourceNumber).IntegerValue = 0 Then mvClassFields.Item(AppealResourceFields.arfAppealResourceNumber).IntegerValue = mvEnv.GetControlNumber("AS")
      mvClassFields.Item(AppealResourceFields.arfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(AppealResourceFields.arfAmendedBy).Value = mvEnv.User.Logname

      If mvClassFields(AppealResourceFields.arfTotalQuantity).Value <> mvClassFields(AppealResourceFields.arfTotalQuantity).SetValue Then
        'the quantity has changed, the remaining quantity needs to be updated.
        IssueQuantity(IntegerValue(mvClassFields(AppealResourceFields.arfTotalQuantity).SetValue) - IntegerValue(mvClassFields(AppealResourceFields.arfTotalQuantity).Value))
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As AppealResourceRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = AppealResourceRecordSetTypes.arrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "ar")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pAppealResourceNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pAppealResourceNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AppealResourceRecordSetTypes.arrtAll) & " FROM appeal_resources ar WHERE appeal_resource_number = " & pAppealResourceNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, AppealResourceRecordSetTypes.arrtAll)
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

    Public Sub InitFromProduct(ByVal pEnv As CDBEnvironment, ByRef pCampaign As String, ByRef pAppeal As String, ByRef pProduct As String)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      With vWhereFields
        .Add("campaign", CDBField.FieldTypes.cftCharacter, pCampaign)
        .Add("appeal", CDBField.FieldTypes.cftCharacter, pAppeal)
        .Add("product", CDBField.FieldTypes.cftCharacter, pProduct)
      End With
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(AppealResourceRecordSetTypes.arrtAll) & " FROM appeal_resources ar WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, AppealResourceRecordSetTypes.arrtAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AppealResourceRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AppealResourceFields.arfAppealResourceNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AppealResourceRecordSetTypes.arrtAll) = AppealResourceRecordSetTypes.arrtAll Then
          .SetItem(AppealResourceFields.arfCampaign, vFields)
          .SetItem(AppealResourceFields.arfAppeal, vFields)
          .SetItem(AppealResourceFields.arfProduct, vFields)
          .SetItem(AppealResourceFields.arfTotalQuantity, vFields)
          .SetItem(AppealResourceFields.arfQuantityRemaining, vFields)
          .SetItem(AppealResourceFields.arfAmendedBy, vFields)
          .SetItem(AppealResourceFields.arfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(AppealResourceFields.arfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByVal pCampaign As String, ByVal pAppeal As String, ByVal pProduct As String, ByVal pTotalQuantity As Integer)
      Init(pEnv)
      With mvClassFields
        .Item(AppealResourceFields.arfCampaign).Value = pCampaign
        .Item(AppealResourceFields.arfAppeal).Value = pAppeal
        .Item(AppealResourceFields.arfProduct).Value = pProduct
        .Item(AppealResourceFields.arfTotalQuantity).Value = CStr(pTotalQuantity)
      End With
    End Sub

    Public Sub Update(ByVal pTotalQuantity As Integer)
      With mvClassFields
        .Item(AppealResourceFields.arfTotalQuantity).Value = CStr(pTotalQuantity)
      End With
    End Sub

    Public Sub Delete()
      If DeleteAllowed() Then
        mvEnv.Connection.DeleteRecords("appeal_resources", mvClassFields.WhereFields)
      End If
    End Sub

    Private Function DeleteAllowed() As Boolean
      Dim vWhereFields As New CDBFields
      Dim vDeleteAllowed As Boolean

      vDeleteAllowed = True
      vWhereFields.Add("appeal_resource_number", CDBField.FieldTypes.cftLong, AppealResourceNumber)
      If mvEnv.Connection.GetCount("collection_resources", vWhereFields) > 0 Then
        RaiseError(DataAccessErrors.daeCannotDeleteAppResource)
      End If
      DeleteAllowed = vDeleteAllowed
    End Function

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(AppealResourceFields.arfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(AppealResourceFields.arfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Appeal() As String
      Get
        Appeal = mvClassFields.Item(AppealResourceFields.arfAppeal).Value
      End Get
    End Property

    Public ReadOnly Property AppealResourceNumber() As Integer
      Get
        AppealResourceNumber = mvClassFields.Item(AppealResourceFields.arfAppealResourceNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Campaign() As String
      Get
        Campaign = mvClassFields.Item(AppealResourceFields.arfCampaign).Value
      End Get
    End Property

    Public ReadOnly Property Product() As String
      Get
        Product = mvClassFields.Item(AppealResourceFields.arfProduct).Value
      End Get
    End Property

    Public ReadOnly Property QuantityRemaining() As Integer
      Get
        QuantityRemaining = mvClassFields.Item(AppealResourceFields.arfQuantityRemaining).IntegerValue
      End Get
    End Property

    Public ReadOnly Property TotalQuantity() As Integer
      Get
        TotalQuantity = mvClassFields.Item(AppealResourceFields.arfTotalQuantity).IntegerValue
      End Get
    End Property

    Public Sub IssueQuantity(ByRef pQuantity As Integer)
      Dim vTemp As Integer
      vTemp = QuantityRemaining - pQuantity
      If vTemp < 0 Then
        RaiseError(DataAccessErrors.daeLowStockOnAppeal, CStr(QuantityRemaining))
      Else
        mvClassFields(AppealResourceFields.arfQuantityRemaining).IntegerValue = vTemp
      End If
    End Sub
  End Class
End Namespace
