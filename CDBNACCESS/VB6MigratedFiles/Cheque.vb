Option Strict Off
'CAREConverter Option Explicit On
Namespace Access
<System.Runtime.InteropServices.ProgId("Cheque_NET.Cheque")> Public Class Cheque
	
	Public Enum ChequeRecordSetTypes 'These are bit values
		chrtAll = &HFFFFs
		'ADD additional recordset types here
	End Enum
	
	'Keep the enum items in the same order as in the InitClassFields function
	Private Enum ChequeFields
		cfAll = 0
		cfChequeReferenceNumber
		cfContactNumber
		cfAddressNumber
		cfAmount
		cfPrintedOn
		cfChequeNumber
		cfReconciledOn
		cfAmendedBy
		cfAmendedOn
		cfChequeStatus
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
				.DatabaseTableName = "cheques"
				'There should be an entry here for each field in the table
				'Keep these in the same order as the Fields enum
				.Add("cheque_reference_number", CDBField.FieldTypes.cftLong)
				.Add("contact_number", CDBField.FieldTypes.cftLong)
				.Add("address_number", CDBField.FieldTypes.cftLong)
				.Add("amount", CDBField.FieldTypes.cftNumeric)
				.Add("printed_on", CDBField.FieldTypes.cftDate)
				.Add("cheque_number", CDBField.FieldTypes.cftLong)
				.Add("reconciled_on", CDBField.FieldTypes.cftDate)
				.Add("amended_by")
				.Add("amended_on", CDBField.FieldTypes.cftDate)
				.Add("cheque_status")
			End With
			
			mvClassFields.Item(ChequeFields.cfChequeReferenceNumber).SetPrimaryKeyOnly
		Else
			mvClassFields.ClearItems()
		End If
		mvExisting = False
	End Sub
	
	Private Sub SetDefaults()
		'Add code here to initialise the class with default values for a new record
	End Sub
	
	Private Sub SetValid(ByVal pField As ChequeFields)
		'Add code here to ensure all values are valid before saving
		mvClassFields.Item(ChequeFields.cfAmendedOn).Value = TodaysDate()
		mvClassFields.Item(ChequeFields.cfAmendedBy).Value = mvEnv.User.Logname
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mvClassFields may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mvClassFields = Nothing
		'UPGRADE_NOTE: Object mvEnv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mvEnv = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'-----------------------------------------------------------
	' PUBLIC PROCEDURES FOLLOW
	'-----------------------------------------------------------
	Public Function GetRecordSetFields(ByVal pRSType As ChequeRecordSetTypes) As String
		Dim vFields As String
		Dim vClassField As ClassField
		
		'Modify below to add each recordset type as required
		If pRSType = ChequeRecordSetTypes.chrtAll Then
			If mvClassFields Is Nothing Then InitClassFields()
			vFields = mvClassFields.FieldNames(mvEnv, "c")
		Else
			
		End If
		GetRecordSetFields = vFields
	End Function
	
	Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pChequeReferenceNumber As Integer = 0, Optional ByRef pChequeNumber As Integer = 0)
		Dim vRecordSet As CDBRecordSet
		Dim vWhere As String
		
		mvEnv = pEnv
		If pChequeReferenceNumber > 0 Or pChequeNumber > 0 Then
			If pChequeReferenceNumber > 0 Then vWhere = "cheque_reference_number = " & pChequeReferenceNumber
			If pChequeNumber > 0 Then
				If Len(vWhere) > 0 Then vWhere = vWhere & " AND "
				vWhere = vWhere & "cheque_number = " & pChequeNumber
			End If
			vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ChequeRecordSetTypes.chrtAll) & " FROM cheques c WHERE " & vWhere)
			If vRecordSet.Fetch() = True Then
				InitFromRecordSet(pEnv, vRecordSet, ChequeRecordSetTypes.chrtAll)
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
	
	Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ChequeRecordSetTypes)
		Dim vFields As CDBFields
		
		mvEnv = pEnv
		InitClassFields()
		vFields = pRecordSet.Fields
		mvExisting = True
		With mvClassFields
			'Always include the primary key attributes
			.SetItem(ChequeFields.cfChequeReferenceNumber, vFields)
			'Modify below to handle each recordset type as required
			If (pRSType And ChequeRecordSetTypes.chrtAll) = ChequeRecordSetTypes.chrtAll Then
				.SetItem(ChequeFields.cfContactNumber, vFields)
				.SetItem(ChequeFields.cfAddressNumber, vFields)
				.SetItem(ChequeFields.cfAmount, vFields)
				.SetItem(ChequeFields.cfPrintedOn, vFields)
				.SetItem(ChequeFields.cfChequeNumber, vFields)
				.SetItem(ChequeFields.cfReconciledOn, vFields)
				.SetItem(ChequeFields.cfAmendedBy, vFields)
				.SetItem(ChequeFields.cfAmendedOn, vFields)
				.SetItem(ChequeFields.cfChequeStatus, vFields)
			End If
		End With
	End Sub
	
	Public Sub Reconcile(ByVal pReconciledOn As String, ByVal pChequeStatus As String)
		mvClassFields(ChequeFields.cfReconciledOn).Value = pReconciledOn
		mvClassFields(ChequeFields.cfChequeStatus).Value = pChequeStatus
	End Sub
	
	Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
		SetValid(ChequeFields.cfAll)
		mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
	End Sub
	
	'-----------------------------------------------------------
	' PROPERTY PROCEDURES FOLLOW
	'-----------------------------------------------------------
	Public ReadOnly Property Existing() As Boolean
		Get
			Existing = mvExisting
		End Get
	End Property
	
	Public ReadOnly Property AddressNumber() As Integer
		Get
			AddressNumber = mvClassFields.Item(ChequeFields.cfAddressNumber).IntegerValue
		End Get
	End Property
	
	Public ReadOnly Property AmendedBy() As String
		Get
			AmendedBy = mvClassFields.Item(ChequeFields.cfAmendedBy).Value
		End Get
	End Property
	
	Public ReadOnly Property AmendedOn() As String
		Get
			AmendedOn = mvClassFields.Item(ChequeFields.cfAmendedOn).Value
		End Get
	End Property
	
	Public ReadOnly Property Amount() As Double
		Get
			Amount = mvClassFields.Item(ChequeFields.cfAmount).DoubleValue
		End Get
	End Property
	
	Public ReadOnly Property ChequeNumber() As String
		Get
			ChequeNumber = mvClassFields.Item(ChequeFields.cfChequeNumber).Value
		End Get
	End Property
	
	Public ReadOnly Property ChequeReferenceNumber() As String
		Get
			ChequeReferenceNumber = mvClassFields.Item(ChequeFields.cfChequeReferenceNumber).Value
		End Get
	End Property
	
	Public ReadOnly Property ChequeStatus() As String
		Get
			ChequeStatus = mvClassFields.Item(ChequeFields.cfChequeStatus).Value
		End Get
	End Property
	
	Public ReadOnly Property ContactNumber() As Integer
		Get
			ContactNumber = mvClassFields.Item(ChequeFields.cfContactNumber).IntegerValue
		End Get
	End Property
	
	Public ReadOnly Property PrintedOn() As String
		Get
			PrintedOn = mvClassFields.Item(ChequeFields.cfPrintedOn).Value
		End Get
	End Property
	
	Public ReadOnly Property ReconciledOn() As String
		Get
			ReconciledOn = mvClassFields.Item(ChequeFields.cfReconciledOn).Value
		End Get
	End Property
End Class
End Namespace
