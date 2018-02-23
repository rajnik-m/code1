Option Strict Off
'CAREConverter Option Explicit On
Namespace Access
<System.Runtime.InteropServices.ProgId("EventContact_NET.EventContact")> Public Class EventContact
	
	Public Enum EventContactRecordSetTypes 'These are bit values
		ecrtAll = &HFFFFs
		'ADD additional recordset types here
	End Enum
	
	'Keep the enum items in the same order as in the InitClassFields function
	Private Enum EventContactFields
		ecfAll = 0
		ecfEventNumber
		ecfContactNumber
		ecfAddressNumber
		ecfEventContactRelationship
		ecfNotes
		ecfAmendedBy
		ecfAmendedOn
	End Enum
	
	Private mvContact As Contact
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
				.DatabaseTableName = "event_contacts"
				'There should be an entry here for each field in the table
				'Keep these in the same order as the Fields enum
				.Add("event_number", CDBField.FieldTypes.cftLong)
				.Add("contact_number", CDBField.FieldTypes.cftLong)
				.Add("address_number", CDBField.FieldTypes.cftLong)
				.Add("event_contact_relationship")
				.Add("notes", CDBField.FieldTypes.cftMemo)
				.Add("amended_by")
				.Add("amended_on", CDBField.FieldTypes.cftDate)
			End With
			
			mvClassFields.Item(EventContactFields.ecfEventNumber).PrimaryKey = True
			mvClassFields.Item(EventContactFields.ecfContactNumber).PrimaryKey = True
		Else
			mvClassFields.ClearItems()
		End If
		mvExisting = False
	End Sub
	
	Private Sub SetDefaults()
		'Add code here to initialise the class with default values for a new record
	End Sub
	
	Private Sub SetValid(ByVal pField As EventContactFields)
		'Add code here to ensure all values are valid before saving
		mvClassFields.Item(EventContactFields.ecfAmendedOn).Value = TodaysDate()
		mvClassFields.Item(EventContactFields.ecfAmendedBy).Value = mvEnv.User.Logname
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
	Public Function GetRecordSetFields(ByVal pRSType As EventContactRecordSetTypes) As String
		Dim vFields As String
		Dim vClassField As ClassField
		
		'Modify below to add each recordset type as required
		If pRSType = EventContactRecordSetTypes.ecrtAll Then
			If mvClassFields Is Nothing Then InitClassFields()
			vFields = mvClassFields.FieldNames(mvEnv, "ec")
		Else
			
		End If
		GetRecordSetFields = vFields
	End Function
	Public ReadOnly Property Contact() As Contact
		Get
			If mvContact Is Nothing Then
				If ContactNumber > 0 Then
					mvContact = New Contact(mvEnv)
					mvContact.Init( ContactNumber)
				End If
			End If
			Contact = mvContact
		End Get
	End Property
	
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
			AmendedBy = mvClassFields.Item(EventContactFields.ecfAmendedBy).Value
		End Get
	End Property
	
	Public ReadOnly Property AmendedOn() As String
		Get
			AmendedOn = mvClassFields.Item(EventContactFields.ecfAmendedOn).Value
		End Get
	End Property
	
	Public ReadOnly Property ContactNumber() As Integer
		Get
			ContactNumber = mvClassFields.Item(EventContactFields.ecfContactNumber).IntegerValue
		End Get
	End Property
	
	Public Property EventNumber() As Integer
		Get
			EventNumber = mvClassFields.Item(EventContactFields.ecfEventNumber).IntegerValue
		End Get
		Set(ByVal Value As Integer)
			mvClassFields.Item(EventContactFields.ecfEventNumber).IntegerValue = Value
		End Set
	End Property
	
'CAREConverter 	Public Property FormValue(ByVal pAttributeName As String) As String
'CAREConverter 		Get
'CAREConverter 			FormValue = mvClassFields.Item(pAttributeName).FormValue
'CAREConverter 		End Get
'CAREConverter 		Set(ByVal Value As String)
'CAREConverter 			mvClassFields.Item(pAttributeName).Value = Value
'CAREConverter 		End Set
'CAREConverter 	End Property
	
	Public ReadOnly Property Notes() As String
		Get
			Notes = mvClassFields.Item(EventContactFields.ecfNotes).MultiLineValue
		End Get
	End Property
	
	Public ReadOnly Property EventContactRelationship() As String
		Get
			EventContactRelationship = mvClassFields.Item(EventContactFields.ecfEventContactRelationship).Value
		End Get
	End Property
	Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
		mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
	End Sub
	
	Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pEventNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0)
		Dim vRecordSet As CDBRecordSet
		
		mvEnv = pEnv
		If pEventNumber > 0 Then
			vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(EventContactRecordSetTypes.ecrtAll) & " FROM event_contacts ec WHERE event_number = " & pEventNumber & " AND contact_number = " & pContactNumber)
			If vRecordSet.Fetch() = True Then
				InitFromRecordSet(pEnv, vRecordSet, EventContactRecordSetTypes.ecrtAll)
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
	
	Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As EventContactRecordSetTypes)
		Dim vFields As CDBFields
		
		mvEnv = pEnv
		InitClassFields()
		vFields = pRecordSet.Fields
		mvExisting = True
		With mvClassFields
			'Always include the primary key attributes
			.SetItem(EventContactFields.ecfEventNumber, vFields)
			.SetItem(EventContactFields.ecfContactNumber, vFields)
			'Modify below to handle each recordset type as required
			If (pRSType And EventContactRecordSetTypes.ecrtAll) = EventContactRecordSetTypes.ecrtAll Then
				.SetItem(EventContactFields.ecfAddressNumber, vFields)
				.SetItem(EventContactFields.ecfEventContactRelationship, vFields)
				.SetItem(EventContactFields.ecfNotes, vFields)
				.SetItem(EventContactFields.ecfAmendedBy, vFields)
				.SetItem(EventContactFields.ecfAmendedOn, vFields)
			End If
		End With
	End Sub
	
	Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAHC As ClassFields.AmendmentHistoryCreation = ClassFields.AmendmentHistoryCreation.ahcDefault)
		SetValid(EventContactFields.ecfAll)
		mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, mvClassFields.CreateAmendmentHistory(mvEnv, pAHC))
	End Sub
	
	Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
		'Auto Generated code for WEB services
		Init(pEnv)
		With mvClassFields
			.Item(EventContactFields.ecfEventNumber).Value = pParams("EventNumber").Value
			.Item(EventContactFields.ecfContactNumber).Value = pParams("ContactNumber").Value
		End With
		Update(pParams)
	End Sub
	
	Public Sub Update(ByRef pParams As CDBParameters)
		'Auto Generated code for WEB services
		With mvClassFields
			If pParams.Exists("AddressNumber") Then .Item(EventContactFields.ecfAddressNumber).Value = pParams("AddressNumber").Value
			If pParams.Exists("EventContactRelationship") Then .Item(EventContactFields.ecfEventContactRelationship).Value = pParams("EventContactRelationship").Value
			If pParams.Exists("Notes") Then .Item(EventContactFields.ecfNotes).Value = pParams("Notes").Value
		End With
	End Sub
End Class
End Namespace
