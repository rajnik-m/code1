

Namespace Access
  Public Class ServiceControl

    Public Enum ServiceControlRecordSetTypes 'These are bit values
      svcrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum ServiceControlFields
      scfAll = 0
      scfContactGroup
      scfModifierGroup
      scfModifierRelationship
      scfModifierActivity
      scfModifierActivityValue
      scfAmendedBy
      scfAmendedOn
      scfDefaultStartTime
      scfDefaultEndTime
      scfFinderType
      scfAppointmentType
      scfLateBookingNotificationDays
      scfGeographicRegionType
      scfRequiresStartDays
    End Enum

    Public Enum ServiceControlFinderTypes
      scftContact
      scftServiceProduct
    End Enum

    Public Enum ServiceControlAppointmentTypes
      scatTimeBased
      scatOvernight
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
          .DatabaseTableName = "service_controls"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("contact_group")
          .Add("modifier_group")
          .Add("modifier_relationship")
          .Add("modifier_activity")
          .Add("modifier_activity_value")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
          .Add("default_start_time")
          .Add("default_end_time")
          .Add("finder_type")
          .Add("appointment_type")
          .Add("late_booking_notification")
          .Add("geographical_region_type")
          .Add("requires_start_days")

          .Item(ServiceControlFields.scfContactGroup).SetPrimaryKeyOnly()
          .Item(ServiceControlFields.scfLateBookingNotificationDays).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets)
          .Item(ServiceControlFields.scfGeographicRegionType).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHolidayLets)
          .Item(ServiceControlFields.scfRequiresStartDays).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataServiceStartDays)
          .Item(ServiceControlFields.scfContactGroup).PrefixRequired = True
          .Item(ServiceControlFields.scfAmendedBy).PrefixRequired = True
          .Item(ServiceControlFields.scfAmendedOn).PrefixRequired = True
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As ServiceControlFields)
      'Add code here to ensure all values are valid before saving
      mvClassFields.Item(ServiceControlFields.scfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(ServiceControlFields.scfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As ServiceControlRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = ServiceControlRecordSetTypes.svcrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pContactGroup As String = "")
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If Len(pContactGroup) > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(ServiceControlRecordSetTypes.svcrtAll) & " FROM service_controls sc WHERE contact_group = '" & pContactGroup & "'")
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, ServiceControlRecordSetTypes.svcrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ServiceControlRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(ServiceControlFields.scfContactGroup, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And ServiceControlRecordSetTypes.svcrtAll) = ServiceControlRecordSetTypes.svcrtAll Then
          .SetItem(ServiceControlFields.scfModifierGroup, vFields)
          .SetItem(ServiceControlFields.scfModifierRelationship, vFields)
          .SetItem(ServiceControlFields.scfModifierActivity, vFields)
          .SetItem(ServiceControlFields.scfModifierActivityValue, vFields)
          .SetItem(ServiceControlFields.scfAmendedBy, vFields)
          .SetItem(ServiceControlFields.scfAmendedOn, vFields)
          .SetOptionalItem(ServiceControlFields.scfDefaultStartTime, vFields)
          .SetOptionalItem(ServiceControlFields.scfDefaultEndTime, vFields)
          .SetOptionalItem(ServiceControlFields.scfFinderType, vFields)
          .SetOptionalItem(ServiceControlFields.scfAppointmentType, vFields)
          .SetOptionalItem(ServiceControlFields.scfLateBookingNotificationDays, vFields)
          .SetOptionalItem(ServiceControlFields.scfGeographicRegionType, vFields)
          .SetOptionalItem(ServiceControlFields.scfRequiresStartDays, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(ServiceControlFields.scfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Function GetModifierQuantity(ByVal pModifierContactNumber As Integer) As String
      Dim vCategory As New ContactCategory(mvEnv)

      vCategory.Init(pModifierContactNumber, ModifierActivity, ModifierActivityValue)
      If vCategory.Existing Then
        GetModifierQuantity = vCategory.Quantity
      Else
        GetModifierQuantity = "1"
      End If
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
        AmendedBy = mvClassFields.Item(ServiceControlFields.scfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(ServiceControlFields.scfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property AppointmentType() As ServiceControlAppointmentTypes
      Get
        Select Case mvClassFields.Item(ServiceControlFields.scfAppointmentType).Value
          Case "N"
            AppointmentType = ServiceControlAppointmentTypes.scatOvernight
          Case Else
            AppointmentType = ServiceControlAppointmentTypes.scatTimeBased
        End Select
      End Get
    End Property

    Public ReadOnly Property ContactGroupCode() As String
      Get
        ContactGroupCode = mvClassFields.Item(ServiceControlFields.scfContactGroup).Value
      End Get
    End Property

    Public ReadOnly Property DefaultEndTime() As String
      Get
        Dim vTime As String = ""
        If Len(mvClassFields.Item(ServiceControlFields.scfDefaultEndTime).Value) > 0 Then
          vTime = mvClassFields.Item(ServiceControlFields.scfDefaultEndTime).Value
          vTime = Left(vTime, 2) & ":" & Right(vTime, 2)
        End If
        DefaultEndTime = vTime
      End Get
    End Property

    Public ReadOnly Property DefaultStartTime() As String
      Get
        Dim vTime As String = ""
        If Len(mvClassFields.Item(ServiceControlFields.scfDefaultStartTime).Value) > 0 Then
          vTime = mvClassFields.Item(ServiceControlFields.scfDefaultStartTime).Value
          vTime = Left(vTime, 2) & ":" & Right(vTime, 2)
        End If
        DefaultStartTime = vTime
      End Get
    End Property

    Public ReadOnly Property FinderType() As ServiceControlFinderTypes
      Get
        Select Case mvClassFields.Item(ServiceControlFields.scfFinderType).Value
          Case "S"
            FinderType = ServiceControlFinderTypes.scftServiceProduct
          Case Else
            FinderType = ServiceControlFinderTypes.scftContact
        End Select
      End Get
    End Property

    Public ReadOnly Property ModifierActivity() As String
      Get
        ModifierActivity = mvClassFields.Item(ServiceControlFields.scfModifierActivity).Value
      End Get
    End Property

    Public ReadOnly Property ModifierActivityValue() As String
      Get
        ModifierActivityValue = mvClassFields.Item(ServiceControlFields.scfModifierActivityValue).Value
      End Get
    End Property

    Public ReadOnly Property ModifierGroupCode() As String
      Get
        ModifierGroupCode = mvClassFields.Item(ServiceControlFields.scfModifierGroup).Value
      End Get
    End Property

    Public ReadOnly Property ModifierRelationship() As String
      Get
        ModifierRelationship = mvClassFields.Item(ServiceControlFields.scfModifierRelationship).Value
      End Get
    End Property

    Public ReadOnly Property LateBookingNotificationDays() As Integer
      Get
        LateBookingNotificationDays = mvClassFields.Item(ServiceControlFields.scfLateBookingNotificationDays).IntegerValue
      End Get
    End Property

    Public ReadOnly Property GeographicRegionType() As String
      Get
        GeographicRegionType = mvClassFields.Item(ServiceControlFields.scfGeographicRegionType).Value
      End Get
    End Property

    Public ReadOnly Property RequiresStartDays() As Boolean
      Get
        RequiresStartDays = mvClassFields.Item(ServiceControlFields.scfRequiresStartDays).Bool
      End Get
    End Property
  End Class
End Namespace
