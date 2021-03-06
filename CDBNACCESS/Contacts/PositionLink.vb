Namespace Access

  Public Class PositionLink
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum PositionLinkFields
      AllFields = 0
      ContactPositionNumber
      LinkedContactNumber
      Relationship
      ValidFrom
      ValidTo
      HistoryOnly
      Notes
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_position_number", CDBField.FieldTypes.cftLong)
        .Add("linked_contact_number", CDBField.FieldTypes.cftLong)
        .Add("relationship")
        .Add("valid_from", CDBField.FieldTypes.cftDate)
        .Add("valid_to", CDBField.FieldTypes.cftDate)
        .Add("history_only")
        .Add("notes", CDBField.FieldTypes.cftMemo)

        .Item(PositionLinkFields.ContactPositionNumber).PrimaryKey = True
        .Item(PositionLinkFields.LinkedContactNumber).PrimaryKey = True
        .Item(PositionLinkFields.Relationship).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cpl"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_position_links"
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
    Public ReadOnly Property ContactPositionNumber() As Integer
      Get
        Return mvClassFields(PositionLinkFields.ContactPositionNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property LinkedContactNumber() As Integer
      Get
        Return mvClassFields(PositionLinkFields.LinkedContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Relationship() As String
      Get
        Return mvClassFields(PositionLinkFields.Relationship).Value
      End Get
    End Property
    Public ReadOnly Property ValidFrom() As String
      Get
        Return mvClassFields(PositionLinkFields.ValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property ValidTo() As String
      Get
        Return mvClassFields(PositionLinkFields.ValidTo).Value
      End Get
    End Property
    Public ReadOnly Property HistoryOnly() As String
      Get
        Return mvClassFields(PositionLinkFields.HistoryOnly).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(PositionLinkFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(PositionLinkFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(PositionLinkFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non AutoGenerated Code"

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields.Item(PositionLinkFields.HistoryOnly).Bool = False
    End Sub

    Public Overrides Function GetAddRecordMandatoryParameters() As String
      Return "ContactPositionNumber,LinkedContactNumber,Relationship"
    End Function

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateCreateParameters(pParameterList)
      'Validate Contact exists
      Dim vContact As New Contact(mvEnv)
      vContact.Init(pParameterList("LinkedContactNumber").IntegerValue)
      If vContact.Existing = False Then Throw New CareException("Contact does not exist", 1, "")
      'Validate ContactPositionNumber / Relationship does not already exist
      Dim vFields As New CDBFields(New CDBField("contact_position_number", pParameterList("ContactPositionNumber").IntegerValue))
      vFields.Add("relationship", pParameterList("Relationship").Value)
      If RecordExists(vFields) Then RaiseError(DataAccessErrors.daeRecordExists, "Contact Position Number / Relationship")
      ValidateValidPeriod(pParameterList)
    End Sub
    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      MyBase.PreValidateUpdateParameters(pParameterList)
      ValidateValidPeriod(pParameterList)
    End Sub

    Public Overrides Sub PreValidateParameterList(ByVal pType As CARERecord.MaintenanceTypes, ByVal pParameterList As CDBParameters)
      MyBase.PreValidateParameterList(pType, pParameterList)
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      SetHistoryOnlyFlag()
    End Sub

    Private Sub SetHistoryOnlyFlag()
      Dim vHistoryOnly As Boolean
      Dim vValidFrom As Date = DateSerial(1, 1, 1)
      Dim vValidTo As Date = DateSerial(9999, 12, 1)
      If IsDate(mvClassFields.Item(PositionLinkFields.ValidFrom).Value) Then vValidFrom = Date.Parse(mvClassFields.Item(PositionLinkFields.ValidFrom).Value)
      If IsDate(mvClassFields.Item(PositionLinkFields.ValidTo).Value) Then vValidTo = Date.Parse(mvClassFields.Item(PositionLinkFields.ValidTo).Value)
      If vValidFrom > Date.Parse(TodaysDate) Then
        'Record not started yet
        vHistoryOnly = True
      ElseIf vValidTo < Date.Parse(TodaysDate) Then
        'Record is now historic
        vHistoryOnly = True
      End If
      mvClassFields.Item(PositionLinkFields.HistoryOnly).Bool = vHistoryOnly
    End Sub

    Public Sub AmalgamateOrganisationPositionLink(ByVal pRecord As PositionLink, ByVal pNewContactPositionNumber As Integer, ByVal pNewValidFrom As String)
      CopyValues(pRecord)
      With mvClassFields
        .Item(PositionLinkFields.ContactPositionNumber).Value = pNewContactPositionNumber.ToString
        .Item(PositionLinkFields.ValidFrom).Value = pNewValidFrom
      End With
    End Sub

    Public Sub MovePositionLinks(ByVal pRecord As PositionLink, ByVal pNewContactPositionNumber As Integer, ByVal pNewPosStart As String, ByVal pNewPosFinish As String)
      'Copy old Position Links values to new Position Link record
      CopyValues(pRecord)
      With mvClassFields
        'Set new Contact Position Number
        .Item(PositionLinkFields.ContactPositionNumber).Value = pNewContactPositionNumber.ToString
        Dim vValidFrom As String = .Item(PositionLinkFields.ValidFrom).Value
        Dim vValidTo As String = .Item(PositionLinkFields.ValidTo).Value
        'If Valid From before new Position Start then set Valid From to be new Position Start
        If IsDate(pNewPosStart) AndAlso IsDate(vValidFrom) AndAlso Date.Compare(CDate(pNewPosStart), CDate(vValidFrom)) > 0 Then
          .Item(PositionLinkFields.ValidFrom).Value = pNewPosStart
          vValidFrom = .Item(PositionLinkFields.ValidFrom).Value
        End If
        'If Valid To after new Position Finish then set Valid To to be new Position Finish
        If IsDate(pNewPosFinish) AndAlso IsDate(vValidTo) AndAlso Date.Compare(CDate(vValidTo), CDate(pNewPosFinish)) > 0 Then
          mvClassFields.Item(PositionLinkFields.ValidTo).Value = pNewPosFinish
          vValidTo = mvClassFields.Item(PositionLinkFields.ValidTo).Value
        End If
        'Validate new activity dates
        If IsDate(vValidFrom) AndAlso IsDate(vValidTo) Then
          If Date.Compare(CDate(vValidFrom), CDate(vValidTo)) > 0 Then RaiseError(DataAccessErrors.daeInvalidDateRange)
        End If
      End With
    End Sub
    Private Sub ValidateValidPeriod(ByVal pParameterList As CDBParameters)

      Dim vValidFrom As String = ""
      Dim vValidTo As String = ""

      'Set defaults to existing 
      If pParameterList.ContainsKey("ValidFrom") Then
        vValidFrom = pParameterList("ValidFrom").Value
      ElseIf Me.Existing Then
        vValidFrom = ValidFrom
      End If
      If pParameterList.ContainsKey("ValidTo") Then
        vValidTo = pParameterList("ValidTo").Value
      ElseIf Me.Existing Then
        vValidTo = ValidTo
      End If

      'Get Contact Position 
      Dim vPosition As New ContactPosition(mvEnv)
      If pParameterList.ParameterExists("ContactPositionNumber").IntegerValue > 0 Then
        vPosition.Init(pParameterList("ContactPositionNumber").IntegerValue)
      ElseIf ContactPositionNumber > 0 Then
        vPosition.Init(ContactPositionNumber)
      End If

      'Check Dates against Position Started date 
      If vPosition.Started <> "" Then
        If vValidFrom > "" Then
          If CDate(vValidFrom) < CDate(vPosition.Started) Then
            RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
          End If
        Else
          RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
        End If
        If vValidTo > "" Then
          If CDate(vValidTo) < CDate(vPosition.Started) Then
            RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
          End If
        End If
      End If

      'Check Dates against Position Finished date 
      If vPosition.Finished <> "" Then
        If vValidTo > "" Then
          If CDate(vValidTo) > CDate(vPosition.Finished) Then
            RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
          End If
        Else
          RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
        End If
        If vValidFrom > "" Then
          If CDate(vValidFrom) > CDate(vPosition.Finished) Then
            RaiseError(DataAccessErrors.daePositionTimesheetDateInconsistentWithPosition)
          End If
        End If
      End If
    End Sub
#End Region

  End Class
End Namespace
